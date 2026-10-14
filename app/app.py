import os, sys, argparse, hmac, secrets, time
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
from merge import merge, build_xlsx
from daily import merge_daily, build_daily_xlsx, compute, compute_period, device_pivot, build_period_xlsx, build_device_xlsx, build_period_device_xlsx, build_total_xlsx, unmatched_devices
import store

VERSION = '1.7.2'


def _secret_key():
    """SECRET_KEY from .env; if it is empty, make one and keep it in data/ (git-ignored) so logins survive a restart."""
    k = os.environ.get('SECRET_KEY', '').strip()
    if k:
        return k
    path = os.path.join(os.path.dirname(store.DB_PATH), '.secret_key')
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if os.path.exists(path):
            k = open(path, encoding='ascii').read().strip()
            if k:
                return k
        k = secrets.token_hex(32)
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, 'w', encoding='ascii') as f:
            f.write(k)
        return k
    except OSError:
        return secrets.token_hex(32)


app = Flask(__name__)
app.secret_key = _secret_key()
app.config['MAX_CONTENT_LENGTH'] = 256 * 1024 * 1024
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax',
                  PERMANENT_SESSION_LIFETIME=timedelta(hours=12))
MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

# Login comes from the environment (.env on the host) - nothing secret is ever in the repo.
ADMIN_USERNAME = os.environ.get('ADMIN_USERNAME', 'admin').strip() or 'admin'
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', '')


def _accounts():
    """{user: password}: ADMIN_USERNAME/ADMIN_PASSWORD plus any 'user:password;user2:password2' in ADMIN_USERS."""
    acc = {ADMIN_USERNAME: ADMIN_PASSWORD} if ADMIN_PASSWORD else {}
    for pair in os.environ.get('ADMIN_USERS', '').split(';'):
        u, sep, pw = pair.partition(':')
        u = u.strip()
        if sep and u and pw:
            acc.setdefault(u, pw)
    return acc


ACCOUNTS = _accounts()
OPEN_ENDPOINTS = {'health', 'login', 'static'}
_fails = {}          # remote addr -> [failures, window start, blocked until]
MAX_FAILS, FAIL_WINDOW, BLOCK_SECONDS = 5, 600, 300


def _same(a, b):
    return hmac.compare_digest((a or '').encode('utf-8'), (b or '').encode('utf-8'))


def _csrf_token():
    if 'csrf' not in session:
        session['csrf'] = secrets.token_hex(16)
    return session['csrf']


app.jinja_env.globals['csrf_token'] = _csrf_token
app.jinja_env.globals['app_version'] = VERSION


@app.before_request
def gate():
    """Everything except /health, /login and static files needs the admin login. No password set = locked, not open."""
    if request.endpoint in OPEN_ENDPOINTS:
        return None
    if not ACCOUNTS:
        return render_template('locked.html'), 503
    if not session.get('auth') or session.get('user') not in ACCOUNTS:
        nxt = request.full_path.rstrip('?') if request.method == 'GET' else ''
        return redirect(url_for('login', next=nxt))
    if request.method == 'POST' and not _same(request.form.get('_csrf'), session.get('csrf')):
        flash('ERROR: session expired - please try again.')
        return redirect(url_for('home'))
    return None


def _safe_next(n):
    ok = n and n.startswith('/') and not n.startswith('//') and '\\' not in n
    return n if ok else url_for('home')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if not ACCOUNTS:
        return render_template('locked.html'), 503
    nxt = request.values.get('next', '')
    if request.method == 'POST':
        ip = request.remote_addr or '?'
        now = time.time()
        rec = _fails.get(ip)
        if rec and rec[2] > now:
            flash('ERROR: too many attempts - try again in %d min.' % (int(rec[2] - now) // 60 + 1))
            return render_template('login.html', next=nxt), 429
        if not _same(request.form.get('_csrf'), session.get('csrf')):
            flash('ERROR: session expired - please try again.')
            return render_template('login.html', next=nxt)
        who = None
        for name, pw in ACCOUNTS.items():          # no early exit: same work whichever account matches
            if _same(request.form.get('username', '').strip(), name) & _same(request.form.get('password', ''), pw):
                who = name
        if who:
            _fails.pop(ip, None)
            session.clear()
            session['auth'] = True
            session['user'] = who
            session.permanent = True
            return redirect(_safe_next(nxt))
        rec = _fails.get(ip)
        if not rec or now - rec[1] > FAIL_WINDOW:
            rec = [0, now, 0]
        rec[0] += 1
        if rec[0] >= MAX_FAILS:
            rec[2] = now + BLOCK_SECONDS
            rec[0], rec[1] = 0, now
        _fails[ip] = rec
        flash('ERROR: wrong user name or password.')
    return render_template('login.html', next=nxt)


@app.route('/logout', methods=['POST'])
def logout():
    session.clear()
    return redirect(url_for('login'))


def _day_arg(name, default):
    """A YYYY-MM-DD query value, or the default when missing/garbled (a hand-edited URL must not 500)."""
    v = request.args.get(name)
    if v:
        try:
            return date.fromisoformat(v).isoformat()
        except ValueError:
            flash('ERROR: "%s" is not a date (use YYYY-MM-DD) - showing the default range.' % v)
    return default


def _view_arg(src):
    """'total' (whole period, the default) or 'day' (one column per day)."""
    return 'day' if (src.get('view') or '').lower() == 'day' else 'total'


def _daily_files(site, detail, pivot, days, view):
    """(download name, workbook) for the Excel buttons of the day/total report."""
    rng = f"{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}"
    if view == 'total':
        return f"{site}_Starlink_Usage_per_User_{rng}.xlsx", build_total_xlsx(detail, pivot, days)
    return f"{site}_Starlink_Usage_per_User_per_Day_{rng}.xlsx", build_daily_xlsx(detail, pivot, days)


def _site_arg():
    s = (request.values.get('site') or '').strip().upper()
    all_sites = store.sites()
    return s or (all_sites[0] if all_sites else ''), all_sites


# ---------- ongoing tracking ----------
@app.route('/')
def home():
    site, all_sites = _site_arg()
    st = store.status(site) if site else None
    return render_template('import.html', site=site, sites=all_sites, st=st, overview=store.sites_overview())


@app.route('/site_panel')
def site_panel():
    """HTML fragment with one site's import status - the Import page swaps it in when a site is picked."""
    site = (request.args.get('site') or '').strip().upper()
    return render_template('_site_status.html', site=site, st=store.status(site) if site else None)


@app.route('/import', methods=['POST'])
def do_import():
    site = (request.form.get('site') or '').strip().upper()
    files = [f for f in request.files.getlist('files') if f and f.filename]
    if not files:
        flash('No files selected.')
        return redirect(url_for('home', site=site))
    kind = request.form.get('kind', 'daily')      # daily | period (a 7-day export)
    p_start = None
    if request.form.get('period_start'):
        try:
            p_start = date.fromisoformat(request.form['period_start'])
        except ValueError:
            flash('ERROR: period start is not a date - ignored.')
    msgs = []
    for f in files:
        fsite = site or store.site_from_name(f.filename)
        if not fsite:
            msgs.append(f'{f.filename}: skipped — no site in file name and none given')
            continue
        try:
            low = f.filename.lower()
            if low.endswith('.xlsx'):
                r = store.import_events(fsite, f.filename, f.stream)
                msgs.append(f'{f.filename}: {r["rows"]} logins read, {r["added"]} new ({r["first"]} → {r["last"]})')
            elif low.endswith('.csv'):
                period = kind == 'period'
                if period:
                    r = store.import_period(fsite, f.filename, f.stream, start=p_start)
                    msgs.append(f'{f.filename}: {r["rows"]} devices stored for period {r["start"]} → {r["end"]}')
                else:
                    r = store.import_usage(fsite, f.filename, f.stream)
                    msgs.append(f'{f.filename}: {r["rows"]} devices stored for {r["day"]}')
            else:
                msgs.append(f'{f.filename}: ERROR unsupported file type (event log = .xlsx, usage = .csv)')
                continue
            site = site or fsite
        except Exception as e:
            msgs.append(f'{f.filename}: ERROR {e}')
    for m in msgs:
        flash(m)
    return redirect(url_for('home', site=site))


@app.route('/delete', methods=['POST'])
def do_delete():
    site = (request.form.get('site') or '').strip().upper()
    what = request.form.get('what', '')
    if not site or what not in ('usage', 'periods', 'logins', 'all'):
        flash('ERROR: choose a site and what to delete.')
        return redirect(url_for('home', site=site))
    rng = []
    for k in ('start', 'end'):
        v = request.form.get(k, '').strip()
        try:
            rng.append(date.fromisoformat(v) if v else None)
        except ValueError:
            flash(f'ERROR: "{v}" is not a date (use YYYY-MM-DD) - nothing deleted.')
            return redirect(url_for('home', site=site))
    if what == 'all' and request.form.get('confirm', '').strip().upper() != site:
        flash(f'ERROR: to delete everything for {site}, type the site name in the confirm box - nothing deleted.')
        return redirect(url_for('home', site=site))
    n = store.delete_data(site, what, rng[0], rng[1])
    label = {'usage': 'daily usage rows', 'periods': 'period usage rows', 'logins': 'logins', 'all': 'rows (everything)'}[what]
    flash(f'Deleted {n} {label} for {site}.')
    return redirect(url_for('home', site=site))


@app.route('/period')
def period():
    site, all_sites = _site_arg()
    try:
        start, end = date.fromisoformat(request.args.get('start', '')), date.fromisoformat(request.args.get('end', ''))
    except ValueError:
        flash('ERROR: bad period dates.')
        return redirect(url_for('home', site=site))
    ev, usage = store.load_period(site, start, end)
    if not usage:
        flash('No such period stored for this site.')
        return redirect(url_for('home', site=site))
    rows, per_user, unmatched, summary = compute_period(ev, usage, start, end)
    if request.args.get('download') == 'device':
        name = f"{site}_Starlink_Usage_per_Device_{start:%Y%m%d}-{end:%Y%m%d}.xlsx"
        return send_file(build_period_device_xlsx(rows), as_attachment=True, download_name=name, mimetype=MIME)
    if request.args.get('download'):
        name = f"{site}_Starlink_Usage_per_User_{start:%Y%m%d}-{end:%Y%m%d}.xlsx"
        return send_file(build_period_xlsx(rows, per_user, unmatched), as_attachment=True, download_name=name, mimetype=MIME)
    return render_template('period.html', site=site, start=start, end=end, rows=rows, per_user=per_user,
                           unmatched=unmatched, summary=summary)


@app.route('/report')
def report():
    site, all_sites = _site_arg()
    st = store.status(site) if site else None
    if not st or not st['days']:
        flash('No usage data imported yet for this site.')
        return redirect(url_for('home', site=site))
    start = _day_arg('start', st['days'][max(0, len(st['days']) - 7)])
    end = _day_arg('end', st['days'][-1])
    ev, byday = store.load_for_report(site, date.fromisoformat(start), date.fromisoformat(end))
    if not byday:
        flash('No usage data in that range.')
        return redirect(url_for('home', site=site))
    detail, pivot, days, summary = compute(ev, byday)
    view = _view_arg(request.args)
    prefix = f"{site}_" if site else ''
    if request.args.get('download') == 'device':
        name = f"{prefix}Starlink_Usage_per_Device{'' if view == 'total' else '_per_Day'}_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
        return send_file(build_device_xlsx(detail, days, view), as_attachment=True, download_name=name, mimetype=MIME)
    if request.args.get('download'):
        name, buf = _daily_files(site, detail, pivot, days, view)
        return send_file(buf, as_attachment=True, download_name=name, mimetype=MIME)
    return render_template('report.html', site=site, sites=all_sites, st=st, start=start, end=end, view=view,
                           detail=detail, pivot=pivot, days=days, summary=summary, devpivot=device_pivot(detail, days),
                           unmatched=unmatched_devices(detail))


# ---------- one-off merge (no storage) ----------
@app.route('/quick', methods=['GET', 'POST'])
def quick():
    if request.method == 'POST':
        ev = request.files.get('events')
        usage_files = [f for f in request.files.getlist('usage') if f and f.filename]
        if not ev or not ev.filename or not usage_files:
            flash('Please select the event log and at least one usage file.')
            return redirect(url_for('quick'))
        try:
            if len(usage_files) > 1:
                detail, pivot, days, summary = merge_daily(ev.stream, [(f.filename, f.stream) for f in usage_files])
                view = _view_arg(request.form)
                if request.form.get('action') == 'download_device':
                    name = f"Starlink_Usage_per_Device{'' if view == 'total' else '_per_Day'}_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
                    return send_file(build_device_xlsx(detail, days, view), as_attachment=True, download_name=name, mimetype=MIME)
                if request.form.get('action') == 'download':
                    name, buf = _daily_files('', detail, pivot, days, view)
                    return send_file(buf, as_attachment=True, download_name=name.lstrip('_'), mimetype=MIME)
                return render_template('index.html', mode='daily', detail=detail, pivot=pivot, days=days, summary=summary,
                                       devpivot=device_pivot(detail, days), unmatched=unmatched_devices(detail), view=view)
            rows, per_user, unmatched, summary = merge(ev.stream, usage_files[0].stream)
        except Exception as e:
            flash(f'Could not process files: {e}')
            return redirect(url_for('quick'))
        if request.form.get('action') == 'download_device':
            name = f"Starlink_Usage_per_Device_{datetime.now():%Y%m%d_%H%M}.xlsx"
            return send_file(build_period_device_xlsx(rows), as_attachment=True, download_name=name, mimetype=MIME)
        if request.form.get('action') == 'download':
            name = f"Starlink_User_MAC_Usage_{datetime.now():%Y%m%d_%H%M}.xlsx"
            return send_file(build_xlsx(rows, per_user, unmatched), as_attachment=True, download_name=name, mimetype=MIME)
        return render_template('index.html', mode='total', rows=rows, per_user=per_user, unmatched=unmatched, summary=summary)
    return render_template('index.html')


@app.route('/health')
def health():
    return {'status': 'ok', 'version': VERSION}


# ---------- CLI ----------
def cli():
    p = argparse.ArgumentParser(prog='app.py', description='Starlink usage by user')
    sub = p.add_subparsers(dest='cmd', required=True)
    m = sub.add_parser('merge', help='One-off merge, no storage')
    m.add_argument('events'); m.add_argument('usage', nargs='+'); m.add_argument('-o', '--output')
    i = sub.add_parser('import', help='Import event logs (.xlsx) and daily usage (.csv) into the store')
    i.add_argument('files', nargs='+'); i.add_argument('--site', help='Site label (default: parsed from file name, e.g. SITE1)')
    r = sub.add_parser('report', help='Per-user-per-day report from the store')
    r.add_argument('--site', required=True); r.add_argument('--start'); r.add_argument('--end'); r.add_argument('-o', '--output')
    r.add_argument('--view', choices=['total', 'day'], default='total', help='total = one row per user/device for the whole range (default); day = one column per day')
    a = p.parse_args()
    if a.cmd == 'merge':
        if len(a.usage) > 1:
            with open(a.events, 'rb') as ef:
                detail, pivot, days, summary = merge_daily(ef, [(os.path.basename(u), open(u, 'rb')) for u in a.usage])
            out = a.output or f"Starlink_Usage_per_User_per_Day_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
            open(out, 'wb').write(build_daily_xlsx(detail, pivot, days).read())
            print(f"Wrote {out}: {summary['days']} days, {summary['users']} users, {summary['matched']}/{summary['device_days']} device-days matched, {summary['total_gb']} GB")
        else:
            with open(a.events, 'rb') as ef, open(a.usage[0], 'rb') as uf:
                rows, per_user, unmatched, summary = merge(ef, uf)
            out = a.output or f"Starlink_User_MAC_Usage_{datetime.now():%Y%m%d_%H%M}.xlsx"
            open(out, 'wb').write(build_xlsx(rows, per_user, unmatched).read())
            print(f"Wrote {out}: {summary['matched']}/{summary['devices']} devices matched, {summary['users']} users, {summary['total_gb']} GB")
    elif a.cmd == 'import':
        for f in a.files:
            name = os.path.basename(f)
            site = (a.site or store.site_from_name(name)).upper()
            if not site:
                print(f'{name}: skipped (no site)'); continue
            with open(f, 'rb') as fh:
                if name.lower().endswith('.xlsx'):
                    r = store.import_events(site, name, fh); print(f'{name}: {r["rows"]} logins, {r["added"]} new ({r["first"]} → {r["last"]})')
                else:
                    r = store.import_usage(site, name, fh); print(f'{name}: {r["rows"]} devices for {r["day"]}')
    elif a.cmd == 'report':
        st = store.status(a.site.upper())
        if not st['days']:
            sys.exit('No data for site')
        start, end = a.start or st['days'][0], a.end or st['days'][-1]
        ev, byday = store.load_for_report(a.site.upper(), date.fromisoformat(start), date.fromisoformat(end))
        detail, pivot, days, summary = compute(ev, byday)
        dflt, buf = _daily_files(a.site.upper(), detail, pivot, days, a.view)
        out = a.output or dflt
        open(out, 'wb').write(buf.read())
        print(f"Wrote {out}: {summary['days']} days, {summary['users']} users, {summary['total_gb']} GB, unmatched {summary['unmatched_gb']} GB")


if __name__ == '__main__':
    if len(sys.argv) > 1:
        cli()
    else:
        app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5020)), debug=False)
