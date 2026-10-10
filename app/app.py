import os, sys, argparse
from datetime import datetime, date
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from merge import merge, build_xlsx
from daily import merge_daily, build_daily_xlsx, compute
import store

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'peplink-usage-tracker-local')
app.config['MAX_CONTENT_LENGTH'] = 256 * 1024 * 1024
MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


def _site_arg():
    s = (request.values.get('site') or '').strip().upper()
    all_sites = store.sites()
    return s or (all_sites[0] if all_sites else ''), all_sites


# ---------- ongoing tracking ----------
@app.route('/')
def home():
    site, all_sites = _site_arg()
    st = store.status(site) if site else None
    return render_template('import.html', site=site, sites=all_sites, st=st)


@app.route('/import', methods=['POST'])
def do_import():
    site = (request.form.get('site') or '').strip().upper()
    files = [f for f in request.files.getlist('files') if f and f.filename]
    if not files:
        flash('No files selected.')
        return redirect(url_for('home', site=site))
    msgs = []
    for f in files:
        fsite = site or store.site_from_name(f.filename)
        if not fsite:
            msgs.append(f'{f.filename}: skipped — no site in file name and none given')
            continue
        try:
            low = f.filename.lower()
            if low.endswith('.xlsx') or 'event' in low:
                r = store.import_events(fsite, f.filename, f.stream)
                msgs.append(f'{f.filename}: {r["rows"]} logins read, {r["added"]} new ({r["first"]} → {r["last"]})')
            else:
                r = store.import_usage(fsite, f.filename, f.stream)
                msgs.append(f'{f.filename}: {r["rows"]} devices stored for {r["day"]}')
            site = site or fsite
        except Exception as e:
            msgs.append(f'{f.filename}: ERROR {e}')
    for m in msgs:
        flash(m)
    return redirect(url_for('home', site=site))


@app.route('/report')
def report():
    site, all_sites = _site_arg()
    st = store.status(site) if site else None
    if not st or not st['days']:
        flash('No usage data imported yet for this site.')
        return redirect(url_for('home', site=site))
    start = request.args.get('start') or st['days'][max(0, len(st['days']) - 7)]
    end = request.args.get('end') or st['days'][-1]
    ev, byday = store.load_for_report(site, date.fromisoformat(start), date.fromisoformat(end))
    if not byday:
        flash('No usage data in that range.')
        return redirect(url_for('home', site=site))
    detail, pivot, days, summary = compute(ev, byday)
    if request.args.get('download'):
        name = f"{site}_Starlink_Usage_per_User_per_Day_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
        return send_file(build_daily_xlsx(detail, pivot, days), as_attachment=True, download_name=name, mimetype=MIME)
    return render_template('report.html', site=site, sites=all_sites, st=st, start=start, end=end,
                           detail=detail, pivot=pivot, days=days, summary=summary)


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
                if request.form.get('action') == 'download':
                    name = f"Starlink_Usage_per_User_per_Day_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
                    return send_file(build_daily_xlsx(detail, pivot, days), as_attachment=True, download_name=name, mimetype=MIME)
                return render_template('index.html', mode='daily', detail=detail, pivot=pivot, days=days, summary=summary)
            rows, per_user, unmatched, summary = merge(ev.stream, usage_files[0].stream)
        except Exception as e:
            flash(f'Could not process files: {e}')
            return redirect(url_for('quick'))
        if request.form.get('action') == 'download':
            name = f"Starlink_User_MAC_Usage_{datetime.now():%Y%m%d_%H%M}.xlsx"
            return send_file(build_xlsx(rows, per_user, unmatched), as_attachment=True, download_name=name, mimetype=MIME)
        return render_template('index.html', mode='total', rows=rows, per_user=per_user, unmatched=unmatched, summary=summary)
    return render_template('index.html')


@app.route('/health')
def health():
    return {'status': 'ok'}


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
        out = a.output or f"{a.site.upper()}_Starlink_Usage_per_User_per_Day_{days[0]:%Y%m%d}-{days[-1]:%Y%m%d}.xlsx"
        open(out, 'wb').write(build_daily_xlsx(detail, pivot, days).read())
        print(f"Wrote {out}: {summary['days']} days, {summary['users']} users, {summary['total_gb']} GB, unmatched {summary['unmatched_gb']} GB")


if __name__ == '__main__':
    if len(sys.argv) > 1:
        cli()
    else:
        app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5020)), debug=False)
