"""Per-day attribution: join daily usage exports (one CSV per day) with portal login events."""
import re, io, collections
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from merge import parse_usage, LOGIN_RE, MAC_RE, dev_label

DATE_RE = re.compile(r'(\d{4}-\d{2}-\d{2})')


def date_from_name(name):
    m = DATE_RE.search(name or '')
    return date.fromisoformat(m.group(1)) if m else None


def dates_from_name(name):
    """All YYYY-MM-DD dates in a file name, in order (a 7-day export carries a start and an end date)."""
    out = []
    for d in DATE_RE.findall(name or ''):
        try:
            out.append(date.fromisoformat(d))
        except ValueError:
            pass
    return out


def parse_events_by_day(fileobj):
    """Return {mac: [(datetime, user), ...]} sorted ascending (successful logins only)."""
    wb = openpyxl.load_workbook(fileobj, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip().lower() if h else '' for h in next(rows)]
    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None
    i_dt, i_mac, i_det = col('date/time', 'date', 'time'), col('client mac', 'mac'), col('details', 'event', 'message')
    if i_dt is None or i_mac is None or i_det is None:
        raise ValueError('Event log must have "Date/Time", "Client MAC" and "Details" columns')
    ev = collections.defaultdict(list)
    for r in rows:
        dt, mac, det = r[i_dt], r[i_mac], r[i_det]
        if not (dt and mac and det):
            continue
        m = LOGIN_RE.search(str(det))
        if not m:
            continue
        if isinstance(dt, str):
            try:
                dt = datetime.fromisoformat(dt)
            except ValueError:
                continue
        ev[str(mac).strip().upper().replace('-', ':')].append((dt, m.group(1)))
    for k in ev:
        ev[k].sort()
    return ev


def attribute(ev_list, day):
    """Pick user for a MAC on a given day.
    Returns (user, basis, users_that_day_str). basis: 'login' | 'carried' | 'inferred' | ''"""
    same = [u for dt, u in ev_list if dt.date() == day]
    if same:
        c = collections.Counter(same)
        return c.most_common(1)[0][0], 'login', ', '.join(f'{u} ({n})' for u, n in c.most_common())
    before = [(dt, u) for dt, u in ev_list if dt.date() < day]
    if before:
        return before[-1][1], 'carried', ''
    after = [(dt, u) for dt, u in ev_list if dt.date() > day]
    if after:  # event log starts after this day (export truncated) - use the next known login
        return after[0][1], 'inferred', ''
    return '', '', ''


def merge_daily(event_fileobj, usage_files):
    """usage_files: list of (filename, fileobj). Returns (detail_rows, pivot, days, summary)."""
    ev = parse_events_by_day(event_fileobj)
    byday = []
    for name, fobj in usage_files:
        d = date_from_name(name)
        if not d:
            raise ValueError(f'Cannot find a date (YYYY-MM-DD) in file name: {name}')
        byday.append((d, parse_usage(fobj)))
    return compute(ev, byday)


def compute(ev, usage_by_day):
    """ev: {mac: [(dt,user)...]} sorted; usage_by_day: [(date, {mac: usage dict})]."""
    days = []
    detail = []
    for d, usage in usage_by_day:
        days.append(d)
        for mac, u in usage.items():
            user, basis, seen = attribute(ev.get(mac, []), d) if u['is_mac'] else ('', '', '')
            detail.append({'day': d, 'mac': mac, 'name': u['name'], 'user': user, 'basis': basis, 'seen': seen,
                           'up': round(u['up'], 2), 'down': round(u['down'], 2), 'tot': round(u['tot'], 2),
                           'is_mac': u['is_mac']})
    days = sorted(set(days))
    detail.sort(key=lambda r: (r['day'], -r['tot']))
    pivot = collections.defaultdict(lambda: {d: 0.0 for d in days})
    macs = collections.defaultdict(set)
    names = {}
    for r in detail:
        if r['user']:
            pivot[r['user']][r['day']] += r['tot']
            macs[r['user']].add(r['mac'])
            if r['name']:
                names[r['mac']] = r['name']
    pivot_rows = []
    for u, per in pivot.items():
        total = sum(per.values())
        pivot_rows.append({'user': u, 'per': per, 'total': round(total, 2), 'gb': round(total / 1024, 2),
                           'macs': sorted(macs[u]), 'devs': [dev_label(m, names.get(m, '')) for m in sorted(macs[u])], 'days_active': sum(1 for v in per.values() if v > 0)})
    pivot_rows.sort(key=lambda r: -r['total'])
    unmatched_tot = sum(r['tot'] for r in detail if not r['user'])
    summary = {'days': len(days), 'first': days[0] if days else None, 'last': days[-1] if days else None,
               'users': len(pivot_rows), 'device_days': len(detail),
               'matched': sum(1 for r in detail if r['user']),
               'carried': sum(1 for r in detail if r['basis'] == 'carried'),
               'inferred': sum(1 for r in detail if r['basis'] == 'inferred'),
               'unmatched_gb': round(unmatched_tot / 1024, 2),
               'total_gb': round(sum(r['tot'] for r in detail) / 1024, 2)}
    return detail, pivot_rows, days, summary


def compute_period(ev, usage, start, end):
    """One usage file covering start..end (e.g. 7 days): one row per device, user = most logins inside the period;
    else last earlier login (carried); else next later login (inferred). Same row shape as merge.merge()."""
    rows = []
    logged = set()
    for mac, u in usage.items():
        evl = ev.get(mac, []) if u['is_mac'] else []
        inside = [(dt, usr) for dt, usr in evl if start <= dt.date() <= end]
        user = users = lt = lu = ''
        n = 0
        if inside:
            c = collections.Counter(usr for _, usr in inside)
            user = c.most_common(1)[0][0]
            users = ', '.join(f'{x} ({k})' for x, k in c.most_common())
            lt, lu, n = inside[-1][0].strftime('%Y-%m-%d %H:%M'), inside[-1][1], len(inside)
        else:
            before = [x for x in evl if x[0].date() < start]
            after = [x for x in evl if x[0].date() > end]
            if before:
                user, users = before[-1][1], '(carried from earlier login)'
            elif after:
                user, users = after[0][1], '(inferred from next login)'
        if evl:
            logged.add(mac)
        rows.append({'user': user, 'mac': mac, 'name': u['name'], 'up': round(u['up'], 2), 'down': round(u['down'], 2),
                     'tot': round(u['tot'], 2), 'gb': round(u['tot'] / 1024, 2), 'logins': n, 'users': users,
                     'last_user': lu, 'last_time': lt, 'pct': round(u.get('pct', 0.0), 2), 'is_mac': u['is_mac']})
    rows.sort(key=lambda r: -r['tot'])
    per_user = collections.OrderedDict()
    for r in rows:
        if r['user']:
            p = per_user.setdefault(r['user'], {'macs': [], 'devs': [], 'tot': 0.0})
            p['macs'].append(r['mac'])
            p['devs'].append(dev_label(r['mac'], r['name']))
            p['tot'] += r['tot']
    per_user = collections.OrderedDict(sorted(per_user.items(), key=lambda x: -x[1]['tot']))
    unmatched = [r for r in rows if not r['user']]
    summary = {'non_mac': sum(1 for r in rows if not r['is_mac']), 'devices': len(rows),
               'matched': len(rows) - len(unmatched), 'unmatched': len(unmatched), 'users': len(per_user),
               'login_macs': len(logged), 'total_gb': round(sum(r['tot'] for r in rows) / 1024, 2)}
    return rows, per_user, unmatched, summary


def _style(sh):
    for c in sh[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F4E78')
    sh.freeze_panes = 'B2'
    sh.auto_filter.ref = sh.dimensions
    for col in sh.columns:
        w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        sh.column_dimensions[get_column_letter(col[0].column)].width = min(max(10, w + 2), 50)
    for row in sh.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, float):
                c.number_format = '#,##0.00'


def build_daily_xlsx(detail, pivot_rows, days):
    out = openpyxl.Workbook()
    s = out.active
    s.title = 'Usage per user per day'
    s.append(['User ID'] + [d.strftime('%d %b') for d in days] + ['Total (MB)', 'Total (GB)', 'Days active', 'No. of devices', 'MACs (hostname)'])
    for r in pivot_rows:
        s.append([r['user']] + [round(r['per'][d], 2) for d in days] + [r['total'], r['gb'], r['days_active'], len(r['macs']), ', '.join(r['devs'])])
    # daily totals row
    s.append(['TOTAL'] + [round(sum(r['per'][d] for r in pivot_rows), 2) for d in days] + [round(sum(r['total'] for r in pivot_rows), 2), '', '', '', ''])
    for c in s[s.max_row]:
        c.font = Font(bold=True)
    s2 = out.create_sheet('Device-day detail')
    s2.append(['Date', 'User ID', 'MAC address', 'Hostname', 'Upload (MB)', 'Download (MB)', 'Total (MB)', 'Basis', 'Users logged in that day (logins)'])
    for r in detail:
        s2.append([r['day'].isoformat(), r['user'], r['mac'], r['name'], r['up'], r['down'], r['tot'],
                   {'login': 'Login same day', 'carried': 'Carried from earlier login', 'inferred': 'Inferred from next login (log starts later)', '': ('No login found' if r['is_mac'] else 'Not a MAC')}[r['basis']],
                   r['seen']])
    s3 = out.create_sheet('Unmatched')
    s3.append(['Date', 'MAC / ID', 'Hostname', 'Total (MB)', 'Note'])
    for r in detail:
        if not r['user']:
            s3.append([r['day'].isoformat(), r['mac'], r['name'], r['tot'], 'No login found' if r['is_mac'] else 'Not a MAC (IP / non-portal traffic)'])
    for sh in out:
        _style(sh)
    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)
    return buf
