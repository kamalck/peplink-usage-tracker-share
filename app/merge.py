"""Core logic: join Starlink portal login events (xlsx or csv) with per-device usage (csv)."""
import csv, io, re, collections
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

MAC_RE = re.compile(r'^([0-9A-F]{2}:){5}[0-9A-F]{2}$')
LOGIN_RE = re.compile(r'guest account ID "([^"]+)" \[login\]')


def dev_label(mac, name):
    """'MAC (hostname)' - the hostname is the Name column of the usage export."""
    return f'{mac} ({name})' if name else mac


def _to_bytes(fileobj):
    data = fileobj.read() if hasattr(fileobj, 'read') else fileobj
    return data.encode('utf-8') if isinstance(data, str) else data


def _decode(data):
    """CSV bytes -> text: UTF-8 (with or without BOM), else Windows-1252."""
    try:
        return data.decode('utf-8-sig')
    except UnicodeDecodeError:
        return data.decode('cp1252')


def event_table(fileobj):
    """Open an event log in either format and return (header, rows): an Excel workbook (detected by its zip signature, whatever the
    file is called) or a CSV (comma, semicolon or tab separated). header = lower-cased column names; rows = lists of cell values."""
    data = _to_bytes(fileobj)
    if data[:2] == b'PK':
        ws = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True).active
        rows = ws.iter_rows(values_only=True)
        try:
            first = next(rows)
        except StopIteration:
            raise ValueError('The event log is empty') from None
        return [str(h).strip().lower() if h else '' for h in first], rows
    text = _decode(data)
    head = text.split('\n', 1)[0]
    delim = max(',;\t', key=head.count) if any(d in head for d in ',;\t') else ','
    rdr = csv.reader(io.StringIO(text, newline=''), delimiter=delim)
    try:
        first = next(rdr)
    except StopIteration:
        raise ValueError('The event log is empty') from None
    return [h.strip().lower() for h in first], rdr


def csv_kind(data):
    """What a CSV is, from its header row: 'events' (portal event log), 'usage' (per-device usage), 'both', or 'unknown'."""
    try:
        header, _ = event_table(data)
    except Exception:
        return 'unknown'
    ev = any(h in header for h in ('client mac', 'mac')) and any(h in header for h in ('details', 'event', 'message'))
    us = any(h in header for h in ('mac address', 'mac')) and any(h in header for h in ('total (mb)', 'total'))
    return 'both' if ev and us else 'events' if ev else 'usage' if us else 'unknown'


def parse_event_log(fileobj):
    """Return (logins{mac: Counter(user)}, last{mac: (dt, user)}) from an event log (.xlsx or .csv)."""
    from daily import parse_events_stats          # same date handling as the stored import (text dates, swapped dates)
    ev, _ = parse_events_stats(fileobj)
    logins = collections.defaultdict(collections.Counter)
    last = {}
    for mac, lst in ev.items():
        for dt, user in lst:
            logins[mac][user] += 1
        last[mac] = lst[-1]
    return logins, last


def _num(value, col, line):
    """A usage figure: tolerates thousands separators ("1,234.5"); a real non-number gets a readable error."""
    t = (value or '').replace(',', '').replace('%', '').strip()
    if not t:
        return 0.0
    try:
        return float(t)
    except ValueError:
        raise ValueError(f'row {line}: "{value}" is not a number in column "{col}"') from None


def parse_usage(fileobj):
    """Return OrderedDict{mac: {name, up, down, tot}} (duplicate MACs summed)."""
    text = fileobj.read()
    if isinstance(text, bytes):
        text = text.decode('utf-8-sig', errors='replace')
    rdr = csv.DictReader(io.StringIO(text))
    fields = {f.strip().lower(): f for f in rdr.fieldnames or []}
    def f(*names):
        for n in names:
            if n in fields:
                return fields[n]
        raise ValueError(f'Usage CSV missing column: {names[0]}')
    c_name = next((fields[n] for n in ('name', 'hostname', 'host name') if n in fields), None)   # optional: blank when absent
    c_mac = f('mac address', 'mac')
    c_up, c_down, c_tot = f('upload (mb)', 'upload'), f('download (mb)', 'download'), f('total (mb)', 'total')
    c_pct = fields.get('% usage')
    usage = collections.OrderedDict()
    for row in rdr:
        mac = (row.get(c_mac) or '').strip().upper().replace('-', ':')
        if not mac:
            continue
        d = usage.setdefault(mac, {'name': ((row.get(c_name) if c_name else '') or '').strip(), 'up': 0.0, 'down': 0.0, 'tot': 0.0, 'pct': 0.0,
                                   'is_mac': bool(MAC_RE.match(mac))})
        d['up'] += _num(row.get(c_up), c_up, rdr.line_num)
        d['down'] += _num(row.get(c_down), c_down, rdr.line_num)
        d['tot'] += _num(row.get(c_tot), c_tot, rdr.line_num)
        if c_pct:
            d['pct'] += _num(row.get(c_pct), c_pct, rdr.line_num)
    return usage


def merge(event_fileobj, usage_fileobj):
    logins, last = parse_event_log(event_fileobj)
    usage = parse_usage(usage_fileobj)
    rows = []
    for mac, d in usage.items():
        c = logins.get(mac)
        if c:
            primary = c.most_common(1)[0][0]
            users = ', '.join(f'{u} ({n})' for u, n in c.most_common())
            lt, lu = (last[mac][0].strftime('%Y-%m-%d %H:%M'), last[mac][1]) if mac in last else ('', '')
            n = sum(c.values())
        else:
            primary = users = lt = lu = ''
            n = 0
        rows.append({
            'user': primary, 'mac': mac, 'name': d['name'],
            'up': round(d['up'], 2), 'down': round(d['down'], 2),
            'tot': round(d['tot'], 2), 'gb': round(d['tot'] / 1024, 2),
            'logins': n, 'users': users, 'last_user': lu, 'last_time': lt,
            'pct': round(d['pct'], 2), 'is_mac': d['is_mac'],
        })
    rows.sort(key=lambda r: -r['tot'])
    per_user = collections.OrderedDict()
    for r in rows:
        if r['user']:
            p = per_user.setdefault(r['user'], {'macs': [], 'devs': [], 'up': 0.0, 'down': 0.0, 'tot': 0.0})
            p['macs'].append(r['mac'])
            p['devs'].append(dev_label(r['mac'], r['name']))
            p['up'] += r['up']
            p['down'] += r['down']
            p['tot'] += r['tot']
    per_user = collections.OrderedDict(sorted(per_user.items(), key=lambda x: -x[1]['tot']))
    unmatched = [r for r in rows if not r['user']]
    non_mac = sum(1 for r in rows if not r['is_mac'])
    summary = {
        'non_mac': non_mac,
        'devices': len(rows), 'matched': len(rows) - len(unmatched), 'unmatched': len(unmatched),
        'users': len(per_user), 'login_macs': len(logins),
        'total_gb': round(sum(r['tot'] for r in rows) / 1024, 2),
        'up_gb': round(sum(r['up'] for r in rows) / 1024, 2), 'down_gb': round(sum(r['down'] for r in rows) / 1024, 2),
        'unmatched_gb': round(sum(r['tot'] for r in unmatched) / 1024, 2),
    }
    return rows, per_user, unmatched, summary


def build_xlsx(rows, per_user, unmatched):
    out = openpyxl.Workbook()
    s = out.active
    s.title = 'Device usage by user'
    s.append(['User ID (primary)', 'MAC address', 'Hostname', 'Upload (MB)', 'Download (MB)',
              'Total (MB)', 'Total (GB)', '% Usage', 'Logins', 'All user IDs seen (logins)', 'Last login user', 'Last login time'])
    for r in rows:
        s.append([r['user'], r['mac'], r['name'], r['up'], r['down'], r['tot'], r['gb'], r['pct'],
                  r['logins'], r['users'], r['last_user'], r['last_time']])
    s2 = out.create_sheet('Usage per user')
    s2.append(['User ID', 'No. of devices', 'Upload (MB)', 'Download (MB)', 'Total (MB)', 'Total (GB)', 'Devices (MAC (hostname))'])
    for u, p in per_user.items():
        s2.append([u, len(p['macs']), round(p['up'], 2), round(p['down'], 2), round(p['tot'], 2), round(p['tot'] / 1024, 2), ', '.join(p['devs'])])
    s3 = out.create_sheet('Unmatched devices')
    s3.append(['MAC address', 'Hostname', 'Upload (MB)', 'Download (MB)', 'Total (MB)', 'Upload (GB)', 'Download (GB)', 'Total (GB)', 'Note'])
    for r in unmatched:
        s3.append([r['mac'], r['name'], r['up'], r['down'], r['tot'], round(r['up'] / 1024, 2), round(r['down'] / 1024, 2), round(r['tot'] / 1024, 2), 'No login found' if r['is_mac'] else 'Not a MAC (IP / non-portal traffic)'])
    for sh in out:
        for c in sh[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F4E78')
        sh.freeze_panes = 'A2'
        sh.auto_filter.ref = sh.dimensions
        for col in sh.columns:
            w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            sh.column_dimensions[get_column_letter(col[0].column)].width = min(max(10, w + 2), 60)
        for row in sh.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = '#,##0.00'
    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)
    return buf
