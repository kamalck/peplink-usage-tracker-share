"""SQLite store for ongoing Starlink usage tracking (multiple sites, overlapping log exports)."""
import os, re, sqlite3, collections
from datetime import datetime, date
from merge import parse_usage
from daily import parse_events_by_day, date_from_name

DB_PATH = os.environ.get('DB_PATH', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'peplink.db'))
SITE_RE = re.compile(r'([A-Za-z0-9]+)_-_Starlink', re.I)

SCHEMA = """
CREATE TABLE IF NOT EXISTS logins (
  site TEXT NOT NULL, ts TEXT NOT NULL, mac TEXT NOT NULL, user TEXT NOT NULL,
  PRIMARY KEY (site, ts, mac, user));
CREATE INDEX IF NOT EXISTS ix_logins_site_mac ON logins(site, mac, ts);
CREATE TABLE IF NOT EXISTS usage_daily (
  site TEXT NOT NULL, day TEXT NOT NULL, mac TEXT NOT NULL, name TEXT, up REAL, down REAL, tot REAL, is_mac INTEGER,
  PRIMARY KEY (site, day, mac));
CREATE TABLE IF NOT EXISTS imports (
  id INTEGER PRIMARY KEY AUTOINCREMENT, imported_at TEXT, site TEXT, kind TEXT, filename TEXT,
  rows INTEGER, added INTEGER, first_ts TEXT, last_ts TEXT);
"""


def site_from_name(name):
    m = SITE_RE.search(name or '')
    return m.group(1).upper() if m else ''


def connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.executescript(SCHEMA)
    return con


def import_events(site, filename, fileobj):
    ev = parse_events_by_day(fileobj)
    rows = [(site, dt.isoformat(sep=' '), mac, user) for mac, lst in ev.items() for dt, user in lst]
    con = connect()
    before = con.execute('SELECT COUNT(*) FROM logins WHERE site=?', (site,)).fetchone()[0]
    con.executemany('INSERT OR IGNORE INTO logins VALUES (?,?,?,?)', rows)
    after = con.execute('SELECT COUNT(*) FROM logins WHERE site=?', (site,)).fetchone()[0]
    ts = sorted(r[1] for r in rows)
    con.execute('INSERT INTO imports (imported_at, site, kind, filename, rows, added, first_ts, last_ts) VALUES (?,?,?,?,?,?,?,?)',
                (datetime.now().isoformat(sep=' ', timespec='seconds'), site, 'events', filename, len(rows), after - before,
                 ts[0] if ts else None, ts[-1] if ts else None))
    con.commit(); con.close()
    return {'rows': len(rows), 'added': after - before, 'first': ts[0] if ts else '', 'last': ts[-1] if ts else ''}


def import_usage(site, filename, fileobj, day=None):
    day = day or date_from_name(filename)
    if not day:
        raise ValueError(f'No date (YYYY-MM-DD) in file name: {filename}')
    usage = parse_usage(fileobj)
    con = connect()
    con.execute('DELETE FROM usage_daily WHERE site=? AND day=?', (site, day.isoformat()))
    con.executemany('INSERT INTO usage_daily VALUES (?,?,?,?,?,?,?,?)',
                    [(site, day.isoformat(), mac, u['name'], u['up'], u['down'], u['tot'], int(u['is_mac'])) for mac, u in usage.items()])
    con.execute('INSERT INTO imports (imported_at, site, kind, filename, rows, added, first_ts, last_ts) VALUES (?,?,?,?,?,?,?,?)',
                (datetime.now().isoformat(sep=' ', timespec='seconds'), site, 'usage', filename, len(usage), len(usage),
                 day.isoformat(), day.isoformat()))
    con.commit(); con.close()
    return {'rows': len(usage), 'day': day}


def sites():
    con = connect()
    s = sorted({r[0] for r in con.execute('SELECT DISTINCT site FROM usage_daily UNION SELECT DISTINCT site FROM logins')})
    con.close()
    return s


def status(site):
    con = connect()
    days = [r[0] for r in con.execute('SELECT DISTINCT day FROM usage_daily WHERE site=? ORDER BY day', (site,))]
    lg = con.execute('SELECT COUNT(*), MIN(ts), MAX(ts) FROM logins WHERE site=?', (site,)).fetchone()
    imports = con.execute('SELECT imported_at, kind, filename, rows, added, first_ts, last_ts FROM imports WHERE site=? ORDER BY id DESC LIMIT 30', (site,)).fetchall()
    con.close()
    # gaps in daily coverage
    gaps = []
    if days:
        d0, d1 = date.fromisoformat(days[0]), date.fromisoformat(days[-1])
        have = set(days)
        cur = d0
        while cur <= d1:
            if cur.isoformat() not in have:
                gaps.append(cur.isoformat())
            cur = date.fromordinal(cur.toordinal() + 1)
    return {'days': days, 'gaps': gaps, 'logins': lg[0], 'log_first': lg[1] or '', 'log_last': lg[2] or '', 'imports': imports}


def load_for_report(site, start, end):
    """Return (ev dict mac->[(dt,user)], usage_by_day list of (day, {mac: {...}}))."""
    con = connect()
    ev = collections.defaultdict(list)
    for ts, mac, user in con.execute('SELECT ts, mac, user FROM logins WHERE site=? ORDER BY ts', (site,)):
        ev[mac].append((datetime.fromisoformat(ts), user))
    byday = collections.OrderedDict()
    for day, mac, name, up, down, tot, is_mac in con.execute(
            'SELECT day, mac, name, up, down, tot, is_mac FROM usage_daily WHERE site=? AND day BETWEEN ? AND ? ORDER BY day',
            (site, start.isoformat(), end.isoformat())):
        byday.setdefault(date.fromisoformat(day), collections.OrderedDict())[mac] = {
            'name': name, 'up': up, 'down': down, 'tot': tot, 'is_mac': bool(is_mac)}
    con.close()
    return ev, list(byday.items())
