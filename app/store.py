"""SQLite store for ongoing Starlink usage tracking (multiple sites, overlapping log exports)."""
import os, re, sqlite3, collections
from contextlib import contextmanager
from datetime import datetime, date, timedelta
from merge import parse_usage
from daily import parse_events_by_day, date_from_name, dates_from_name

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
CREATE TABLE IF NOT EXISTS usage_period (
  site TEXT NOT NULL, start TEXT NOT NULL, end TEXT NOT NULL, mac TEXT NOT NULL, name TEXT, up REAL, down REAL, tot REAL,
  pct REAL, is_mac INTEGER,
  PRIMARY KEY (site, start, end, mac));
CREATE TABLE IF NOT EXISTS imports (
  id INTEGER PRIMARY KEY AUTOINCREMENT, imported_at TEXT, site TEXT, kind TEXT, filename TEXT,
  rows INTEGER, added INTEGER, first_ts TEXT, last_ts TEXT);
"""


def site_from_name(name):
    m = SITE_RE.search(name or '')
    return m.group(1).upper() if m else ''


_ready = False


@contextmanager
def connect():
    """Open the database; commit on success, roll back on error, ALWAYS close (sqlite3's own `with` never closes)."""
    global _ready
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=30)
    try:
        if not _ready:
            con.executescript(SCHEMA)
            _ready = True
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


def import_events(site, filename, fileobj):
    ev = parse_events_by_day(fileobj)
    rows = [(site, dt.isoformat(sep=' '), mac, user) for mac, lst in ev.items() for dt, user in lst]
    ts = sorted(r[1] for r in rows)
    with connect() as con:
        before = con.execute('SELECT COUNT(*) FROM logins WHERE site=?', (site,)).fetchone()[0]
        con.executemany('INSERT OR IGNORE INTO logins VALUES (?,?,?,?)', rows)
        after = con.execute('SELECT COUNT(*) FROM logins WHERE site=?', (site,)).fetchone()[0]
        con.execute('INSERT INTO imports (imported_at, site, kind, filename, rows, added, first_ts, last_ts) VALUES (?,?,?,?,?,?,?,?)',
                    (datetime.now().isoformat(sep=' ', timespec='seconds'), site, 'events', filename, len(rows), after - before,
                     ts[0] if ts else None, ts[-1] if ts else None))
    return {'rows': len(rows), 'added': after - before, 'first': ts[0] if ts else '', 'last': ts[-1] if ts else ''}


def import_usage(site, filename, fileobj, day=None):
    day = day or date_from_name(filename)
    if not day:
        raise ValueError(f'No date (YYYY-MM-DD) in file name: {filename}')
    usage = parse_usage(fileobj)
    with connect() as con:
        con.execute('DELETE FROM usage_daily WHERE site=? AND day=?', (site, day.isoformat()))
        con.executemany('INSERT INTO usage_daily VALUES (?,?,?,?,?,?,?,?)',
                        [(site, day.isoformat(), mac, u['name'], u['up'], u['down'], u['tot'], int(u['is_mac'])) for mac, u in usage.items()])
        con.execute('INSERT INTO imports (imported_at, site, kind, filename, rows, added, first_ts, last_ts) VALUES (?,?,?,?,?,?,?,?)',
                    (datetime.now().isoformat(sep=' ', timespec='seconds'), site, 'usage', filename, len(usage), len(usage),
                     day.isoformat(), day.isoformat()))
    return {'rows': len(usage), 'day': day}


def import_period(site, filename, fileobj, start=None, end=None):
    """A usage export covering several days (e.g. 7). Dates come from the form, else from two dates in the file name."""
    if not start:
        ds = dates_from_name(filename)
        if len(ds) >= 2:
            start, end = ds[0], ds[1]
        elif len(ds) == 1:
            start = ds[0]
    if not start:
        raise ValueError(f'No period start date: give one in the form or put two dates (YYYY-MM-DD) in the file name: {filename}')
    end = end or start + timedelta(days=6)
    if end < start:
        start, end = end, start
    usage = parse_usage(fileobj)
    with connect() as con:
        con.execute('DELETE FROM usage_period WHERE site=? AND start=? AND end=?', (site, start.isoformat(), end.isoformat()))
        con.executemany('INSERT INTO usage_period VALUES (?,?,?,?,?,?,?,?,?,?)',
                        [(site, start.isoformat(), end.isoformat(), mac, u['name'], u['up'], u['down'], u['tot'], u['pct'], int(u['is_mac']))
                         for mac, u in usage.items()])
        con.execute('INSERT INTO imports (imported_at, site, kind, filename, rows, added, first_ts, last_ts) VALUES (?,?,?,?,?,?,?,?)',
                    (datetime.now().isoformat(sep=' ', timespec='seconds'), site, 'period', filename, len(usage), len(usage),
                     start.isoformat(), end.isoformat()))
    return {'rows': len(usage), 'start': start, 'end': end}


def load_period(site, start, end):
    """Return (ev, usage {mac: {...}}) for a stored period, or (ev, None) when it does not exist."""
    ev = collections.defaultdict(list)
    usage = collections.OrderedDict()
    with connect() as con:
        for ts, mac, user in con.execute('SELECT ts, mac, user FROM logins WHERE site=? ORDER BY ts', (site,)):
            ev[mac].append((datetime.fromisoformat(ts), user))
        for mac, name, up, down, tot, pct, is_mac in con.execute(
                'SELECT mac, name, up, down, tot, pct, is_mac FROM usage_period WHERE site=? AND start=? AND end=?',
                (site, start.isoformat(), end.isoformat())):
            usage[mac] = {'name': name, 'up': up, 'down': down, 'tot': tot, 'pct': pct, 'is_mac': bool(is_mac)}
    return ev, (usage or None)


def delete_data(site, what, start=None, end=None):
    """Delete stored data for one site. what: usage | periods | logins | all. start/end (dates, optional) limit
    usage/logins to that range and periods to those overlapping it; 'all' ignores them. Returns rows removed."""
    s, e = (start.isoformat() if start else ''), (end.isoformat() if end else '9999-12-31')
    n = 0
    with connect() as con:
        if what in ('usage', 'all'):
            n += con.execute('DELETE FROM usage_daily WHERE site=? AND day BETWEEN ? AND ?',
                             (site, '' if what == 'all' else s, '9999-12-31' if what == 'all' else e)).rowcount
        if what in ('periods', 'all'):
            n += con.execute('DELETE FROM usage_period WHERE site=? AND end>=? AND start<=?',
                             (site, '' if what == 'all' else s, '9999-12-31' if what == 'all' else e)).rowcount
        if what in ('logins', 'all'):
            n += con.execute('DELETE FROM logins WHERE site=? AND substr(ts,1,10) BETWEEN ? AND ?',
                             (site, '' if what == 'all' else s, '9999-12-31' if what == 'all' else e)).rowcount
        if what == 'all':
            con.execute('DELETE FROM imports WHERE site=?', (site,))
    return n


def sites():
    with connect() as con:
        return sorted({r[0] for r in con.execute(
            'SELECT DISTINCT site FROM usage_daily UNION SELECT DISTINCT site FROM logins UNION SELECT DISTINCT site FROM usage_period')})


def status(site):
    with connect() as con:
        days = [r[0] for r in con.execute('SELECT DISTINCT day FROM usage_daily WHERE site=? ORDER BY day', (site,))]
        periods = con.execute('SELECT start, end, COUNT(*), SUM(tot) FROM usage_period WHERE site=? GROUP BY start, end ORDER BY start', (site,)).fetchall()
        lg = con.execute('SELECT COUNT(*), MIN(ts), MAX(ts) FROM logins WHERE site=?', (site,)).fetchone()
        imports = con.execute('SELECT imported_at, kind, filename, rows, added, first_ts, last_ts FROM imports WHERE site=? ORDER BY id DESC LIMIT 30', (site,)).fetchall()
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
    return {'days': days, 'gaps': gaps, 'logins': lg[0], 'log_first': lg[1] or '', 'log_last': lg[2] or '', 'imports': imports, 'periods': periods}


def load_for_report(site, start, end):
    """Return (ev dict mac->[(dt,user)], usage_by_day list of (day, {mac: {...}}))."""
    ev = collections.defaultdict(list)
    byday = collections.OrderedDict()
    with connect() as con:
        for ts, mac, user in con.execute('SELECT ts, mac, user FROM logins WHERE site=? ORDER BY ts', (site,)):
            ev[mac].append((datetime.fromisoformat(ts), user))
        for day, mac, name, up, down, tot, is_mac in con.execute(
                'SELECT day, mac, name, up, down, tot, is_mac FROM usage_daily WHERE site=? AND day BETWEEN ? AND ? ORDER BY day',
                (site, start.isoformat(), end.isoformat())):
            byday.setdefault(date.fromisoformat(day), collections.OrderedDict())[mac] = {
                'name': name, 'up': up, 'down': down, 'tot': tot, 'is_mac': bool(is_mac)}
    return ev, list(byday.items())
