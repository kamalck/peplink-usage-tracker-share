"""Shared helpers: load the app against a scratch database (never data/) and build small sample files."""
import importlib, io, os, sys, tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app'))
import openpyxl

HEADER = "Name,MAC Address,Upload (MB),Download (MB),Total (MB)\n"
A, B = 'AA:BB:CC:00:00:01', 'AA:BB:CC:00:00:02'


def load(**env):
    """Fresh app + modules on a new scratch database. Returns (app module, store module)."""
    for k in ('SESSION_COOKIE_NAME', 'SESSION_COOKIE_SECURE', 'TRUSTED_PROXIES', 'ADMIN_USERS', 'ADMIN_USERNAME'):
        os.environ.pop(k, None)
    os.environ.update(DB_PATH=os.path.join(tempfile.mkdtemp(), 'test.db'), ADMIN_PASSWORD='pw', SECRET_KEY='k', SESSION_COOKIE_SECURE='0')
    os.environ.update(env)
    for m in ('app', 'store', 'merge', 'daily'):
        sys.modules.pop(m, None)
    mod = importlib.import_module('app')
    return mod, importlib.import_module('store')


def usage_csv(*rows):
    """rows: (name, mac, up, down) -> CSV text, total = up + down."""
    return (HEADER + ''.join(f'{n},{m},{u},{d},{u + d}\n' for n, m, u, d in rows)).encode()


def events_xlsx(*logins):
    """logins: (datetime-or-text, mac, user)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date/Time', 'Client MAC', 'Details'])
    for when, mac, user in logins:
        ws.append([when, mac, f'guest account ID "{user}" [login]'])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def signed_in(app_mod, user='admin', password='pw'):
    c = app_mod.app.test_client()
    c.get('/login')
    with c.session_transaction() as s:
        t = s['csrf']
    c.post('/login', data={'_csrf': t, 'username': user, 'password': password})
    c.get('/')
    with c.session_transaction() as s:
        return c, s['csrf']


def flashes(client):
    with client.session_transaction() as s:
        return [m for _, m in s.pop('_flashes', [])]
