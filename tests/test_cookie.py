"""Login cookie flags. Run:  python -m unittest discover -s tests   (uses a scratch database, never data/)."""
import importlib, os, sys, tempfile, unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app'))


def load_app(secure=None):
    os.environ.update(DB_PATH=os.path.join(tempfile.mkdtemp(), 'test.db'), ADMIN_PASSWORD='pw', SECRET_KEY='test-key')
    os.environ.pop('SESSION_COOKIE_NAME', None)
    if secure is None:
        os.environ.pop('SESSION_COOKIE_SECURE', None)
    else:
        os.environ['SESSION_COOKIE_SECURE'] = secure
    for m in ('app', 'store', 'merge', 'daily'):
        sys.modules.pop(m, None)
    return importlib.import_module('app').app


def login_cookie(flask_app, scheme='https'):
    c = flask_app.test_client()
    base = f'{scheme}://localhost'
    c.get('/login', base_url=base)
    with c.session_transaction(base_url=base) as s:        # CSRF token set by the login page
        token = s['csrf']
    r = c.post('/login', data={'_csrf': token, 'username': 'admin', 'password': 'pw'}, base_url=base)
    return [h for h in r.headers.getlist('Set-Cookie') if h.startswith('peplink_session=')]


class CookieTests(unittest.TestCase):
    def test_defaults_are_secure_httponly_lax_and_uniquely_named(self):
        cookies = login_cookie(load_app())
        self.assertEqual(len(cookies), 1, cookies)
        c = cookies[0]
        self.assertIn('Secure', c)
        self.assertIn('HttpOnly', c)
        self.assertIn('SameSite=Lax', c)
        self.assertNotIn('Domain=', c)
        self.assertFalse(c.lower().startswith('session='))

    def test_secure_can_be_turned_off_for_plain_http(self):
        app = load_app(secure='0')
        self.assertFalse(app.config['SESSION_COOKIE_SECURE'])
        c = login_cookie(app, scheme='http')[0]
        self.assertNotIn('Secure', c)
        self.assertIn('HttpOnly', c)
        self.assertIn('SameSite=Lax', c)


if __name__ == '__main__':
    unittest.main()
