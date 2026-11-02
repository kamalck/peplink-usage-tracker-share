"""Core behaviour: attribution rules, import / delete rules, report views, Excel layout, security headers, audit trail, backup.
Run:  python -m unittest discover -s tests      (scratch databases only)"""
import io, os, unittest
from datetime import date, datetime

import openpyxl
from support import A, B, events_xlsx, flashes, load, signed_in, usage_csv


class Attribution(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()
        import daily
        self.daily = daily

    def test_same_day_login_wins_with_most_logins(self):
        ev = [(datetime(2026, 10, 3, 9), 'ana'), (datetime(2026, 10, 3, 10), 'bob'), (datetime(2026, 10, 3, 11), 'bob')]
        user, basis, _ = self.daily.attribute(ev, date(2026, 10, 3))
        self.assertEqual((user, basis), ('bob', 'login'))

    def test_carried_from_earlier_login(self):
        user, basis, _ = self.daily.attribute([(datetime(2026, 10, 1, 9), 'ana')], date(2026, 10, 5))
        self.assertEqual((user, basis), ('ana', 'carried'))

    def test_inferred_from_next_login_when_log_starts_later(self):
        user, basis, _ = self.daily.attribute([(datetime(2026, 10, 6, 9), 'bob')], date(2026, 10, 3))
        self.assertEqual((user, basis), ('bob', 'inferred'))

    def test_no_login_at_all(self):
        self.assertEqual(self.daily.attribute([], date(2026, 10, 3))[:2], ('', ''))

    def test_period_total_equals_sum_of_days_and_gb_is_mb_over_1024(self):
        ev = {A: [(datetime(2026, 10, 1, 8), 'ana')]}
        byday = [(date(2026, 10, 1), {A: dict(name='h', up=10, down=90, tot=100, is_mac=True)}),
                 (date(2026, 10, 2), {A: dict(name='h', up=20, down=180, tot=200, is_mac=True)})]
        detail, pivot, days, summary = self.daily.compute(ev, byday)
        self.assertEqual(pivot[0]['total'], 300)
        self.assertEqual((pivot[0]['up'], pivot[0]['down']), (30, 270))
        self.assertEqual(pivot[0]['gb'], round(300 / 1024, 2))
        self.assertEqual(summary['total_gb'], round(300 / 1024, 2))


class Imports(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()

    def test_reimport_same_day_replaces_not_adds(self):
        for _ in range(2):
            self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(usage_csv(('h', A, 1, 9))))
        ev, byday = self.store.load_for_report('S', date(2026, 10, 1), date(2026, 10, 1))
        self.assertEqual(byday[0][1][A]['tot'], 10)

    def test_empty_usage_file_is_refused_and_keeps_the_day(self):
        self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(usage_csv(('h', A, 1, 9))))
        with self.assertRaises(ValueError):
            self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(usage_csv()))
        self.assertEqual(self.store.status('S')['days'], ['2026-10-01'])
        with self.assertRaises(ValueError):
            self.store.import_period('S', 'p_2026-10-01_2026-10-07.csv', io.BytesIO(usage_csv()))

    def test_thousands_separators_are_accepted(self):
        csv = (b'Name,MAC Address,Upload (MB),Download (MB),Total (MB)\nh,' + A.encode() + b',"1,234.5","2,000.0","3,234.5"\n')
        self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(csv))
        _, byday = self.store.load_for_report('S', date(2026, 10, 1), date(2026, 10, 1))
        self.assertAlmostEqual(byday[0][1][A]['tot'], 3234.5)

    def test_real_non_number_gets_a_readable_error(self):
        csv = b'Name,MAC Address,Upload (MB),Download (MB),Total (MB)\nh,' + A.encode() + b',abc,1,1\n'
        with self.assertRaisesRegex(ValueError, 'not a number'):
            self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(csv))

    def test_missing_hostname_column_is_ok(self):
        csv = b'MAC Address,Upload (MB),Download (MB),Total (MB)\n' + A.encode() + b',1,2,3\n'
        self.store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(csv))
        _, byday = self.store.load_for_report('S', date(2026, 10, 1), date(2026, 10, 1))
        self.assertEqual(byday[0][1][A]['name'], '')

    def test_events_deduplicated_and_unreadable_dates_counted(self):
        f = lambda: events_xlsx((datetime(2026, 10, 1, 8), A, 'ana'), ('not a date', A, 'ana'))
        first = self.store.import_events('S', 'e.xlsx', f())
        self.assertEqual((first['rows'], first['added'], first['skipped']), (1, 1, 1))
        again = self.store.import_events('S', 'e.xlsx', f())
        self.assertEqual(again['added'], 0)

    def test_period_reversed_dates_are_swapped(self):
        r = self.store.import_period('S', 'p_2026-10-09_2026-10-03.csv', io.BytesIO(usage_csv(('h', A, 1, 2))))
        self.assertEqual((r['start'], r['end']), (date(2026, 10, 3), date(2026, 10, 9)))


class EventDates(unittest.TestCase):
    """Excel-mangled event logs: real dates with day/month swapped next to text dates (found in a real export)."""
    def setUp(self):
        self.app, self.store = load()

    def stored(self):
        import sqlite3
        con = sqlite3.connect(self.store.DB_PATH)
        try:
            return sorted(r[0][:10] for r in con.execute('SELECT ts FROM logins'))
        finally:
            con.close()

    def test_swapped_real_dates_are_repaired_using_the_text_dates(self):
        # true dates 1-3 Oct became real dates 10 Jan/Feb/Mar; 13-15 Oct stayed text "MM/DD/YYYY HH:MM"
        f = events_xlsx((datetime(2026, 1, 10, 8), A, 'ana'), (datetime(2026, 2, 10, 9), A, 'ana'), (datetime(2026, 3, 10, 9), B, 'bob'),
                        ('10/13/2026 07:15', A, 'ana'), ('10/14/2026 23:57', B, 'bob'), ('10/15/2026 00:01', A, 'ana'))
        r = self.store.import_events('S', 'e.xlsx', f)
        self.assertEqual((r['rows'], r['skipped'], r['repaired'], r['text_dates']), (6, 0, 3, 3))
        self.assertEqual(self.stored(), ['2026-10-01', '2026-10-02', '2026-10-03', '2026-10-13', '2026-10-14', '2026-10-15'])

    def test_day_first_text_dates_are_recognised(self):
        f = events_xlsx(('13/10/2026 07:15', A, 'ana'), ('31/10/2026 23:57', B, 'bob'))
        self.store.import_events('S', 'e.xlsx', f)
        self.assertEqual(self.stored(), ['2026-10-13', '2026-10-31'])

    def test_iso_text_dates_and_plain_real_dates_are_left_alone(self):
        f = events_xlsx(('2026-10-13 07:15:00', A, 'ana'), (datetime(2026, 1, 10, 8), B, 'bob'), (datetime(2026, 6, 10, 8), B, 'bob'))
        r = self.store.import_events('S', 'e.xlsx', f)
        self.assertEqual(r['repaired'], 0)                   # no slash-text dates to prove a swap: nothing is guessed
        self.assertEqual(self.stored(), ['2026-01-10', '2026-06-10', '2026-10-13'])

    def test_future_dated_logins_are_not_imported(self):
        f = events_xlsx((datetime(2026, 10, 1, 8), A, 'ana'), (datetime(2099, 1, 10, 8), A, 'ana'))
        r = self.store.import_events('S', 'e.xlsx', f)
        self.assertEqual((r['added'], r['future']), (1, 1))

    def test_import_page_explains_repairs(self):
        c, t = signed_in(self.app)
        f = events_xlsx((datetime(2026, 1, 10, 8), A, 'ana'), ('10/13/2026 07:15', A, 'ana'))
        c.post('/import', data={'_csrf': t, 'site': 'S', 'files': [(f, 'e.xlsx')]}, content_type='multipart/form-data')
        self.assertIn('swapped', ' '.join(flashes(c)))


class EventCsv(unittest.TestCase):
    """Event logs may be .csv as well as .xlsx (the portal's own CSV: month-first text dates, no padding, CRLF)."""
    HEAD = 'Date/Time,Device,SSID,Client MAC,Client Name,Type,Details,Longitude,Latitude\r\n'

    def line(self, when, mac, user, sep=','):
        return sep.join([when, 'dev', 'ssid', mac, 'name', 'WiFi', f'guest account ID "{user}" [login]', '', '']) + '\r\n'

    def setUp(self):
        self.app, self.store = load()

    def stored(self):
        import sqlite3
        con = sqlite3.connect(self.store.DB_PATH)
        try:
            return sorted(r[0][:16] for r in con.execute('SELECT ts FROM logins'))
        finally:
            con.close()

    def test_portal_csv_with_bom_crlf_and_unpadded_dates(self):
        body = ('\ufeff' + self.HEAD + self.line('10/1/2026 0:00', A, 'ana') + self.line('10/31/2026 23:57', B, 'bob')
                + 'Failed login,x\r\n' + self.line('10/2/2026 9:05', A, 'ana')).encode('utf-8')
        r = self.store.import_events('S', 'e.csv', io.BytesIO(body))
        self.assertEqual((r['rows'], r['skipped'], r['repaired']), (3, 0, 0))
        self.assertEqual(self.stored(), ['2026-10-01 00:00', '2026-10-02 09:05', '2026-10-31 23:57'])

    def test_semicolon_csv_and_windows_1252_text(self):
        body = (self.HEAD.replace(',', ';') + self.line('10/13/2026 7:15', A, 'ana', sep=';')
                + self.line('10/14/2026 8:00', B, 'jos\xe9', sep=';')).encode('cp1252')
        r = self.store.import_events('S', 'e.csv', io.BytesIO(body))
        self.assertEqual((r['rows'], r['added']), (2, 2))

    def test_same_logins_from_xlsx_and_csv_are_not_duplicated(self):
        self.store.import_events('S', 'e.xlsx', events_xlsx((datetime(2026, 10, 13, 7, 15), A, 'ana')))
        r = self.store.import_events('S', 'e.csv', io.BytesIO((self.HEAD + self.line('10/13/2026 7:15', A, 'ana')).encode()))
        self.assertEqual((r['rows'], r['added']), (1, 0))

    def test_csv_kind_detection(self):
        import merge
        self.assertEqual(merge.csv_kind(self.HEAD.encode()), 'events')
        self.assertEqual(merge.csv_kind(usage_csv(('h', A, 1, 2))), 'usage')
        self.assertEqual(merge.csv_kind(b'Client MAC,Details,MAC Address,Total (MB)\n'), 'both')
        self.assertEqual(merge.csv_kind(b'a,b\n1,2\n'), 'unknown')

    def test_import_page_routes_each_csv_by_its_header(self):
        c, t = signed_in(self.app)
        files = [(io.BytesIO((self.HEAD + self.line('10/1/2026 0:00', A, 'ana')).encode()), 'S_events.csv'),
                 (io.BytesIO(usage_csv(('h', A, 1, 2))), 'S_usage_2026-10-01.csv'),
                 (io.BytesIO(b'a,b\n1,2\n'), 'S_other_2026-10-01.csv')]
        c.post('/import', data={'_csrf': t, 'site': 'S', 'files': files}, content_type='multipart/form-data')
        text = ' | '.join(flashes(c))
        self.assertIn('S_events.csv: 1 logins read, 1 new', text)
        self.assertIn('S_usage_2026-10-01.csv: 1 devices stored', text)
        self.assertIn('S_other_2026-10-01.csv: ERROR not recognised', text)
        self.assertEqual(self.store.status('S')['days'], ['2026-10-01'])

    def test_quick_merge_accepts_a_csv_event_log(self):
        c, t = signed_in(self.app)
        ev = (self.HEAD + self.line('10/1/2026 0:00', A, 'ana')).encode()
        r = c.post('/quick', data={'_csrf': t, 'action': 'preview', 'events': (io.BytesIO(ev), 'e.csv'),
                                   'usage': [(io.BytesIO(usage_csv(('h', A, 1, 2))), 'a_2026-10-01.csv'), (io.BytesIO(usage_csv(('h', A, 1, 2))), 'a_2026-10-02.csv')]},
                   content_type='multipart/form-data')
        self.assertIn('ana', r.get_data(as_text=True))
        single = c.post('/quick', data={'_csrf': t, 'action': 'preview', 'events': (io.BytesIO(ev), 'e.csv'),
                                        'usage': (io.BytesIO(usage_csv(('h', A, 1, 2))), 'a.csv')}, content_type='multipart/form-data')
        self.assertIn('ana', single.get_data(as_text=True))


class DaysLoaded(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()

    def test_segments_show_ranges_and_gaps_inside_the_span_only(self):
        segs = self.store.day_segments(['2026-10-01', '2026-10-02', '2026-10-05', '2026-10-07', '2026-10-08'])
        self.assertEqual([(x['kind'], x['a'], x['b'], x['n']) for x in segs],
                         [('ok', '2026-10-01', '2026-10-02', 2), ('miss', '2026-10-03', '2026-10-04', 2), ('ok', '2026-10-05', '2026-10-05', 1),
                          ('miss', '2026-10-06', '2026-10-06', 1), ('ok', '2026-10-07', '2026-10-08', 2)])
        self.assertEqual(self.store.day_segments([]), [])
        self.assertEqual([x['kind'] for x in self.store.day_segments(['2026-10-01', '2026-10-02'])], ['ok'])

    def test_import_page_shows_ranges_and_red_missing_days_and_report_lists_gaps(self):
        for d in ('01', '02', '05'):
            self.store.import_usage('S', f'x_2026-10-{d}.csv', io.BytesIO(usage_csv(('h', A, 1, 2))))
        c, _t = signed_in(self.app)
        html = c.get('/site_panel?site=S').get_data(as_text=True)
        self.assertIn('2 days missing', html)
        self.assertIn('rng miss', html)
        self.assertIn('missing 2026-10-03', html)
        self.assertIn('2 ranges', html)
        rep = c.get('/report?site=S&start=2026-10-01&end=2026-10-05').get_data(as_text=True)
        self.assertIn('2 days missing in store', rep)

    def test_complete_history_says_no_gaps(self):
        for d in ('01', '02', '03'):
            self.store.import_usage('S', f'x_2026-10-{d}.csv', io.BytesIO(usage_csv(('h', A, 1, 2))))
        c, _t = signed_in(self.app)
        html = c.get('/site_panel?site=S').get_data(as_text=True)
        self.assertIn('No gaps', html)
        self.assertNotIn('rng miss', html)


class StaleData(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()

    def load_day(self, days_ago):
        from datetime import timedelta
        d = date.today() - timedelta(days=days_ago)
        self.store.import_usage('S', f'x_{d}.csv', io.BytesIO(usage_csv(('h', A, 1, 2))))

    def test_age_rule(self):
        today = date(2026, 11, 10)
        self.assertEqual(self.store.data_age('2026-11-10', today), (0, False))
        self.assertEqual(self.store.data_age('2026-11-08', today), (2, False))     # still normal: yesterday's export may not be in yet
        self.assertEqual(self.store.data_age('2026-11-07', today), (3, True))
        self.assertEqual(self.store.data_age('', today), (None, False))

    def test_warning_shows_on_old_data_only(self):
        self.load_day(7)
        c, _t = signed_in(self.app)
        self.assertIn('Last data 7 days old', c.get('/site_panel?site=S').get_data(as_text=True))
        self.assertIn('7 days old', c.get('/').get_data(as_text=True))             # also in the sites table

    def test_no_warning_when_data_is_current(self):
        self.load_day(1)
        c, _t = signed_in(self.app)
        self.assertNotIn('Last data', c.get('/site_panel?site=S').get_data(as_text=True))
        self.assertNotIn('days old', c.get('/').get_data(as_text=True))


class DeleteAndAudit(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()
        for d in ('01', '02', '03'):
            self.store.import_usage('S', f'x_2026-10-{d}.csv', io.BytesIO(usage_csv(('h', A, 1, 2))), user='admin')
        self.store.import_events('S', 'e.xlsx', events_xlsx((datetime(2026, 10, 1, 8), A, 'ana'), (datetime(2026, 10, 3, 8), A, 'ana')), user='admin')

    def test_usage_delete_respects_the_date_range(self):
        n = self.store.delete_data('S', 'usage', date(2026, 10, 2), date(2026, 10, 3), user='bob')
        self.assertEqual(n, 2)
        self.assertEqual(self.store.status('S')['days'], ['2026-10-01'])

    def test_logins_delete_respects_the_range(self):
        self.store.delete_data('S', 'logins', date(2026, 10, 1), date(2026, 10, 2))
        self.assertEqual(self.store.status('S')['logins'], 1)

    def test_everything_needs_the_site_name_typed(self):
        c, t = signed_in(self.app)
        c.post('/delete', data={'_csrf': t, 'site': 'S', 'what': 'all', 'confirm': 'nope'})
        self.assertEqual(len(self.store.status('S')['days']), 3)
        c.post('/delete', data={'_csrf': t, 'site': 'S', 'what': 'all', 'confirm': 's'})
        self.assertEqual(self.store.status('S')['days'], [])

    def test_audit_trail_records_who_imported_and_deleted(self):
        self.store.delete_data('S', 'usage', date(2026, 10, 1), date(2026, 10, 1), user='bob')
        rows = self.store.status('S')['imports']
        self.assertEqual(rows[0][1], 'delete')
        self.assertEqual(rows[0][7], 'bob')
        self.assertIn('admin', [r[7] for r in rows])

    def test_old_database_without_user_column_is_migrated_in_place(self):
        import sqlite3
        db = self.store.DB_PATH
        con = sqlite3.connect(db)
        con.executescript('DROP TABLE imports; CREATE TABLE imports (id INTEGER PRIMARY KEY AUTOINCREMENT, imported_at TEXT, site TEXT, kind TEXT, filename TEXT, rows INTEGER, added INTEGER, first_ts TEXT, last_ts TEXT);')
        con.commit(); con.close()
        self.store._ready = False
        self.assertEqual(len(self.store.status('S')['days']), 3)       # data kept
        self.store.import_usage('S', 'x_2026-10-09.csv', io.BytesIO(usage_csv(('h', A, 1, 2))), user='admin')
        self.assertEqual(self.store.status('S')['imports'][0][7], 'admin')


class ViewsAndExcel(unittest.TestCase):
    def setUp(self):
        self.app, self.store = load()
        self.store.import_events('S', 'e.xlsx', events_xlsx((datetime(2026, 10, 1, 8), A, 'ana')))
        for d in ('01', '02'):
            self.store.import_usage('S', f'x_2026-10-{d}.csv', io.BytesIO(usage_csv(('laptop', A, 10, 90), ('', B, 1, 2), ('ip', '10.0.0.9', 1, 1))))
        self.c, self.t = signed_in(self.app)
        self.q = '/report?site=S&start=2026-10-01&end=2026-10-02'

    def sheet(self, url, name):
        wb = openpyxl.load_workbook(io.BytesIO(self.c.get(url).data))
        return [[c.value for c in r] for r in wb[name]]

    def test_total_period_is_default_and_has_upload_download(self):
        html = self.c.get(self.q).get_data(as_text=True)
        self.assertIn('GB uploaded', html)
        self.assertIn('Upload GB', html)           # unmatched shows MB and GB

    def test_total_excel_layout_macs_last_and_unmatched_gb(self):
        users = self.sheet(self.q + '&view=total&download=1', 'Usage per user')
        self.assertEqual(users[0][-1], 'MACs (hostname)')
        self.assertEqual(users[1][3], 180)         # download MB over two days
        unm = self.sheet(self.q + '&view=total&download=1', 'Unmatched')
        self.assertIn('Total (GB)', unm[0])

    def test_day_excel_has_metric_sheets_and_device_download_follows_view(self):
        wb = openpyxl.load_workbook(io.BytesIO(self.c.get(self.q + '&view=day&download=1').data))
        for name in ('Usage per user per day', 'User per day - Upload', 'User per day - Download', 'Device per day - Download', 'Device-day detail'):
            self.assertIn(name, wb.sheetnames)
        dev_total = openpyxl.load_workbook(io.BytesIO(self.c.get(self.q + '&view=total&download=device').data))
        self.assertEqual(dev_total.sheetnames, ['Usage per device'])

    def test_period_page_and_downloads(self):
        self.store.import_period('S', 'p_2026-10-01_2026-10-07.csv', io.BytesIO(usage_csv(('laptop', A, 10, 90), ('', B, 1, 2))))
        url = '/period?site=S&start=2026-10-01&end=2026-10-07'
        self.assertIn('GB downloaded', self.c.get(url).get_data(as_text=True))
        wb = openpyxl.load_workbook(io.BytesIO(self.c.get(url + '&download=1').data))
        self.assertEqual(wb.sheetnames[0], 'Usage per device')
        self.assertEqual(self.c.get(url + '&download=device').status_code, 200)

    def test_bad_inputs_redirect_instead_of_crashing(self):
        for u in ('/report?site=S&start=2026-10-05&end=2026-10-01', '/report?site=S&start=garbage', '/period?site=S&start=x&end=y', '/report?site=NOPE'):
            self.assertIn(self.c.get(u).status_code, (200, 302), u)


class Security(unittest.TestCase):
    def test_signed_in_pages_are_not_cacheable_and_have_hardening_headers(self):
        app, _ = load()
        c, _t = signed_in(app)
        r = c.get('/')
        self.assertEqual(r.headers['Cache-Control'], 'no-store')
        self.assertEqual(r.headers['X-Content-Type-Options'], 'nosniff')
        self.assertEqual(r.headers['X-Frame-Options'], 'SAMEORIGIN')
        r = c.get('/static/kck-ui/kck-ui.css')
        self.assertNotEqual(r.headers.get('Cache-Control'), 'no-store')
        r.close()

    def test_login_required_and_health_open(self):
        app, _ = load()
        c = app.app.test_client()
        self.assertEqual(c.get('/').status_code, 302)
        self.assertEqual(c.get('/health').status_code, 200)
        self.assertEqual(c.post('/delete', data={'site': 'S', 'what': 'all'}).status_code, 302)

    def test_extra_accounts_and_csrf(self):
        app, _ = load(ADMIN_USERS='bob:pw2')
        c, t = signed_in(app, 'bob', 'pw2')
        self.assertEqual(c.get('/').status_code, 200)
        c.post('/delete', data={'site': 'S', 'what': 'usage'})        # no CSRF token: refused
        self.assertIn('expired', ' '.join(flashes(c)))

    def test_lockout_after_five_failures_and_per_client_with_proxy(self):
        app, _ = load()
        c = app.app.test_client()
        for _ in range(6):
            c.get('/login')
            with c.session_transaction() as s:
                t = s['csrf']
            r = c.post('/login', data={'_csrf': t, 'username': 'admin', 'password': 'bad'})
        self.assertEqual(r.status_code, 429)
        # behind one proxy each client is counted separately: client 2 is not locked out by client 1
        app2, _ = load(TRUSTED_PROXIES='1')
        c2 = app2.app.test_client()
        def attempt(ip):
            c2.get('/login', headers={'X-Forwarded-For': ip})
            with c2.session_transaction() as s:
                t = s['csrf']
            return c2.post('/login', data={'_csrf': t, 'username': 'admin', 'password': 'bad'}, headers={'X-Forwarded-For': ip}).status_code
        for _ in range(6):
            last = attempt('10.1.1.1')
        self.assertEqual(last, 429)
        self.assertEqual(attempt('10.2.2.2'), 200)

    def test_plain_http_login_failure_explains_the_secure_cookie(self):
        app, _ = load(SESSION_COOKIE_SECURE='1')
        c = app.app.test_client()
        c.get('/login')                                  # Secure cookie is not kept over http, so the CSRF token is lost
        r = c.post('/login', data={'_csrf': 'whatever', 'username': 'admin', 'password': 'pw'})
        text = r.get_data(as_text=True)
        self.assertIn('session expired', text)
        self.assertIn('SESSION_COOKIE_SECURE=0', text)
        app0, _ = load(SESSION_COOKIE_SECURE='0')        # Secure off: no cookie hint, login works over http
        c0, _t = signed_in(app0)
        self.assertEqual(c0.get('/').status_code, 200)

    def test_old_lockout_entries_are_pruned(self):
        app, _ = load()
        app._fails.update({f'1.1.1.{i}': [1, 0, 0] for i in range(50)})
        app._prune_fails(10_000_000)
        self.assertEqual(app._fails, {})


class Backup(unittest.TestCase):
    def test_backup_is_a_usable_copy_and_old_ones_are_pruned(self):
        import sqlite3, tempfile, time
        app, store = load()
        store.import_usage('S', 'x_2026-10-01.csv', io.BytesIO(usage_csv(('h', A, 1, 2))))
        d = tempfile.mkdtemp()
        for _ in range(3):
            out = app.backup_db(d, keep=2)
            time.sleep(1.05)
        self.assertEqual(len(os.listdir(d)), 2)
        con = sqlite3.connect(out)
        n = con.execute('SELECT COUNT(*) FROM usage_daily').fetchone()[0]
        con.close()
        self.assertEqual(n, 1)


if __name__ == '__main__':
    unittest.main()
