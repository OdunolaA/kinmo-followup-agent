import datetime as dt
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import followup as f  # noqa: E402

TODAY = dt.date(2026, 11, 16)


class FollowupTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.cfg = f.load_config(os.path.join(self.tmp, "none.json"))
        self.cfg.update(log_path=os.path.join(self.tmp, "log.json"), drafts_dir=os.path.join(self.tmp, "drafts"),
                        sender_name="Test Sender")
        self.people = {a["name"]: a for a in (f.assess(r, [], self.cfg, TODAY) for r in
                       f.read_tracker(f.resolve(self.cfg, "tracker_path"), "KS"))}

    def test_comment_dates_are_day_month(self):
        dates, future = f.comment_dates("Proposal sent 22/10\nFollow up 2/11", TODAY)
        self.assertEqual(dates, [dt.date(2026, 10, 22), dt.date(2026, 11, 2)])
        self.assertEqual(future, [])

    def test_future_dates_are_flagged(self):
        dates, future = f.comment_dates("Proposal sent 27/11", TODAY)
        self.assertEqual((dates, future), ([], [dt.date(2026, 11, 27)]))

    def test_statuses(self):
        p = self.people
        self.assertEqual(p["Tola Adeyemi"]["status"], "due")          # proposal, 25 days
        self.assertEqual(p["Ngozi Eze"]["stage"], "price")            # "too expensive"
        self.assertEqual(p["Musa Bello"]["status"], "due")            # Last Contacted column, 13 days
        self.assertEqual(p["Kunle Ajayi"]["status"], "wait")          # 4 days ago
        self.assertEqual(p["Aisha Musa"]["status"], "nodate")
        self.assertEqual(p["Emeka Okoro"]["status"], "skip")          # partnership secured
        self.assertEqual(p["Bisi Cole"]["status"], "skip")            # referral, not LinkedIn
        self.assertEqual(p["Femi Lawal"]["status"], "skip")           # replied

    def test_sent_log_resets_the_clock_and_counts(self):
        rec = next(r for r in f.read_tracker(f.resolve(self.cfg, "tracker_path"), "KS") if r["Contact Name"] == "Tola Adeyemi")
        a = f.assess(rec, [{"name": "Tola Adeyemi", "date": "2026-11-14"}], self.cfg, TODAY)
        self.assertEqual(a["status"], "wait")
        log = [{"name": "Tola Adeyemi", "date": "2026-11-01"}, {"name": "Tola Adeyemi", "date": "2026-11-08"}]
        self.assertEqual(f.assess(rec, log, self.cfg, TODAY)["status"], "done")

    def test_drafts_use_the_right_template(self):
        t = f.load_templates()
        tola = f.draft(self.people["Tola Adeyemi"], self.cfg, t, TODAY)
        self.assertIn("proposal I sent on 22 October", tola)
        self.assertTrue(tola.endswith("Test Sender\nKinmo Solutions · building factories across Africa"))
        ngozi = f.draft(self.people["Ngozi Eze"], self.cfg, t, TODAY)
        self.assertIn("technical due diligence", ngozi)
        ifeoma = f.draft(self.people["Ifeoma Obi"], self.cfg, t, TODAY)
        self.assertIn("work as a supplier", ifeoma)

    def test_run_writes_a_page(self):
        due, others = f.run(self.cfg, TODAY)
        self.assertEqual({a["name"] for a in due}, {"Tola Adeyemi", "Ngozi Eze", "Musa Bello", "Ifeoma Obi"})
        page = os.path.join(self.cfg["drafts_dir"], "2026-11-16.html")
        with open(page) as fh:
            html = fh.read()
        self.assertIn("Copy message", html)
        self.assertIn("https://www.linkedin.com/in/example-musa", html)


if __name__ == "__main__":
    unittest.main()
