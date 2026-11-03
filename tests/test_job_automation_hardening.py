import unittest
import tempfile
import shutil
import os
import openpyxl
import pandas as pd
from scraper import MultiPlatformJobScraper
from ai_gatekeeper import AIGatekeeper, parse_json_loose
from excel_generator import ExcelDashboardGenerator


class TestJobAutomationHardening(unittest.TestCase):
    def setUp(self):
        self.config = {
            "search": {
                "max_experience_years": 2,
                "max_range_ceiling_years": 3.0,
                "job_titles": ["Data Analyst", "SQL Developer"]
            },
            "output": {"seen_jobs_file": "seen_jobs_test.json"}
        }
        self.scraper = MultiPlatformJobScraper(self.config)

    def test_excluded_title_blacklist(self):
        """Verify that off-domain and disguised senior titles are strictly dropped."""
        # Non-tech / Off-domain
        self.assertTrue(self.scraper._is_senior_title("Receptionist"))
        self.assertTrue(self.scraper._is_senior_title("Front Desk Receptionist"))
        self.assertTrue(self.scraper._is_senior_title("Front Office Executive"))
        self.assertTrue(self.scraper._is_senior_title("Customer Service Executive"))
        self.assertTrue(self.scraper._is_senior_title("Voice Process Associate"))
        self.assertTrue(self.scraper._is_senior_title("People Operations Intern"))

        # Lateral / Senior designations
        self.assertTrue(self.scraper._is_senior_title("Senior Data Analyst"))
        self.assertTrue(self.scraper._is_senior_title("Lead SQL Developer"))
        self.assertTrue(self.scraper._is_senior_title("Data Architect"))
        self.assertTrue(self.scraper._is_senior_title("Infrastructure Engineer II"))

        # Authentic entry-level roles MUST pass
        self.assertFalse(self.scraper._is_senior_title("Data Analyst"))
        self.assertFalse(self.scraper._is_senior_title("Junior Data Analyst"))
        self.assertFalse(self.scraper._is_senior_title("SQL Developer"))
        self.assertFalse(self.scraper._is_senior_title("Power BI Developer"))
        self.assertFalse(self.scraper._is_senior_title("Data Annotator"))
        self.assertFalse(self.scraper._is_senior_title("Graduate Trainee"))

    def test_range_ceiling_rejections(self):
        """Verify that wide ranges exceeding max_upper_years are dropped."""
        # Wide ranges (must be rejected)
        self.assertFalse(self.scraper._within_ceiling("0 - 7 Years", max_years=2, max_upper_years=3.0))
        self.assertFalse(self.scraper._within_ceiling("0 - 5 Years", max_years=2, max_upper_years=3.0))
        self.assertFalse(self.scraper._within_ceiling("2 - 5 Years", max_years=2, max_upper_years=3.0))
        self.assertFalse(self.scraper._within_ceiling("2 To 6 Years", max_years=2, max_upper_years=3.0))
        self.assertFalse(self.scraper._within_ceiling("2-4 Years", max_years=2, max_upper_years=3.0))
        self.assertFalse(self.scraper._within_ceiling("3-5 YOE", max_years=2, max_upper_years=3.0))

        # Authentic entry-level ranges (must pass)
        self.assertTrue(self.scraper._within_ceiling("0 - 2 Years", max_years=2, max_upper_years=3.0))
        self.assertTrue(self.scraper._within_ceiling("1 - 3 Years", max_years=2, max_upper_years=3.0))
        self.assertTrue(self.scraper._within_ceiling("6 months to 2 years", max_years=2, max_upper_years=3.0))
        self.assertTrue(self.scraper._within_ceiling("Fresher", max_years=2, max_upper_years=3.0))
        self.assertTrue(self.scraper._within_ceiling("Entry / Fresher", max_years=2, max_upper_years=3.0))

    def test_closed_job_pattern_matching(self):
        """Verify that closed, expired, and filled job postings are detected."""
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("Note: No longer accepting applications for this role.")))
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("We are not accepting applications at this time.")))
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("This job has expired.")))
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("Job posting has expired.")))
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("Position has been filled.")))
        self.assertTrue(bool(self.scraper.CLOSED_JOB_PATTERN.search("Applications are closed.")))

        # Active job descriptions should NOT match
        self.assertFalse(bool(self.scraper.CLOSED_JOB_PATTERN.search("We are looking for a Data Analyst with SQL and Power BI experience.")))
        self.assertFalse(bool(self.scraper.CLOSED_JOB_PATTERN.search("Join our Data Analytics team in Hyderabad!")))

    def test_enrich_jobs_drops_closed_postings(self):
        """Verify that _enrich_jobs automatically prunes closed postings."""
        df = pd.DataFrame([
            {
                "title": "Data Analyst",
                "company": "Active Corp",
                "job_url_direct": "https://www.linkedin.com/jobs/view/1234567890",
                "description": "Looking for entry level Data Analyst with SQL.",
                "location": "Hyderabad"
            },
            {
                "title": "Data Analyst",
                "company": "Closed Corp",
                "job_url_direct": "https://www.linkedin.com/jobs/view/9876543210",
                "description": "No longer accepting applications. Thank you.",
                "location": "Hyderabad"
            }
        ])
        enriched = self.scraper._enrich_jobs(df)
        companies = list(enriched["company"])
        self.assertIn("Active Corp", companies)
        self.assertNotIn("Closed Corp", companies)
        self.assertGreaterEqual(self.scraper.last_run_stats.get("dropped_closed_job", 0), 1)

    def test_ai_gatekeeper_fail_open(self):
        """Verify that AI Gatekeeper fails open when no key is set."""
        gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
        df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}])
        res = gk.filter_dataframe(df)
        self.assertEqual(len(res), 1)

    def test_parse_json_loose(self):
        """Verify JSON parsing handles raw and markdown-wrapped responses."""
        raw = '{"decisions": [{"id": 0, "fit": true}]}'
        md = '```json\n{"decisions": [{"id": 0, "fit": true}]}\n```'
        self.assertEqual(parse_json_loose(raw), {"decisions": [{"id": 0, "fit": True}]})
        self.assertEqual(parse_json_loose(md), {"decisions": [{"id": 0, "fit": True}]})

    def test_ignored_and_not_accepting_statuses(self):
        """Verify that 'Ignored' and 'Not Accepting Applications' exist and are excluded from Applied tab."""
        temp_dir = tempfile.mkdtemp()
        try:
            cfg = {
                "output": {
                    "leads_dir": os.path.join(temp_dir, "leads"),
                    "stable_filename": os.path.join(temp_dir, "Tracker.xlsx"),
                    "applications_file": os.path.join(temp_dir, "apps.json"),
                    "seen_jobs_file": os.path.join(temp_dir, "seen.json"),
                }
            }
            gen = ExcelDashboardGenerator(cfg)
            self.assertIn("Ignored", gen.status_options)
            self.assertIn("Not Accepting Applications", gen.status_options)

            # Seed ledger with 1 Applied, 1 Ignored, and 1 Not Accepting Applications
            gen.applications_ledger = {
                "https://example.com/1": {"status": "Applied", "notes": "", "company": "Co A", "title": "Analyst"},
                "https://example.com/2": {"status": "Ignored", "notes": "Not interested", "company": "Co B", "title": "SQL Dev"},
                "https://example.com/3": {"status": "Not Accepting Applications", "notes": "Expired", "company": "Co C", "title": "BI Dev"}
            }
            df = pd.DataFrame([
                {"title": "Analyst", "company": "Co A", "final_apply_url": "https://example.com/1"},
                {"title": "SQL Dev", "company": "Co B", "final_apply_url": "https://example.com/2"},
                {"title": "BI Dev", "company": "Co C", "final_apply_url": "https://example.com/3"}
            ])
            out_file = gen.generate(df)
            wb = openpyxl.load_workbook(out_file)
            ws_applied = wb["Applied & In Progress"]
            applied_companies = [ws_applied.cell(r, 2).value for r in range(2, ws_applied.max_row + 1)]
            self.assertIn("Co A", applied_companies)
            self.assertNotIn("Co B", applied_companies)
            self.assertNotIn("Co C", applied_companies)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
