import unittest
import os
import shutil
import tempfile
import pandas as pd
import openpyxl
from unittest.mock import patch, MagicMock
from extractor import JobRequirementExtractor
from main import RUN_LOG_HEADERS, LEGACY_RUN_LOG_HEADERS
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



    def test_ai_gatekeeper_credential_isolation(self):
        """Verify that AI Gatekeeper fails open cleanly without leaking environment credentials."""
        with patch.dict(os.environ, {}, clear=True):
            gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
            self.assertFalse(gk.is_active)
            df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}])
            res = gk.filter_dataframe(df)
            self.assertEqual(len(res), 1)

    def test_ai_gatekeeper_malformed_fit_score(self):
        """Verify that non-numeric fit_score (e.g. 'N/A') does not crash and defaults safely."""
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "test_key"}):
            gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
            mock_res = MagicMock()
            mock_res.status_code = 200
            mock_res.json.return_value = {
                "choices": [{
                    "message": {
                        "content": '{"decisions": [{"id": 0, "is_fresher_eligible": true, "is_domain_match": true, "fit_score": "N/A"}]}'
                    }
                }]
            }
            with patch("requests.post", return_value=mock_res):
                df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}], index=[0])
                res = gk.filter_dataframe(df)
                self.assertEqual(len(res), 1)

    def test_ai_gatekeeper_null_decisions(self):
        """Verify that 'decisions': null does not raise TypeError and fails open."""
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "test_key"}):
            gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
            mock_res = MagicMock()
            mock_res.status_code = 200
            mock_res.json.return_value = {
                "choices": [{
                    "message": {
                        "content": '{"decisions": null}'
                    }
                }]
            }
            with patch("requests.post", return_value=mock_res):
                df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}], index=[0])
                res = gk.filter_dataframe(df)
                self.assertEqual(len(res), 1)

    def test_ai_gatekeeper_string_booleans(self):
        """Verify that string booleans ('false', 'true') are strictly parsed and not evaluated as truthy strings."""
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "test_key"}):
            gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
            mock_res = MagicMock()
            mock_res.status_code = 200
            mock_res.json.return_value = {
                "choices": [{
                    "message": {
                        "content": '{"decisions": [{"id": 0, "is_fresher_eligible": "false", "is_domain_match": "true", "fit_score": 8.0}]}'
                    }
                }]
            }
            with patch("requests.post", return_value=mock_res):
                df = pd.DataFrame([{"title": "Senior Data Analyst", "company": "BigCorp"}], index=[0])
                res = gk.filter_dataframe(df)
                self.assertEqual(len(res), 0)

    def test_extractor_nan_location_and_description(self):
        """Verify that float('nan') or None in location/description does not throw AttributeError."""
        extractor = JobRequirementExtractor(["SQL", "Python"])
        mode = extractor.extract_work_mode(float('nan'), float('nan'))
        self.assertEqual(mode, "On-site")
        
        job_result = extractor.process_job({"location": float('nan'), "description": float('nan')})
        self.assertEqual(job_result["work_mode"], "On-site")
        self.assertEqual(job_result["skills_count"], 0)

    def test_archive_human_edit_not_overwritten_by_unchanged_primary(self):
        """Verify that an archive human edit (e.g. to 'Interviewing') is never overwritten by an untouched primary tracker."""
        temp_dir = tempfile.mkdtemp()
        try:
            leads_dir = os.path.join(temp_dir, "leads")
            os.makedirs(leads_dir, exist_ok=True)
            stable_tracker = os.path.join(temp_dir, "Tracker_Latest.xlsx")
            archive_tracker = os.path.join(leads_dir, "Job_Leads_2026-10-14_011000.xlsx")
            apps_file = os.path.join(temp_dir, "applications.json")
            
            cfg = {
                "output": {
                    "leads_dir": leads_dir,
                    "stable_filename": stable_tracker,
                    "applications_file": apps_file,
                    "seen_jobs_file": os.path.join(temp_dir, "seen.json"),
                }
            }
            gen = ExcelDashboardGenerator(cfg)
            df = pd.DataFrame([{
                "title": "Data Analyst",
                "company": "Cognizant",
                "location": "Hyderabad",
                "platform": "LinkedIn",
                "final_apply_url": "https://www.linkedin.com/jobs/view/123456",
                "skills": "SQL",
                "experience": "0-2 years",
                "work_mode": "Hybrid",
                "job_age": "Recent"
            }])
            
            # 1. Generate primary tracker (stamped with 'Not Applied')
            gen.generate(df, output_path=stable_tracker)
            
            # 2. Generate archive tracker (stamped with 'Not Applied')
            gen.generate(df, output_path=archive_tracker)
            
            # 3. Simulate human editing archive tracker to 'Interviewing'
            wb_arch = openpyxl.load_workbook(archive_tracker)
            ws_arch = wb_arch["Active Job Leads"]
            # Column K is Application Status
            status_col = None
            for col in range(1, ws_arch.max_column + 1):
                if ws_arch.cell(1, col).value == "Application Status":
                    status_col = col
                    break
            self.assertIsNotNone(status_col)
            ws_arch.cell(2, status_col).value = "Interviewing"
            wb_arch.save(archive_tracker)
            wb_arch.close()
            
            # 4. Sync and load applications
            ledger = gen._sync_and_load_applications()
            key = "https://www.linkedin.com/jobs/view/123456"
            self.assertIn(key, ledger)
            # The human edit 'Interviewing' MUST be preserved and NOT reverted to 'Not Applied'!
            self.assertEqual(ledger[key]["status"], "Interviewing")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_primary_tracker_isolation_ignores_other_profiles(self):
        """Verify that Tracker_Latest_IT_Support.xlsx is strictly ignored by the main tracker scanner."""
        temp_dir = tempfile.mkdtemp()
        try:
            stable_tracker = os.path.join(temp_dir, "Tracker_Latest.xlsx")
            it_support_tracker = os.path.join(temp_dir, "Tracker_Latest_IT_Support.xlsx")
            
            wb1 = openpyxl.Workbook()
            wb1.save(stable_tracker)
            wb1.close()
            
            wb2 = openpyxl.Workbook()
            wb2.save(it_support_tracker)
            wb2.close()
            
            cfg = {
                "output": {
                    "leads_dir": os.path.join(temp_dir, "leads"),
                    "stable_filename": stable_tracker,
                    "applications_file": os.path.join(temp_dir, "apps.json"),
                    "seen_jobs_file": os.path.join(temp_dir, "seen.json"),
                }
            }
            gen = ExcelDashboardGenerator(cfg)
            candidates = [p for p, is_prim in gen._get_ordered_sync_candidates()]
            self.assertIn(stable_tracker, candidates)
            self.assertNotIn(it_support_tracker, candidates)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def test_consultant_title_handling(self):
        """Verify that entry-level consultant titles are kept while senior consultants are dropped."""
        self.assertFalse(self.scraper._is_senior_title("Graduate Data Analytics Consultant"))
        self.assertFalse(self.scraper._is_senior_title("Associate Consultant"))
        self.assertFalse(self.scraper._is_senior_title("Trainee Consultant"))
        self.assertFalse(self.scraper._is_senior_title("Junior Consultant"))
        self.assertTrue(self.scraper._is_senior_title("Senior Consultant"))
        self.assertTrue(self.scraper._is_senior_title("Lead Consultant"))
        self.assertTrue(self.scraper._is_senior_title("Package Consultant"))

    def test_run_log_headers_has_dropped_closed_job(self):
        """Verify that dropped_closed_job is in RUN_LOG_HEADERS and legacy schemas are preserved."""
        self.assertIn("dropped_closed_job", RUN_LOG_HEADERS)
        self.assertEqual(len(RUN_LOG_HEADERS), 15)
        self.assertEqual(len(LEGACY_RUN_LOG_HEADERS), 4)


if __name__ == "__main__":
    unittest.main()
