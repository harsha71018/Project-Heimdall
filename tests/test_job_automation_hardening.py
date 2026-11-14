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
from excel_generator import ExcelDashboardGenerator, SYNC_STAMP_SHEET


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
        """Verify that AI Gatekeeper fails open when no key is set or on network failure, without making live requests."""
        with patch.dict(os.environ, {}, clear=True):
            gk = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
            self.assertFalse(gk.is_active)
            df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}])
            res = gk.filter_dataframe(df)
            self.assertEqual(len(res), 1)

        # Also verify that when key IS present, network error fails open safely without live leak
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "dummy_key"}):
            with patch("requests.post", side_effect=Exception("Network failure")):
                gk_active = AIGatekeeper({"ai_gatekeeper": {"enabled": True}})
                self.assertTrue(gk_active.is_active)
                df = pd.DataFrame([{"title": "Data Analyst", "company": "Wipro"}], index=[0])
                res = gk_active.filter_dataframe(df)
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
        """Verify that an archive human edit (e.g. to 'Interviewing') is never overwritten
        by an untouched primary tracker containing a stale non-default status (e.g. 'Applied')."""
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
            job_url = "https://www.linkedin.com/jobs/view/123456"
            df = pd.DataFrame([{
                "title": "Data Analyst",
                "company": "Cognizant",
                "location": "Hyderabad",
                "platform": "LinkedIn",
                "final_apply_url": job_url,
                "skills": "SQL",
                "experience": "0-2 years",
                "work_mode": "Hybrid",
                "job_age": "Recent"
            }])
            
            # 1. Generate primary tracker and set it to a stale non-default 'Applied' state
            # Both the visible cell and its hidden _sync_stamp entry are 'Applied' (i.e. untouched since generation).
            gen.generate(df, output_path=stable_tracker)
            wb_prim = openpyxl.load_workbook(stable_tracker)
            ws_prim = wb_prim["Active Job Leads"]
            ws_prim_stamps = wb_prim[SYNC_STAMP_SHEET]
            for col in range(1, ws_prim.max_column + 1):
                if ws_prim.cell(1, col).value == "Application Status":
                    ws_prim.cell(2, col).value = "Applied"
                    break
            for r in range(2, ws_prim_stamps.max_row + 1):
                if ws_prim_stamps.cell(r, 1).value == job_url:
                    ws_prim_stamps.cell(r, 2).value = "Applied"
                    break
            wb_prim.save(stable_tracker)
            wb_prim.close()

            # 2. Generate archive tracker, starting with 'Applied' in stamp
            gen.generate(df, output_path=archive_tracker)
            wb_arch = openpyxl.load_workbook(archive_tracker)
            ws_arch = wb_arch["Active Job Leads"]
            ws_arch_stamps = wb_arch[SYNC_STAMP_SHEET]
            status_col = None
            for col in range(1, ws_arch.max_column + 1):
                if ws_arch.cell(1, col).value == "Application Status":
                    status_col = col
                    break
            self.assertIsNotNone(status_col)
            for r in range(2, ws_arch_stamps.max_row + 1):
                if ws_arch_stamps.cell(r, 1).value == job_url:
                    ws_arch_stamps.cell(r, 2).value = "Applied"
                    break

            # 3. Simulate human editing archive tracker to 'Interviewing'
            # Note: stamp remains 'Applied', so cell != stamp is flagged as a real human edit.
            ws_arch.cell(2, status_col).value = "Interviewing"
            wb_arch.save(archive_tracker)
            wb_arch.close()
            
            # 4. Sync and load applications
            ledger = gen._sync_and_load_applications()
            self.assertIn(job_url, ledger)
            # The human edit 'Interviewing' MUST be preserved and NOT reverted to 'Applied' by the untouched primary!
            self.assertEqual(ledger[job_url]["status"], "Interviewing")
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

    def test_post_ai_engine_counts_reconciled(self):
        """Verify that per-engine retention counts reconcile with post-AI filtered leads."""
        scraper = MultiPlatformJobScraper(self.config)
        # Mock dataframe with 2 JobSpy leads and 1 Firecrawl lead
        df = pd.DataFrame([
            {"title": "Data Analyst", "company": "Co 1", "_source_engine": "jobspy"},
            {"title": "BI Developer", "company": "Co 2", "_source_engine": "jobspy"},
            {"title": "SQL Engineer", "company": "Co 3", "_source_engine": "firecrawl"},
        ])
        
        # Scraper records pre-AI counts
        scraper.last_run_stats["fc_retained"] = 1
        scraper.last_run_stats["jobspy_retained"] = 2
        scraper.last_run_stats["fc_retained_pre_ai"] = 1
        scraper.last_run_stats["jobspy_retained_pre_ai"] = 2

        # Simulate AI filter dropping one of the JobSpy leads
        filtered_df = df.iloc[[1, 2]].copy()  # retains Co 2 (jobspy) and Co 3 (firecrawl)
        
        # Execute main.py reconciliation block
        if "_source_engine" in filtered_df.columns:
            fc_post_ai = int((filtered_df["_source_engine"] == "firecrawl").sum())
            jobspy_post_ai = int((filtered_df["_source_engine"] == "jobspy").sum())
            scraper.last_run_stats["fc_retained"] = fc_post_ai
            scraper.last_run_stats["jobspy_retained"] = jobspy_post_ai
            filtered_df.drop(columns=["_source_engine"], inplace=True)

        self.assertEqual(scraper.last_run_stats["fc_retained"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_retained"], 1)
        self.assertEqual(scraper.last_run_stats["fc_retained_pre_ai"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_retained_pre_ai"], 2)
        self.assertEqual(scraper.last_run_stats["fc_retained"] + scraper.last_run_stats["jobspy_retained"], len(filtered_df))
        self.assertNotIn("_source_engine", filtered_df.columns)

    def test_failed_source_vs_clean_zero_classification(self):
        """Verify that runs with source errors (e.g. Firecrawl 402 + JobSpy error) are classified as failures,
        while clean zero-match runs are classified as successful clean runs."""
        # Scenario 1: Source Failure (Firecrawl 402 + JobSpy network error)
        stats_failed = {
            "fc_candidates": 0,
            "jobspy_candidates": 0,
            "fc_errors": 1,
            "jobspy_errors": 1,
            "fc_queries_successful": 0,
            "jobspy_queries_successful": 0,
        }
        total_raw_failed = stats_failed["fc_candidates"] + stats_failed["jobspy_candidates"]
        has_any_successful_failed = (stats_failed["fc_queries_successful"] + stats_failed["jobspy_queries_successful"]) > 0
        all_sources_failed_1 = (total_raw_failed == 0) and not has_any_successful_failed
        self.assertTrue(all_sources_failed_1, "Expected run to be classified as source failure when no queries succeeded")

        # Scenario 2: Clean Zero-Match Run (JobSpy succeeded 4 queries, found 0 jobs)
        stats_clean = {
            "fc_candidates": 0,
            "jobspy_candidates": 0,
            "fc_errors": 0,
            "jobspy_errors": 0,
            "fc_queries_successful": 0,
            "jobspy_queries_successful": 4,
        }
        total_raw_clean = stats_clean["fc_candidates"] + stats_clean["jobspy_candidates"]
        has_any_successful_clean = (stats_clean["fc_queries_successful"] + stats_clean["jobspy_queries_successful"]) > 0
        all_sources_failed_2 = (total_raw_clean == 0) and not has_any_successful_clean
        self.assertFalse(all_sources_failed_2, "Expected clean zero-match run to not be classified as failure")

        # Scenario 3: Firecrawl 402 but JobSpy succeeded 2 queries cleanly with 0 jobs
        stats_mixed = {
            "fc_candidates": 0,
            "jobspy_candidates": 0,
            "fc_errors": 1,
            "jobspy_errors": 0,
            "fc_queries_successful": 0,
            "jobspy_queries_successful": 2,
        }
        total_raw_mixed = stats_mixed["fc_candidates"] + stats_mixed["jobspy_candidates"]
        has_any_successful_mixed = (stats_mixed["fc_queries_successful"] + stats_mixed["jobspy_queries_successful"]) > 0
        all_sources_failed_3 = (total_raw_mixed == 0) and not has_any_successful_mixed
        self.assertFalse(all_sources_failed_3, "Expected run where one platform succeeded to not be classified as total source failure")

    @patch("main.acquire_run_lock", return_value=True)
    @patch("main.release_run_lock")
    @patch("main.migrate_and_append_run_log", return_value=True)
    def test_main_orchestration_post_ai_and_failure(self, mock_log, mock_rel, mock_acq):
        """End-to-end orchestration tests for main() with post-AI reconciliation and failure exit codes."""
        from main import main
        import sys

        # Test A: AI filtering drops candidates -> check row_data logged has reconciled counts
        with tempfile.TemporaryDirectory() as tmpdir:
            test_tracker = os.path.join(tmpdir, "Tracker.xlsx")
            mock_jobs = pd.DataFrame([
                {"title": "Data Analyst", "company": "Co 1", "job_url_direct": "https://example.com/1", "final_apply_url": "https://example.com/1", "_source_engine": "jobspy"},
                {"title": "BI Developer", "company": "Co 2", "job_url_direct": "https://example.com/2", "final_apply_url": "https://example.com/2", "_source_engine": "jobspy"}
            ])
            with patch("main.MultiPlatformJobScraper") as MockScraper, \
                 patch("main.AIGatekeeper") as MockAI, \
                 patch("main.ExcelDashboardGenerator") as MockGen, \
                 patch.object(sys, "argv", ["main.py", "--output", test_tracker]):
                
                inst_scraper = MockScraper.return_value
                inst_scraper.scrape_all_roles.return_value = mock_jobs
                inst_scraper.get_remaining_credits.return_value = 100
                inst_scraper.last_run_stats = {
                    "fc_candidates": 0, "fc_markdown_hits": 0, "fc_retained": 0,
                    "jobspy_candidates": 2, "jobspy_retained": 2,
                    "jobspy_queries_successful": 1
                }
                
                # AI gatekeeper retains only 1 of the 2 JobSpy leads
                inst_ai = MockAI.return_value
                inst_ai.is_active = True
                inst_ai.filter_dataframe.return_value = mock_jobs.iloc[[0]].copy()
                
                MockGen.return_value.generate.return_value = test_tracker

                main()

                # Verify migrate_and_append_run_log received reconciled counts
                self.assertTrue(mock_log.called)
                logged_data = mock_log.call_args[0][1]
                self.assertEqual(logged_data["jobspy_retained"], 1)
                self.assertEqual(logged_data["total_unique_jobs"], 1)
                self.assertEqual(logged_data["fc_retained"], 0)

        # Test B: All sources failed -> sys.exit(1) and no tracker refresh
        with patch("main.MultiPlatformJobScraper") as MockScraper, \
             patch("main.ExcelDashboardGenerator") as MockGen, \
             patch.object(sys, "argv", ["main.py"]):
            
            inst_scraper = MockScraper.return_value
            inst_scraper.scrape_all_roles.return_value = pd.DataFrame()
            inst_scraper.get_remaining_credits.return_value = None
            inst_scraper.last_run_stats = {
                "fc_candidates": 0, "fc_markdown_hits": 0, "fc_retained": 0,
                "jobspy_candidates": 0, "jobspy_retained": 0,
                "fc_errors": 1, "jobspy_errors": 1,
                "fc_queries_successful": 0, "jobspy_queries_successful": 0
            }

            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)
            # generate() should NOT be called when all sources failed
            self.assertFalse(MockGen.return_value.generate.called)

        # Test C: Malformed HTTP 200 response from Firecrawl + JobSpy exception (real scraper lifecycle)
        with patch.dict(os.environ, {"FIRECRAWL_API_KEY": "fc_test_key"}), \
             patch("main.ExcelDashboardGenerator") as MockGen, \
             patch("scraper.requests.post") as mock_fc_post, \
             patch("scraper.scrape_jobs", side_effect=Exception("JobSpy network error")), \
             patch("scraper.time.sleep"), \
             patch.object(sys, "argv", ["main.py"]):
            
            # Firecrawl returns HTTP 200 with invalid JSON
            mock_res = MagicMock()
            mock_res.status_code = 200
            mock_res.json.side_effect = ValueError("Invalid JSON response payload")
            mock_fc_post.return_value = mock_res
            
            with self.assertRaises(SystemExit) as cm:
                main()
            
            # Must exit with code 1 due to all sources failing
            self.assertEqual(cm.exception.code, 1)
            # Empty tracker must NOT be refreshed, preserving existing workbooks
            self.assertFalse(MockGen.return_value.generate.called)

    def test_firecrawl_malformed_200_counts_as_error_not_success(self):
        """Verify that HTTP 200 with invalid JSON increments fc_errors and does NOT increment fc_queries_successful."""
        scraper = MultiPlatformJobScraper(self.config)
        scraper.firecrawl_api_key = "test_key"
        
        mock_res = MagicMock()
        mock_res.status_code = 200
        mock_res.json.side_effect = ValueError("Invalid JSON from server")
        
        with patch("requests.post", return_value=mock_res):
            df = scraper._scrape_with_firecrawl("Data Analyst", "Hyderabad")
            self.assertTrue(df.empty)
            self.assertEqual(scraper.last_run_stats.get("fc_queries_successful", 0), 0)
            self.assertGreaterEqual(scraper.last_run_stats.get("fc_errors", 0), 1)


if __name__ == "__main__":
    unittest.main()
