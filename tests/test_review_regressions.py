"""Regression scenarios from the independent review of commit 379bc7d."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock

import openpyxl
import pandas as pd

from excel_generator import ExcelDashboardGenerator, SYNC_STAMP_SHEET
from extractor import JobRequirementExtractor
from scraper import MultiPlatformJobScraper
from url_utils import canonicalize_job_url
from run_lock import acquire_run_lock, release_run_lock
from ai_gatekeeper import AIGatekeeper


class ReviewRegressions(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config = {
            "search": {"job_titles": ["Data Analyst"], "platforms": ["linkedin"]},
            "output": {
                "filename": str(self.root / "Tracker.xlsx"),
                "stable_filename": str(self.root / "Tracker_Latest.xlsx"),
                "leads_dir": str(self.root / "leads"),
                "applications_file": str(self.root / "applications.json"),
                "seen_jobs_file": str(self.root / "seen_jobs.json"),
            },
        }
        self.url = "https://www.linkedin.com/jobs/view/1234567890"
        self.alias = "https://in.linkedin.com/jobs/view/data-analyst-at-acme-1234567890"

    def job(self, url=None, description="Minimum 1 year of SQL experience"):
        return {"company": "Acme", "title": "Data Analyst", "location": "Hyderabad",
                "job_url": url or self.url, "final_apply_url": url or self.url,
                "description": description, "platform": "LinkedIn"}

    def write_apps(self, ledger):
        path = self.root / "applications.json"
        path.write_text(json.dumps(ledger), encoding="utf-8")
        return path

    def make_editable_pair(self):
        self.write_apps({self.url: {"status": "Applied", "notes": "old note"}})
        gen = ExcelDashboardGenerator(self.config)
        primary = self.config["output"]["stable_filename"]
        archive = str(self.root / "leads" / "Job_Leads_2026-10-14_011000.xlsx")
        for path in (primary, archive):
            gen.generate(pd.DataFrame([self.job()]), output_path=path)
        return primary, archive

    @staticmethod
    def edit(path, timestamp, status=None, notes=None):
        wb = openpyxl.load_workbook(path)
        try:
            ws = wb["Active Job Leads"]
            if status is not None:
                ws.cell(2, 11, status)
            if notes is not None:
                ws.cell(2, 12, notes)
            wb.save(path)
        finally:
            wb.close()
        os.utime(path, (timestamp, timestamp))

    def test_notes_only_edit_does_not_roll_back_status_in_either_time_order(self):
        for primary_time, archive_time in ((100, 200), (200, 100)):
            with self.subTest(primary_time=primary_time):
                primary, archive = self.make_editable_pair()
                self.edit(primary, primary_time, status="Interviewing")
                self.edit(archive, archive_time, notes="Recruiter called")
                ledger = ExcelDashboardGenerator(self.config).applications_ledger
                self.assertEqual(ledger[self.url]["status"], "Interviewing")
                self.assertEqual(ledger[self.url]["notes"], "Recruiter called")
                again = ExcelDashboardGenerator(self.config).applications_ledger
                self.assertEqual(again[self.url], ledger[self.url])
                # Remove only these fixtures before the second subcase.
                Path(primary).unlink()
                Path(archive).unlink()

    def test_same_field_newer_edit_wins_and_explicit_clear_survives(self):
        primary, archive = self.make_editable_pair()
        self.edit(primary, 100, status="Interviewing")
        self.edit(archive, 200, status="Not Applied", notes="")
        record = ExcelDashboardGenerator(self.config).applications_ledger[self.url]
        self.assertEqual(record["status"], "Not Applied")
        self.assertEqual(record["notes"], "")

    def test_later_notes_save_does_not_replay_previously_consumed_status(self):
        primary, archive = self.make_editable_pair()
        self.edit(primary, 100, status="Interviewing")
        ExcelDashboardGenerator(self.config)
        self.edit(archive, 200, status="Rejected")
        self.assertEqual(ExcelDashboardGenerator(self.config).applications_ledger[self.url]["status"], "Rejected")
        self.edit(primary, 300, notes="Recruiter follow-up")
        record = ExcelDashboardGenerator(self.config).applications_ledger[self.url]
        self.assertEqual(record["status"], "Rejected")
        self.assertEqual(record["notes"], "Recruiter follow-up")

    def test_regenerated_workbook_starts_new_edit_baseline(self):
        primary, archive = self.make_editable_pair()
        self.edit(primary, 100, status="Interviewing")
        gen = ExcelDashboardGenerator(self.config)
        gen.generate(pd.DataFrame([self.job()]), output_path=primary)
        self.edit(primary, 300, status="Applied")
        self.assertEqual(ExcelDashboardGenerator(self.config).applications_ledger[self.url]["status"], "Applied")

    def test_native_ids_collapse_variants_without_collapsing_distinct_jobs(self):
        pairs = [
            (self.url, self.alias),
            ("https://www.naukri.com/job-listings-analyst-a-141126123456",
             "https://www.naukri.com/job-listings-analyst-b-141126123456?src=x"),
            ("https://in.indeed.com/rc/clk?jk=0123456789abcdef&from=serp",
             "https://www.indeed.co.in/viewjob?jk=0123456789abcdef"),
        ]
        for first, second in pairs:
            self.assertEqual(canonicalize_job_url(first), canonicalize_job_url(second))
        self.assertNotEqual(canonicalize_job_url(self.url), canonicalize_job_url(self.url[:-1] + "1"))
        self.assertNotEqual(canonicalize_job_url(self.url), canonicalize_job_url(self.url.replace("linkedin.com", "linkedin.com.evil.test")))
        scraper = MultiPlatformJobScraper(self.config)
        self.assertEqual(len(scraper._deduplicate_jobs(pd.DataFrame([self.job(), self.job(self.alias)]))), 1)

    def test_application_migration_preserves_status_notes_raw_link_and_backup(self):
        original = {self.alias: {"status": "Applied", "notes": "Referral: Sam", "url": self.alias}}
        path = self.write_apps(original)
        original_bytes = path.read_bytes()
        gen = ExcelDashboardGenerator(self.config)
        self.assertEqual(gen.applications_ledger[self.url]["status"], "Applied")
        self.assertEqual(gen.applications_ledger[self.url]["notes"], "Referral: Sam")
        self.assertEqual(gen.applications_ledger[self.url]["url"], self.alias)
        backups = list(self.root.glob("applications.json.identity-*.bak"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), original_bytes)
        output = gen.generate(pd.DataFrame([self.job(self.alias)]), output_path=str(self.root / "out.xlsx"))
        wb = openpyxl.load_workbook(output)
        try:
            self.assertEqual(wb["Active Job Leads"].cell(2, 11).value, "Applied")
            self.assertEqual(wb["Active Job Leads"].cell(2, 10).hyperlink.target, self.alias)
        finally:
            wb.close()
        ExcelDashboardGenerator(self.config)
        self.assertEqual(len(list(self.root.glob("applications.json.identity-*.bak"))), 1)

    def test_migration_collision_merges_field_versions_and_retains_source_records(self):
        self.write_apps({
            self.url: {"status": "Interviewing", "notes": "old", "_field_edited_at": {"status": 200, "notes": 100}},
            self.alias: {"status": "Applied", "notes": "new", "_field_edited_at": {"status": 100, "notes": 200}},
        })
        record = ExcelDashboardGenerator(self.config).applications_ledger[self.url]
        self.assertEqual((record["status"], record["notes"]), ("Interviewing", "new"))
        self.assertEqual(len(record["_identity_conflicts"]), 2)

    def test_migration_aborts_if_backup_cannot_be_written(self):
        path = self.write_apps({self.alias: {"status": "Applied"}})
        original = path.read_bytes()
        with patch("ledger_utils.shutil.copy2", side_effect=PermissionError("read only")):
            with self.assertRaises(PermissionError):
                ExcelDashboardGenerator(self.config)
        self.assertEqual(path.read_bytes(), original)

    def test_migration_does_not_resurrect_notes_cleared_in_newer_record(self):
        self.write_apps({
            self.url: {"status": "Not Applied", "notes": "", "_edited_at": 200},
            self.alias: {"status": "Applied", "notes": "old note", "_edited_at": 100},
        })
        record = ExcelDashboardGenerator(self.config).applications_ledger[self.url]
        self.assertEqual((record["status"], record["notes"]), ("Not Applied", ""))

    def test_old_sync_stamp_alias_still_detects_human_edits(self):
        primary, _ = self.make_editable_pair()
        wb = openpyxl.load_workbook(primary)
        try:
            for row in wb[SYNC_STAMP_SHEET].iter_rows(min_row=2):
                row[0].value = self.alias
            wb["Active Job Leads"].cell(2, 11, "Interviewing")
            wb.save(primary)
        finally:
            wb.close()
        self.assertEqual(ExcelDashboardGenerator(self.config).applications_ledger[self.url]["status"], "Interviewing")

    def test_legacy_duplicate_rows_keep_their_own_stamp_baselines(self):
        self.write_apps({self.url: {"status": "Interviewing", "notes": "", "_edited_at": 100}})
        wb = openpyxl.Workbook()
        try:
            ws = wb.active
            ws.title = "Active Job Leads"
            ws.append(["Company", "Title", "Apply Action", "Application Status", "Notes"])
            ws.append(["Acme", "Data Analyst", self.alias + "?utm_source=test", "Applied", ""])
            ws.append(["Acme", "Data Analyst", self.url, "Not Applied", ""])
            stamp = wb.create_sheet(SYNC_STAMP_SHEET)
            stamp.append(["key", "status", "notes"])
            stamp.append([self.alias, "Applied", ""])
            stamp.append([self.url, "Not Applied", ""])
            path = self.config["output"]["stable_filename"]
            wb.save(path)
        finally:
            wb.close()
        os.utime(path, (200, 200))
        record = ExcelDashboardGenerator(self.config).applications_ledger[self.url]
        self.assertEqual(record["status"], "Interviewing")

    def test_seen_migration_keeps_latest_date_and_backup(self):
        path = self.root / "seen_jobs.json"
        today = datetime.now().strftime("%Y-%m-%d")
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        path.write_text(json.dumps({self.alias: today, self.url: yesterday}), encoding="utf-8")
        original = path.read_bytes()
        scraper = MultiPlatformJobScraper(self.config)
        self.assertEqual(scraper.seen_jobs, {self.url: today})
        result = scraper._enrich_jobs(pd.DataFrame([self.job()]))
        self.assertIn("SEEN", result.iloc[0]["lead_status"])
        scraper.save_seen_jobs()
        self.assertEqual(json.loads(path.read_text()), {self.url: today})
        self.assertEqual(next(self.root.glob("seen_jobs.json.identity-*.bak")).read_bytes(), original)

    def test_closed_first_copy_cannot_suppress_open_second_copy(self):
        for closed_url in (self.url, f"https://www.naukri.com/job-listings-data-analyst-acme-0-to-2-years-{datetime.now():%d%m%y}123456"):
            scraper = MultiPlatformJobScraper(self.config)
            scraper.firecrawl_api_key = "test"
            closed = pd.DataFrame([self.job(closed_url, "No longer accepting applications")])
            opened = pd.DataFrame([self.job(self.alias)])
            with patch.object(scraper, "_scrape_with_firecrawl", return_value=closed), patch("scraper.scrape_jobs", return_value=opened):
                result = scraper.scrape_all_roles()
            self.assertEqual(len(result), 1)
            self.assertEqual(result.iloc[0]["final_apply_url"], self.alias)
            self.assertEqual(scraper.last_run_stats["jobspy_retained"], 1)
            self.assertEqual(scraper.last_run_stats["fc_retained"], 0)

    def test_all_mandatory_experience_requirements_are_checked(self):
        extractor = JobRequirementExtractor()
        for description in (
            "Minimum 1 year of SQL experience and minimum 5 years of overall professional experience.",
            "Experience: 1-2 years with SQL. Requires 5 years of professional experience.",
        ):
            evidence = extractor.extract_experience(description)
            self.assertIn("5", evidence)
            self.assertFalse(MultiPlatformJobScraper(self.config)._within_ceiling(evidence, 2, 3))

    def test_preferred_experience_is_not_a_mandatory_minimum(self):
        evidence = JobRequirementExtractor().extract_experience("Minimum 1 year of SQL experience. 5 years of experience preferred.")
        self.assertNotIn("5", evidence)
        self.assertTrue(MultiPlatformJobScraper(self.config)._within_ceiling(evidence, 2, 3))

    def test_slug_cannot_override_higher_description_requirement(self):
        url = f"https://www.naukri.com/job-listings-data-analyst-acme-0-to-2-years-{datetime.now():%d%m%y}123456"
        scraper = MultiPlatformJobScraper(self.config)
        result = scraper._enrich_jobs(pd.DataFrame([self.job(url, "Minimum 5 years of professional experience.")]))
        self.assertTrue(result.empty)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 1)
        self.assertFalse(scraper._within_ceiling("0-5 years", 2, 3))
        self.assertTrue(scraper._within_ceiling("1-3 years", 2, 3))

    def test_ai_receives_requirement_excerpt_beyond_first_400_characters(self):
        description = "About our team. " * 40 + "Minimum 1 year of SQL experience required."
        frame = MultiPlatformJobScraper(self.config)._enrich_jobs(pd.DataFrame([self.job(description=description)]))
        response = MagicMock(status_code=200)
        response.json.return_value = {"choices": [{"message": {"content": json.dumps({"decisions": [
            {"id": 0, "is_fresher_eligible": True, "is_domain_match": True, "fit_score": 8}
        ]})}}]}
        with patch.dict(os.environ, {"OPENROUTER_API_KEY": "test"}), patch("ai_gatekeeper.requests.post", return_value=response) as post:
            result = AIGatekeeper(self.config).filter_dataframe(frame)
        self.assertEqual(len(result), 1)
        prompt = post.call_args.kwargs["json"]["messages"][-1]["content"]
        payload = json.loads(prompt.split("\n", 1)[1])[0]
        self.assertNotIn("Minimum", payload["snippet"])
        self.assertIn("Minimum 1 year of SQL experience required.", " ".join(payload["experience_evidence"]))

    def test_process_lock_rejects_old_active_owner_and_recovers_after_crash(self):
        path = str(self.root / "heimdall.lock")
        repo = str(Path(__file__).resolve().parents[1])
        script = "import sys; sys.path.insert(0, sys.argv[1]); from run_lock import acquire_run_lock; print(acquire_run_lock(sys.argv[2]), flush=True); sys.stdin.read()"
        process = subprocess.Popen([sys.executable, "-B", "-c", script, repo, path],
                                   stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        try:
            self.assertEqual(process.stdout.readline().strip(), "True")
            os.utime(path, (1, 1))
            self.assertFalse(acquire_run_lock(path))
            release_run_lock(path)  # A failed contender must not release the owner.
            self.assertFalse(acquire_run_lock(path))
            process.kill()
            process.communicate(timeout=10)
            self.assertTrue(acquire_run_lock(path))
            self.assertFalse(acquire_run_lock(path))
            release_run_lock(path)
            self.assertTrue(acquire_run_lock(path))
        finally:
            release_run_lock(path)
            if process.poll() is None:
                process.kill()
                process.communicate(timeout=10)

    def test_lock_io_error_blocks_run(self):
        self.assertFalse(acquire_run_lock(str(self.root / "missing" / "heimdall.lock")))

    @patch("scraper.scrape_jobs")
    @patch("scraper.time.sleep")
    def test_jobspy_platform_error_increments_errors_and_does_not_count_as_successful(self, mock_sleep, mock_scrape):
        """When JobSpy reports platform errors via df.attrs, count as errors and do NOT increment success."""
        df_error = pd.DataFrame()
        df_error.attrs = {"errors": {"linkedin": "HTTP 429 Rate Limit"}}
        mock_scrape.return_value = df_error

        cfg = {
            "search": {"job_titles": ["Data Analyst"], "platforms": ["linkedin"]},
            "output": {"seen_jobs_file": str(self.root / "seen.json")},
            "firecrawl": {"enabled": False}
        }
        scraper = MultiPlatformJobScraper(cfg)
        result = scraper.scrape_all_roles()

        self.assertTrue(result.empty)
        self.assertEqual(scraper.last_run_stats["jobspy_errors"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_queries_successful"], 0)

    @patch("scraper.scrape_jobs")
    @patch("scraper.time.sleep")
    def test_jobspy_partial_platform_error_counts_both_error_and_success(self, mock_sleep, mock_scrape):
        """When 1 platform fails but another succeeds, record error and mark query successful."""
        df_partial = pd.DataFrame([{
            "title": "Data Analyst",
            "company": "Acme",
            "location": "Hyderabad",
            "job_url_direct": "https://www.linkedin.com/jobs/view/1234567890",
            "description": "SQL required 1 year",
            "site": "linkedin"
        }])
        df_partial.attrs = {"errors": {"indeed": "Cloudflare challenge"}}
        mock_scrape.return_value = df_partial

        cfg = {
            "search": {"job_titles": ["Data Analyst"], "platforms": ["linkedin", "indeed"]},
            "output": {"seen_jobs_file": str(self.root / "seen.json")},
            "firecrawl": {"enabled": False}
        }
        scraper = MultiPlatformJobScraper(cfg)
        result = scraper.scrape_all_roles()

        self.assertEqual(len(result), 1)
        self.assertEqual(scraper.last_run_stats["jobspy_errors"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_queries_successful"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_candidates"], 1)

    @patch("scraper.scrape_jobs")
    @patch("scraper.time.sleep")
    def test_jobspy_single_platform_partial_failure_retains_jobs_and_records_error(self, mock_sleep, mock_scrape):
        """When a single requested platform returns valid jobs alongside a pagination/rate-limit error,
        the valid jobs must NOT be discarded and the query must count as successful."""
        df_with_jobs = pd.DataFrame([{
            "title": "Data Analyst",
            "company": "Acme",
            "location": "Hyderabad",
            "job_url_direct": "https://www.linkedin.com/jobs/view/1234567890",
            "description": "SQL required 1 year",
            "site": "linkedin"
        }])
        # Simulate page 1 succeeded with 1 lead, but page 2 hit HTTP 429
        df_with_jobs.attrs = {"errors": {"linkedin": "HTTP 429 Rate Limit on page 2"}}
        mock_scrape.return_value = df_with_jobs

        cfg = {
            "search": {"job_titles": ["Data Analyst"], "platforms": ["linkedin"]},
            "output": {"seen_jobs_file": str(self.root / "seen.json")},
            "firecrawl": {"enabled": False}
        }
        scraper = MultiPlatformJobScraper(cfg)
        result = scraper.scrape_all_roles()

        self.assertEqual(len(result), 1, "Returned jobs must be preserved despite platform error attribute")
        self.assertEqual(scraper.last_run_stats["jobspy_candidates"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_errors"], 1)
        self.assertEqual(scraper.last_run_stats["jobspy_queries_successful"], 1)

    def test_extract_location_prefers_default_and_is_deterministic(self):
        """Verify _extract_location_from_metadata prefers default_loc if matched,
        and otherwise selects deterministically from sorted KNOWN_CITIES."""
        scraper = MultiPlatformJobScraper(self.config)

        # Multi-city slug containing pune, hyderabad, and bengaluru
        slug = "https://www.naukri.com/job-listings-data-analyst-acme-pune-hyderabad-bengaluru-0-to-2-years-1234567890"

        # 1. When default_loc='Hyderabad' is present in slug, it must be preferred
        loc = scraper._extract_location_from_metadata(slug, "Data Analyst opening", default_loc="Hyderabad")
        self.assertEqual(loc, "Hyderabad")

        # 2. When default_loc='Chennai' is NOT in slug, it must deterministically pick Bengaluru
        # ('bengaluru' is first alphabetically among bengaluru, hyderabad, pune)
        loc2 = scraper._extract_location_from_metadata(slug, "Data Analyst opening", default_loc="Chennai")
        self.assertEqual(loc2, "Bengaluru")

        # 3. When default_loc is generic 'India', it shouldn't match as a city and should pick Bengaluru
        loc3 = scraper._extract_location_from_metadata(slug, "Data Analyst opening", default_loc="India")
        self.assertEqual(loc3, "Bengaluru")

        # 4. When no known city matches anywhere, fall back to default_loc
        loc_fallback = scraper._extract_location_from_metadata("https://example.com/job/1", "text", default_loc="Mysore")
        self.assertEqual(loc_fallback, "Mysore")

    @patch("scraper.scrape_jobs")
    def test_dropped_closed_job_resets_across_runs(self, mock_scrape):
        """Verify dropped_closed_job counter resets to 0 at the start of scrape_all_roles."""
        mock_scrape.return_value = pd.DataFrame()
        cfg = {
            "search": {"job_titles": ["Data Analyst"], "platforms": ["linkedin"]},
            "output": {"seen_jobs_file": str(self.root / "seen.json")},
            "firecrawl": {"enabled": False}
        }
        scraper = MultiPlatformJobScraper(cfg)
        scraper.last_run_stats["dropped_closed_job"] = 12
        scraper.scrape_all_roles()
        self.assertEqual(scraper.last_run_stats["dropped_closed_job"], 0)

    @patch.dict(os.environ, {"OPENROUTER_API_KEY": "test-key-abc"}, clear=False)
    @patch("requests.post")
    def test_ai_gatekeeper_prompt_and_referer_dynamic(self, mock_post):
        """Verify AIGatekeeper reflects custom max_experience_years, location,
        correct project referer, and tracks last_dropped_count."""
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "choices": [{
                "message": {
                    "content": json.dumps({
                        "decisions": [
                            {"id": 0, "is_fresher_eligible": False, "is_domain_match": True, "fit_score": 4.0, "reason": "Requires 5 YOE"}
                        ]
                    })
                }
            }]
        }
        mock_post.return_value = mock_response

        cfg = {
            "ai_gatekeeper": {"enabled": True, "model": "test/model:free"},
            "search": {
                "job_titles": ["Data Engineer"],
                "max_experience_years": 4,
                "location": "Germany"
            }
        }
        gatekeeper = AIGatekeeper(cfg)
        df = pd.DataFrame([{
            "title": "Senior Data Engineer",
            "company": "Tech Corp",
            "extracted_experience": "5 Years",
            "location": "Berlin",
            "description": "5+ years required."
        }])

        filtered_df = gatekeeper.filter_dataframe(df)

        self.assertTrue(filtered_df.empty)
        self.assertEqual(gatekeeper.last_dropped_count, 1)

        # Inspect API call args
        self.assertTrue(mock_post.called)
        kwargs = mock_post.call_args[1]
        headers = kwargs["headers"]
        self.assertEqual(headers.get("HTTP-Referer"), "https://github.com/harsha71018/Project-Heimdall")

        payload = kwargs["json"]
        system_content = payload["messages"][0]["content"]
        self.assertIn("entry-level (0-4 YOE) staffing in Germany", system_content)
        self.assertIn("strictly <= 4 years of experience", system_content)
        self.assertIn("0-4 years", system_content)
        self.assertIn(">4 years", system_content)


if __name__ == "__main__":
    unittest.main()
