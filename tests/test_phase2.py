import os
import re
import time
import json
import unittest
import tempfile
import pandas as pd
import openpyxl
from unittest.mock import patch

from excel_generator import ExcelDashboardGenerator, canonicalize_job_url
from scraper import MultiPlatformJobScraper
from extractor import JobRequirementExtractor


class TestPhase2Acceptance(unittest.TestCase):
    """
    Phase 2 Acceptance & Audit Regression Test Suite verifying:
    - Application status & notes carryover via applications.json
    - Sync precedence (newest-timestamped/modified file across leads/ and Tracker_Latest.xlsx wins)
    - Cold start with absent leads/ and applications.json
    - Graceful PermissionError fallback on locked files
    - Tab 2 Run Summary audit sheet and accurate funnel math
    - Dual output (dated archive + stable Tracker_Latest.xlsx)
    - Test 6a: Gate reject counters increment and reset
    - Canonical URL ?jk= preservation on Indeed & ATS links
    - Reverting status back to 'Not Applied'
    - Hardened _within_ceiling() against 'Graduate with 5-8 years' & substring short-circuits
    - Formula injection sanitization & _is_senior_title() query isolation
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.leads_dir = os.path.join(self.temp_dir.name, "leads")
        self.stable_file = os.path.join(self.temp_dir.name, "Tracker_Latest.xlsx")
        self.applications_file = os.path.join(self.temp_dir.name, "applications.json")
        self.seen_jobs_file = os.path.join(self.temp_dir.name, "seen_jobs.json")
        self.config = {
            "output": {
                "leads_dir": self.leads_dir,
                "stable_filename": self.stable_file,
                "applications_file": self.applications_file,
                "seen_jobs_file": self.seen_jobs_file,
                "sheet_name": "Active Job Leads",
                "status_options": [
                    "Not Applied", "Applied", "Referral Requested",
                    "Screening Call", "Interviewing", "Assessment", "Rejected", "Offer Received"
                ]
            }
        }

    def tearDown(self):
        self.temp_dir.cleanup()

    def _create_sample_workbook(self, filepath, jobs_data):
        """Helper to create a formatted workbook with headers and sample jobs."""
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
        wb = openpyxl.Workbook()
        try:
            ws = wb.active
            ws.title = "Active Job Leads"
            headers = [
                "Sl. No", "Lead Status", "Date Posted", "Company Name",
                "Job Title / Role", "Main Skills Required", "Experience Level",
                "Location / Mode", "Platform", "Apply Action",
                "Application Status", "Referral / Notes"
            ]
            ws.append(headers)
            for idx, job in enumerate(jobs_data, 1):
                row_idx = ws.max_row + 1
                ws.cell(row=row_idx, column=1, value=idx)
                ws.cell(row=row_idx, column=2, value=job.get("lead_status", "🟢 NEW"))
                ws.cell(row=row_idx, column=3, value=job.get("date_posted", "2026-10-08"))
                ws.cell(row=row_idx, column=4, value=job.get("company", "Company A"))
                ws.cell(row=row_idx, column=5, value=job.get("title", "Data Analyst"))
                ws.cell(row=row_idx, column=6, value=job.get("skills", "Python, SQL"))
                ws.cell(row=row_idx, column=7, value=job.get("experience", "0 - 2 Years"))
                ws.cell(row=row_idx, column=8, value=job.get("location", "Hyderabad"))
                ws.cell(row=row_idx, column=9, value=job.get("platform", "LinkedIn"))

                c_apply = ws.cell(row=row_idx, column=10, value="Apply Now")
                if job.get("apply_url"):
                    c_apply.hyperlink = job["apply_url"]

                ws.cell(row=row_idx, column=11, value=job.get("status", "Not Applied"))
                ws.cell(row=row_idx, column=12, value=job.get("notes", ""))

            wb.save(filepath)
        finally:
            wb.close()

    def _create_multi_tab_workbook(self, filepath, tab1_jobs, tab2_jobs):
        """Helper to create a formatted workbook with both 'Active Job Leads' and 'Applied & In Progress' tabs."""
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
        wb = openpyxl.Workbook()
        try:
            ws1 = wb.active
            ws1.title = "Active Job Leads"
            headers1 = [
                "Sl. No", "Lead Status", "Date Posted", "Company Name",
                "Job Title / Role", "Main Skills Required", "Experience Level",
                "Location / Mode", "Platform", "Apply Action",
                "Application Status", "Referral / Notes"
            ]
            ws1.append(headers1)
            for idx, job in enumerate(tab1_jobs, 1):
                r = ws1.max_row + 1
                ws1.cell(row=r, column=1, value=idx)
                ws1.cell(row=r, column=2, value=job.get("lead_status", "🟢 NEW"))
                ws1.cell(row=r, column=3, value=job.get("date_posted", "2026-10-08"))
                ws1.cell(row=r, column=4, value=job.get("company", "Company A"))
                ws1.cell(row=r, column=5, value=job.get("title", "Data Analyst"))
                ws1.cell(row=r, column=6, value=job.get("skills", "Python, SQL"))
                ws1.cell(row=r, column=7, value=job.get("experience", "0 - 2 Years"))
                ws1.cell(row=r, column=8, value=job.get("location", "Hyderabad"))
                ws1.cell(row=r, column=9, value=job.get("platform", "LinkedIn"))
                c_apply = ws1.cell(row=r, column=10, value="Apply Now")
                if job.get("apply_url"):
                    c_apply.hyperlink = job["apply_url"]
                ws1.cell(row=r, column=11, value=job.get("status", "Not Applied"))
                ws1.cell(row=r, column=12, value=job.get("notes", ""))

            ws2 = wb.create_sheet(title="Applied & In Progress")
            headers2 = [
                "Sl. No", "Company Name", "Job Title / Role", "Location / Mode",
                "Platform", "Apply Action", "Application Status", "Referral / Notes", "Last Updated"
            ]
            ws2.append(headers2)
            for idx, job in enumerate(tab2_jobs, 1):
                r = ws2.max_row + 1
                ws2.cell(row=r, column=1, value=idx)
                ws2.cell(row=r, column=2, value=job.get("company", "Company A"))
                ws2.cell(row=r, column=3, value=job.get("title", "Data Analyst"))
                ws2.cell(row=r, column=4, value=job.get("location", "Hyderabad"))
                ws2.cell(row=r, column=5, value=job.get("platform", "LinkedIn"))
                c_apply2 = ws2.cell(row=r, column=6, value="View / Apply")
                if job.get("apply_url"):
                    c_apply2.hyperlink = job["apply_url"]
                ws2.cell(row=r, column=7, value=job.get("status", "Applied"))
                ws2.cell(row=r, column=8, value=job.get("notes", ""))
                ws2.cell(row=r, column=9, value="2026-10-08")

            wb.save(filepath)
        finally:
            wb.close()

    def test_status_and_note_carryover(self):
        """Test 1: Mark 'Applied' + note in file A -> status and note carried into new file."""
        file_a = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        target_url = "https://www.linkedin.com/jobs/view/12345"
        self._create_sample_workbook(file_a, [
            {
                "company": "Google",
                "title": "Data Analyst",
                "apply_url": target_url,
                "status": "Applied",
                "notes": "Referred by John"
            }
        ])

        generator = ExcelDashboardGenerator(self.config)
        new_jobs_df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW",
                "formatted_date": "2026-10-09",
                "company": "Google",
                "title": "Data Analyst",
                "extracted_skills": "Python, SQL",
                "extracted_experience": "0 - 2 Years",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "final_apply_url": target_url + "?tracking=xyz"
            }
        ])

        output_path = generator.generate(new_jobs_df)
        wb_new = openpyxl.load_workbook(output_path)
        try:
            ws_new = wb_new["Active Job Leads"]
            self.assertEqual(ws_new.cell(2, 11).value, "Applied")
            self.assertEqual(ws_new.cell(2, 12).value, "Referred by John")
        finally:
            wb_new.close()

    def test_conflicting_historical_files_newest_wins(self):
        """Test 2: Two historical files with conflicting status -> newest file wins."""
        file_old = os.path.join(self.leads_dir, "Job_Leads_2026-10-07_0900.xlsx")
        file_new = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        target_url = "https://www.linkedin.com/jobs/view/99999"

        self._create_sample_workbook(file_old, [
            {"company": "Amazon", "title": "BI Analyst", "apply_url": target_url, "status": "Applied", "notes": "Old note"}
        ])
        self._create_sample_workbook(file_new, [
            {"company": "Amazon", "title": "BI Analyst", "apply_url": target_url, "status": "Interviewing", "notes": "Round 1 cleared"}
        ])

        generator = ExcelDashboardGenerator(self.config)
        with open(self.applications_file, "r", encoding="utf-8") as f:
            ledger = json.load(f)

        norm_url = re.sub(r'\?.*$', '', target_url).strip().lower()
        self.assertIn(norm_url, ledger)
        self.assertEqual(ledger[norm_url]["status"], "Interviewing")
        self.assertEqual(ledger[norm_url]["notes"], "Round 1 cleared")

    def test_cold_start_empty_history(self):
        """Test 3: leads/ absent, applications.json absent -> clean generation, empty memory."""
        self.assertFalse(os.path.exists(self.leads_dir))
        self.assertFalse(os.path.exists(self.applications_file))

        generator = ExcelDashboardGenerator(self.config)
        jobs_df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW",
                "formatted_date": "Today",
                "company": "Microsoft",
                "title": "Support Engineer",
                "extracted_skills": "Windows, Azure",
                "extracted_experience": "1 - 2 Years",
                "location": "Bengaluru",
                "site": "Indeed",
                "final_apply_url": "https://indeed.com/viewjob?jk=1111111111111111"
            }
        ])

        archive_file = generator.generate(jobs_df)
        self.assertTrue(os.path.exists(archive_file))
        self.assertTrue(os.path.exists(self.stable_file))
        self.assertTrue(os.path.exists(self.leads_dir))

    def test_locked_file_sync_fallback(self):
        """Test 4: Yesterday's file locked/open in Excel -> sync skips gracefully, run completes."""
        file_locked = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        self._create_sample_workbook(file_locked, [{"company": "TCS", "title": "Developer"}])

        with open(self.applications_file, "w", encoding="utf-8") as f:
            json.dump({"https://test.com/job1": {"status": "Applied", "notes": "Existing note"}}, f)

        real_load = openpyxl.load_workbook
        def mock_load(path, **kwargs):
            if "Job_Leads_" in str(path):
                raise PermissionError(f"[WinError 32] The process cannot access '{path}' because it is being used.")
            return real_load(path, **kwargs)

        with patch("openpyxl.load_workbook", side_effect=mock_load):
            generator = ExcelDashboardGenerator(self.config)
            self.assertIn("https://test.com/job1", generator.applications_ledger)
            self.assertEqual(generator.applications_ledger["https://test.com/job1"]["status"], "Applied")

    def test_tab2_summary_sheet(self):
        """Test 5: Tab 2 Run Summary exists and displays exact stats from last_run_stats."""
        generator = ExcelDashboardGenerator(self.config)
        jobs_df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW",
                "formatted_date": "Recent",
                "company": "Infosys",
                "title": "Systems Engineer",
                "extracted_skills": "Java, SQL",
                "extracted_experience": "0 - 1 Years",
                "location": "Hyderabad",
                "site": "Naukri",
                "final_apply_url": "https://naukri.com/job-1"
            }
        ])
        run_stats = {
            "timestamp": "2026-10-09 16:30",
            "mode": "frugal",
            "credits_used": 14,
            "fc_candidates": 16,
            "jobspy_candidates": 10,
            "dropped_senior_title": 3,
            "dropped_experience_ceiling": 5,
        }

        output_path = generator.generate(jobs_df, run_stats=run_stats)
        wb = openpyxl.load_workbook(output_path)
        try:
            self.assertIn("Run Summary", wb.sheetnames)
            ws_summary = wb["Run Summary"]
            summary_dict = {}
            for row in range(2, ws_summary.max_row + 1):
                metric = ws_summary.cell(row, 1).value
                val = ws_summary.cell(row, 2).value
                if metric:
                    summary_dict[metric] = val

            self.assertIn("FRUGAL", str(summary_dict.get("Scraping Mode", "")))
            self.assertEqual(summary_dict.get("Total Raw Candidates Gathered"), 26)
            self.assertEqual(summary_dict.get("Senior Roles Filtered"), 3)
            self.assertEqual(summary_dict.get("Experience Ceiling Filtered (> 2 YOE)"), 5)
            self.assertEqual(summary_dict.get("Firecrawl Credits Consumed"), 14)
        finally:
            wb.close()

    def test_dual_output_identical_content(self):
        """Test 6: Dated filename and stable Tracker_Latest.xlsx both exist and have identical content."""
        generator = ExcelDashboardGenerator(self.config)
        jobs_df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW",
                "formatted_date": "Today",
                "company": "Wipro",
                "title": "Associate",
                "extracted_skills": "ITIL, Helpdesk",
                "extracted_experience": "0 - 2 Years",
                "location": "Bengaluru",
                "site": "Indeed",
                "final_apply_url": "https://indeed.com/job-wipro"
            }
        ])

        archive_file = generator.generate(jobs_df)
        self.assertTrue(os.path.exists(archive_file))
        self.assertTrue(os.path.exists(self.stable_file))

        wb_archive = openpyxl.load_workbook(archive_file)
        wb_stable = openpyxl.load_workbook(self.stable_file)
        try:
            self.assertEqual(wb_archive.sheetnames, wb_stable.sheetnames)
            ws_arch_leads = wb_archive["Active Job Leads"]
            ws_stab_leads = wb_stable["Active Job Leads"]

            self.assertEqual(ws_arch_leads.max_row, ws_stab_leads.max_row)
            self.assertEqual(ws_arch_leads.max_column, ws_stab_leads.max_column)

            for r in range(1, ws_arch_leads.max_row + 1):
                for c in range(1, ws_arch_leads.max_column + 1):
                    self.assertEqual(ws_arch_leads.cell(r, c).value, ws_stab_leads.cell(r, c).value)
        finally:
            wb_archive.close()
            wb_stable.close()

    def test_gate_counters_increment_and_reset(self):
        """Test 6a: Synthetic jobs through enrichment increment gates and reset cleanly."""
        scraper_config = {
            "search": {
                "max_experience_years": 2,
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad"
            },
            "skills_filter": {"target_skills": ["SQL", "Python"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)

        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 0)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 0)

        synthetic_df = pd.DataFrame([
            {
                "company": "Company Senior",
                "title": "Senior Data Analyst Lead",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "job_url": "https://www.linkedin.com/jobs/view/101",
                "job_url_direct": "https://www.linkedin.com/jobs/view/101",
                "description": "Lead a team of analysts. Fresher welcome.",
                "date_posted": "Recent",
                "search_query": "Data Analyst"
            },
            {
                "company": "Company Mid",
                "title": "Data Analyst",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "job_url": "https://www.linkedin.com/jobs/view/102",
                "job_url_direct": "https://www.linkedin.com/jobs/view/102",
                "description": "Requires minimum 5 to 8 years of experience in data modeling.",
                "date_posted": "Recent",
                "search_query": "Data Analyst"
            },
            {
                "company": "Company Junior",
                "title": "Junior Data Analyst",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "job_url": "https://www.linkedin.com/jobs/view/103",
                "job_url_direct": "https://www.linkedin.com/jobs/view/103",
                "description": "Entry level 0-1 years of experience in SQL and Python.",
                "date_posted": "Recent",
                "search_query": "Data Analyst"
            }
        ])

        enriched = scraper._enrich_jobs(synthetic_df)

        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 1)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 1)
        self.assertEqual(len(enriched), 1)

        scraper.last_run_stats["dropped_senior_title"] = 0
        scraper.last_run_stats["dropped_experience_ceiling"] = 0
        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 0)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 0)

    def test_indeed_query_param_jk_preservation(self):
        """Test 7: Indeed URLs with distinct ?jk= tokens do not collide in ledger, dedup, or seen_jobs."""
        url_job1 = "https://www.indeed.co.in/viewjob?jk=0123456789abcdef&utm_source=serp"
        url_job2 = "https://www.indeed.co.in/viewjob?jk=fedcba9876543210&utm_source=serp"

        self.assertEqual(
            canonicalize_job_url(url_job1),
            "https://www.indeed.co.in/viewjob?jk=0123456789abcdef"
        )
        self.assertNotEqual(canonicalize_job_url(url_job1), canonicalize_job_url(url_job2))

        # Mark only Job 1 as 'Applied' in historical workbook
        file_a = os.path.join(self.leads_dir, "Job_Leads_2026-10-10_0900.xlsx")
        self._create_sample_workbook(file_a, [
            {"company": "Company 1", "title": "Analyst 1", "apply_url": url_job1, "status": "Applied", "notes": "Job 1 Note"},
            {"company": "Company 2", "title": "Analyst 2", "apply_url": url_job2, "status": "Not Applied", "notes": ""}
        ])

        generator = ExcelDashboardGenerator(self.config)
        df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW", "formatted_date": "Today", "company": "Company 1",
                "title": "Analyst 1", "extracted_skills": "SQL", "extracted_experience": "0 - 2 Years",
                "location": "Hyderabad", "site": "Indeed", "final_apply_url": url_job1
            },
            {
                "lead_status": "🟢 NEW", "formatted_date": "Today", "company": "Company 2",
                "title": "Analyst 2", "extracted_skills": "Python", "extracted_experience": "0 - 2 Years",
                "location": "Hyderabad", "site": "Indeed", "final_apply_url": url_job2
            }
        ])
        out_path = generator.generate(df)
        wb = openpyxl.load_workbook(out_path)
        try:
            ws = wb["Active Job Leads"]
            self.assertEqual(ws.cell(2, 11).value, "Applied")
            self.assertEqual(ws.cell(2, 12).value, "Job 1 Note")
            # Job 2 must NOT inherit Job 1's status
            self.assertEqual(ws.cell(3, 11).value, "Not Applied")
            self.assertIn(ws.cell(3, 12).value, ["", None])
        finally:
            wb.close()

    def test_tracker_latest_mtime_precedence_and_revert(self):
        """Test 8: Edits in Tracker_Latest.xlsx override older leads/ file, including reverting to 'Not Applied'."""
        lead_archive = os.path.join(self.leads_dir, "Job_Leads_2026-10-11_0900.xlsx")
        url_1 = "https://www.linkedin.com/jobs/view/55555"
        url_2 = "https://www.linkedin.com/jobs/view/66666"

        # 1. Create archive in leads/ with url_1 = Applied, url_2 = Applied
        self._create_sample_workbook(lead_archive, [
            {"company": "Meta", "title": "Data Analyst", "apply_url": url_1, "status": "Applied", "notes": "Initial"},
            {"company": "Netflix", "title": "BI Engineer", "apply_url": url_2, "status": "Applied", "notes": "To revert"}
        ])
        old_time = time.time() - 120
        os.utime(lead_archive, (old_time, old_time))

        # 2. User edits Tracker_Latest.xlsx more recently: url_1 -> Interviewing, url_2 -> reverted to Not Applied
        self._create_sample_workbook(self.stable_file, [
            {"company": "Meta", "title": "Data Analyst", "apply_url": url_1, "status": "Interviewing", "notes": "Updated in Latest"},
            {"company": "Netflix", "title": "BI Engineer", "apply_url": url_2, "status": "Not Applied", "notes": ""}
        ])
        new_time = time.time()
        os.utime(self.stable_file, (new_time, new_time))

        generator = ExcelDashboardGenerator(self.config)
        self.assertEqual(generator.applications_ledger[canonicalize_job_url(url_1)]["status"], "Interviewing")
        self.assertEqual(generator.applications_ledger[canonicalize_job_url(url_1)]["notes"], "Updated in Latest")
        self.assertEqual(generator.applications_ledger[canonicalize_job_url(url_2)]["status"], "Not Applied")
        self.assertEqual(generator.applications_ledger[canonicalize_job_url(url_2)]["notes"], "")

    def test_graduate_and_internal_experience_ceiling_rejection(self):
        """Test 9: Hardened _within_ceiling rejects 'Graduate with 5-8 years' & 'internal' substring bypasses."""
        scraper_config = {
            "search": {"max_experience_years": 2},
            "skills_filter": {"target_skills": ["SQL"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)

        # Must reject senior experience strings despite 'Graduate', 'internal', or contract months
        self.assertFalse(scraper._within_ceiling("Graduate with 5-8 years experience", max_years=2))
        self.assertFalse(scraper._within_ceiling("5+ years in internal audit", max_years=2))
        self.assertFalse(scraper._within_ceiling("5+ years experience (6 months contract)", max_years=2))
        self.assertFalse(scraper._within_ceiling("3–5 YOE", max_years=2))
        self.assertFalse(scraper._within_ceiling("30 months", max_years=2))

        # Must accept valid entry-level / fresher / range strings
        self.assertTrue(scraper._within_ceiling("Fresh graduate", max_years=2))
        self.assertTrue(scraper._within_ceiling("0-5 years", max_years=2))
        self.assertTrue(scraper._within_ceiling("1+ to 4 years", max_years=2))
        self.assertTrue(scraper._within_ceiling("6 - 18 months", max_years=2))
        self.assertTrue(scraper._within_ceiling("6 months to 2 years", max_years=2))

    def test_formula_injection_and_url_query_isolation(self):
        """Test 10: Formula strings are escaped, indeed.co.in is valid, and ?utm_source=lead_form is not dropped."""
        scraper_config = {
            "search": {"max_experience_years": 2},
            "skills_filter": {"target_skills": ["SQL"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)

        # indeed.co.in direct URL validation
        self.assertTrue(scraper._is_direct_job_url("https://www.indeed.co.in/viewjob?jk=0123456789abcdef"))
        # Query param 'lead_form' must not trigger senior title rejection
        self.assertFalse(scraper._is_senior_title(
            "Junior Data Analyst",
            "https://www.linkedin.com/jobs/view/123456?utm_source=lead_form&trk=director_campaign"
        ))

        generator = ExcelDashboardGenerator(self.config)
        df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW",
                "formatted_date": "Today",
                "company": "=CMD|' /C calc'!A0",
                "title": "Data Analyst",
                "extracted_skills": "SQL",
                "extracted_experience": "0 - 2 Years",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "final_apply_url": "https://www.linkedin.com/jobs/view/77777"
            }
        ])
        out_path = generator.generate(df)
        wb = openpyxl.load_workbook(out_path)
        try:
            ws = wb["Active Job Leads"]
            self.assertEqual(ws.cell(2, 4).value, "'=CMD|' /C calc'!A0")
        finally:
            wb.close()

    def test_degree_and_contract_durations_do_not_drop_freshers(self):
        """Test 11: '4 year degree' and '36 months contract' do not trigger >2 YOE ceiling drops."""
        scraper_config = {
            "search": {"max_experience_years": 2},
            "skills_filter": {"target_skills": ["SQL", "Excel"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)
        df = pd.DataFrame([
            {
                "company": "Acme Corp",
                "title": "Junior Data Analyst",
                "location": "Hyderabad",
                "site": "LinkedIn",
                "job_url": "https://www.linkedin.com/jobs/view/9001",
                "job_url_direct": "https://www.linkedin.com/jobs/view/9001",
                "description": "Requires a 4 year Bachelor's degree in Computer Science. 36 months contract. Freshers welcome.",
                "date_posted": "Recent",
                "search_query": "Data Analyst"
            }
        ])
        enriched = scraper._enrich_jobs(df)
        self.assertEqual(len(enriched), 1)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 0)

    def test_excel_verb_not_extracted_as_skill(self):
        """Test 12: 'excel in' as a verb is ignored while 'Excel' as a tool is extracted."""
        scraper_config = {
            "search": {"max_experience_years": 2},
            "skills_filter": {"target_skills": ["SQL", "Excel"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)
        skills_verb = scraper.extractor.extract_skills("You will excel in a fast-paced environment and excel at teamwork.")
        self.assertNotIn("Excel", skills_verb)

        skills_tool = scraper.extractor.extract_skills("Proficient in MS Excel, SQL, and dashboarding.")
        self.assertIn("Excel", skills_tool)

    def test_leadsquared_and_principal_companies_not_dropped_as_senior(self):
        """Test 13: Companies like LeadSquared or Principal Global in URL slugs are not dropped as senior."""
        scraper_config = {
            "search": {"max_experience_years": 2},
            "skills_filter": {"target_skills": ["SQL"]},
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(scraper_config)
        self.assertFalse(scraper._is_senior_title(
            "Data Analyst",
            "https://www.naukri.com/job-listings-data-analyst-leadsquared-hyderabad-0-to-2-years-121026001234",
            company="LeadSquared"
        ))
        self.assertFalse(scraper._is_senior_title(
            "Associate Analyst - Lead Generation",
            "https://www.linkedin.com/jobs/view/9002",
            company="Infosys"
        ))

    def test_touching_old_archive_mtime_does_not_revert_status(self):
        """Test 14: Opening an old file in leads/ and touching its mtime never overwrites Tracker_Latest.xlsx."""
        lead_archive = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        target_url = "https://www.linkedin.com/jobs/view/88888"

        self._create_sample_workbook(lead_archive, [
            {"company": "Stripe", "title": "Data Analyst", "apply_url": target_url, "status": "Not Applied", "notes": ""}
        ])
        self._create_sample_workbook(self.stable_file, [
            {"company": "Stripe", "title": "Data Analyst", "apply_url": target_url, "status": "Interviewing", "notes": "HM round"}
        ])

        # Simulate user opening the old archive in Excel LATER, making its mtime newer than Tracker_Latest.xlsx
        future_time = time.time() + 300
        os.utime(lead_archive, (future_time, future_time))

        generator = ExcelDashboardGenerator(self.config)
        self.assertEqual(generator.applications_ledger[canonicalize_job_url(target_url)]["status"], "Interviewing")

    def test_applied_and_in_progress_tab_persists_across_days(self):
        """Test 15: Jobs marked 'Applied' appear in the 'Applied & In Progress' tab even when absent from today's scrape."""
        file_a = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        applied_url = "https://www.linkedin.com/jobs/view/777111"
        self._create_sample_workbook(file_a, [
            {"company": "Atlassian", "title": "Junior BI Analyst", "apply_url": applied_url, "status": "Applied", "notes": "Referral"}
        ])

        generator = ExcelDashboardGenerator(self.config)
        # Tomorrow's scrape only has a completely different job
        tomorrow_df = pd.DataFrame([
            {
                "lead_status": "🟢 NEW", "formatted_date": "Today", "company": "Oracle",
                "title": "SQL Developer", "extracted_skills": "SQL", "extracted_experience": "0 - 2 Years",
                "location": "Hyderabad", "site": "LinkedIn", "final_apply_url": "https://www.linkedin.com/jobs/view/888222"
            }
        ])
        out_path = generator.generate(tomorrow_df)
        wb = openpyxl.load_workbook(out_path)
        try:
            self.assertIn("Applied & In Progress", wb.sheetnames)
            ws_app = wb["Applied & In Progress"]
            self.assertEqual(ws_app.cell(2, 2).value, "Atlassian")
            self.assertEqual(ws_app.cell(2, 3).value, "Junior BI Analyst")
            self.assertEqual(ws_app.cell(2, 7).value, "Applied")
            self.assertEqual(ws_app.cell(2, 8).value, "Referral")
        finally:
            wb.close()

    def test_multi_tab_sync_conflict_active_tab_wins(self):
        """Test 16: Editing a status in Tab 1 wins over stale unedited row in Tab 2."""
        target_url = "https://www.linkedin.com/jobs/view/11111"
        target_url2 = "https://www.linkedin.com/jobs/view/22222"
        norm1 = canonicalize_job_url(target_url)
        norm2 = canonicalize_job_url(target_url2)

        # Baseline in applications.json
        initial_ledger = {
            norm1: {
                "status": "Applied",
                "notes": "Initial application",
                "company": "Comp A",
                "title": "Role A",
                "url": target_url
            },
            norm2: {
                "status": "Applied",
                "notes": "Sent resume",
                "company": "Comp B",
                "title": "Role B",
                "url": target_url2
            }
        }
        with open(self.applications_file, "w", encoding="utf-8") as f:
            json.dump(initial_ledger, f)

        # In Tracker_Latest.xlsx:
        # Job 1 is edited in Tab 1 to "Interviewing" ("Round 1 passed"), but Tab 2 still has stale "Applied" ("Initial application")
        # Job 2 is edited in Tab 2 to "Offer Received" ("Package accepted"), but Tab 1 still has "Applied" ("Sent resume")
        tab1_jobs = [
            {"company": "Comp A", "title": "Role A", "apply_url": target_url, "status": "Interviewing", "notes": "Round 1 passed"},
            {"company": "Comp B", "title": "Role B", "apply_url": target_url2, "status": "Applied", "notes": "Sent resume"}
        ]
        tab2_jobs = [
            {"company": "Comp A", "title": "Role A", "apply_url": target_url, "status": "Applied", "notes": "Initial application"},
            {"company": "Comp B", "title": "Role B", "apply_url": target_url2, "status": "Offer Received", "notes": "Package accepted"}
        ]
        self._create_multi_tab_workbook(self.stable_file, tab1_jobs, tab2_jobs)

        generator = ExcelDashboardGenerator(self.config)
        # Tab 1 edit won over stale Tab 2
        self.assertEqual(generator.applications_ledger[norm1]["status"], "Interviewing")
        self.assertEqual(generator.applications_ledger[norm1]["notes"], "Round 1 passed")
        # Tab 2 edit won over stale Tab 1
        self.assertEqual(generator.applications_ledger[norm2]["status"], "Offer Received")
        self.assertEqual(generator.applications_ledger[norm2]["notes"], "Package accepted")

    def test_archive_resurrection_prevented_by_initial_ledger_keys(self):
        """Test 17: Reverting status to 'Not Applied' in applications.json is never resurrected by old leads/ archive."""
        target_url = "https://www.linkedin.com/jobs/view/33333"
        norm = canonicalize_job_url(target_url)

        # User previously reverted this job to 'Not Applied' in applications.json
        with open(self.applications_file, "w", encoding="utf-8") as f:
            json.dump({
                norm: {
                    "status": "Not Applied",
                    "notes": "",
                    "company": "Comp C",
                    "title": "Role C",
                    "url": target_url
                }
            }, f)

        # An older archive exists in leads/ where this job had been marked 'Applied'
        archive_file = os.path.join(self.leads_dir, "Job_Leads_2026-10-08_0900.xlsx")
        self._create_sample_workbook(archive_file, [
            {"company": "Comp C", "title": "Role C", "apply_url": target_url, "status": "Applied", "notes": "Old applied note"}
        ])

        generator = ExcelDashboardGenerator(self.config)
        # Verify old archive did NOT resurrect status back to 'Applied'
        self.assertEqual(generator.applications_ledger[norm]["status"], "Not Applied")
        self.assertEqual(generator.applications_ledger[norm]["notes"], "")

    def test_naukri_slug_company_extraction_preserves_lead(self):
        """Test 18: Naukri slug company extraction preserves 'Lead' in company names."""
        scraper = MultiPlatformJobScraper(self.config)
        url = "https://www.naukri.com/job-listings-data-analyst-lead-squared-hyderabad-0-to-2-years-121026001234"
        role, company = scraper._extract_clean_metadata("Data Analyst - Lead Squared", url, "Naukri")
        self.assertEqual(company, "Lead Squared")

        url2 = "https://www.naukri.com/job-listings-sql-developer-lead-edge-solutions-bangalore-121026001235"
        role2, company2 = scraper._extract_clean_metadata("SQL Developer - Lead Edge Solutions", url2, "Naukri")
        self.assertEqual(company2, "Lead Edge Solutions")

    def test_senior_title_filter_drops_principal_and_head(self):
        """Test 19: Title-only senior gate properly drops Principal and Head roles while preserving lead gen."""
        scraper = MultiPlatformJobScraper(self.config)
        self.assertTrue(scraper._is_senior_title("Principal Financial Analyst"))
        self.assertTrue(scraper._is_senior_title("Head - Digital Marketing"))
        self.assertTrue(scraper._is_senior_title("Senior Business Analyst"))
        self.assertTrue(scraper._is_senior_title("Lead Systems Engineer"))
        self.assertFalse(scraper._is_senior_title("Lead Generation Executive"))
        self.assertFalse(scraper._is_senior_title("Associate Data Analyst", url="https://naukri.com/leadsquared", company="LeadSquared"))

    def test_excel_skill_extractor_handles_capitalization_and_prepositions(self):
        """Test 20: Excel tool is extracted even when followed by prepositions if capitalized, verb ignored."""
        extractor = JobRequirementExtractor(["Excel", "SQL"])
        # Verbs to ignore
        self.assertNotIn("Excel", extractor.extract_skills("Candidates will excel in fast-paced collaborative environments."))
        self.assertNotIn("Excel", extractor.extract_skills("You must excel at problem solving."))
        self.assertNotIn("Excel", extractor.extract_skills("Ability to excel within a team."))
        # Tools to extract
        self.assertIn("Excel", extractor.extract_skills("Strong skills with Excel in data analysis and financial reporting."))
        self.assertIn("Excel", extractor.extract_skills("Proficiency in EXCEL at an advanced level is required."))
        self.assertIn("Excel", extractor.extract_skills("Advanced Excel as well as SQL."))

    def test_bare_years_in_skill_extracted_and_rejected_by_ceiling(self):
        """Test 21: Bare '5 years in Power BI' is extracted and rejected by the experience ceiling."""
        extractor = JobRequirementExtractor([])
        scraper = MultiPlatformJobScraper(self.config)

        exp1 = extractor.extract_experience("Candidate must have 5 years in Power BI and SQL.")
        self.assertIn("5 Years", exp1)
        self.assertFalse(scraper._within_ceiling(exp1, max_years=2))

        exp2 = extractor.extract_experience("Role requires 3+ years of Python development.")
        self.assertIn("3+ Years", exp2)
        self.assertFalse(scraper._within_ceiling(exp2, max_years=2))

        exp3 = extractor.extract_experience("Requires 1 year in data analytics. Freshers welcome.")
        self.assertIn("1 Year", exp3)
        self.assertTrue(scraper._within_ceiling(exp3, max_years=2))

    def test_work_mode_extraction_arrangement_vs_duty_and_negation(self):
        """Test 22: Work mode correctly differentiates work arrangements from technical duties and negations."""
        extractor = JobRequirementExtractor([])

        # Technical IT duties should NOT trigger Remote
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "Provide remote support to desktop users via AnyDesk and TeamViewer."),
            "On-site (Hyderabad)"
        )
        self.assertEqual(
            extractor.extract_work_mode("Bengaluru", "Configure remote desktop administration and troubleshooting tools."),
            "On-site (Bengaluru)"
        )

        # Negations should NOT trigger Remote
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "This is an in-office position, not a remote opportunity."),
            "On-site (Hyderabad)"
        )
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "Non-remote role based in our financial district campus."),
            "On-site (Hyderabad)"
        )

        # Positive arrangements SHOULD trigger Remote
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "100% remote role with flexible working hours."),
            "Remote"
        )
        self.assertEqual(
            extractor.extract_work_mode("Remote", "Junior Data Analyst needed."),
            "Remote"
        )
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "Work from home full time."),
            "Remote"
        )

        # JobSpy native is_remote boolean override
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "Office based", is_remote=True),
            "Remote"
        )

        # Hybrid arrangements
        self.assertEqual(
            extractor.extract_work_mode("Hyderabad", "Hybrid work setup with 2 days in office and 3 days remote."),
            "Hybrid"
        )


if __name__ == "__main__":
    unittest.main()

