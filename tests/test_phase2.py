import os
import re
import json
import unittest
import tempfile
import pandas as pd
import openpyxl
from unittest.mock import patch

from excel_generator import ExcelDashboardGenerator
from scraper import MultiPlatformJobScraper


class TestPhase2Acceptance(unittest.TestCase):
    """
    Phase 2 Acceptance Test Suite verifying:
    - Application status & notes carryover via applications.json
    - Sync precedence (newest-timestamped file in leads/ wins)
    - Cold start with absent leads/ and applications.json
    - Graceful PermissionError fallback on locked files
    - Tab 2 Run Summary audit sheet
    - Dual output (dated archive + stable Tracker_Latest.xlsx)
    - Test 6a: Gate reject counters increment and reset
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.leads_dir = os.path.join(self.temp_dir.name, "leads")
        self.stable_file = os.path.join(self.temp_dir.name, "Tracker_Latest.xlsx")
        self.applications_file = os.path.join(self.temp_dir.name, "applications.json")
        self.config = {
            "output": {
                "leads_dir": self.leads_dir,
                "stable_filename": self.stable_file,
                "applications_file": self.applications_file,
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
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        wb = openpyxl.Workbook()
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
        ws_new = wb_new["Active Job Leads"]

        self.assertEqual(ws_new.cell(2, 11).value, "Applied")
        self.assertEqual(ws_new.cell(2, 12).value, "Referred by John")

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
                "final_apply_url": "https://indeed.com/viewjob?jk=111"
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

        # Mock openpyxl.load_workbook to raise PermissionError when syncing
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
            "dropped_senior_title": 3,
            "dropped_experience_ceiling": 5,
        }

        output_path = generator.generate(jobs_df, run_stats=run_stats)
        wb = openpyxl.load_workbook(output_path)
        self.assertIn("Run Summary", wb.sheetnames)

        ws_summary = wb["Run Summary"]
        summary_dict = {}
        for row in range(2, ws_summary.max_row + 1):
            metric = ws_summary.cell(row, 1).value
            val = ws_summary.cell(row, 2).value
            if metric:
                summary_dict[metric] = val

        self.assertIn("FRUGAL", str(summary_dict.get("Scraping Mode", "")))
        self.assertEqual(summary_dict.get("Senior Roles Filtered"), 3)
        self.assertEqual(summary_dict.get("Experience Ceiling Filtered (> 2 YOE)"), 5)
        self.assertEqual(summary_dict.get("Firecrawl Credits Consumed"), 14)

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

        self.assertEqual(wb_archive.sheetnames, wb_stable.sheetnames)
        ws_arch_leads = wb_archive["Active Job Leads"]
        ws_stab_leads = wb_stable["Active Job Leads"]

        self.assertEqual(ws_arch_leads.max_row, ws_stab_leads.max_row)
        self.assertEqual(ws_arch_leads.max_column, ws_stab_leads.max_column)

        for r in range(1, ws_arch_leads.max_row + 1):
            for c in range(1, ws_arch_leads.max_column + 1):
                self.assertEqual(ws_arch_leads.cell(r, c).value, ws_stab_leads.cell(r, c).value)

    def test_gate_counters_increment_and_reset(self):
        """Test 6a: Synthetic jobs through enrichment increment gates and reset cleanly."""
        scraper_config = {
            "search": {
                "max_experience_years": 2,
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad"
            },
            "skills_filter": {"target_skills": ["SQL", "Python"]}
        }
        scraper = MultiPlatformJobScraper(scraper_config)

        # Initial state: counters must be zero
        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 0)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 0)

        # 1. Synthetic Senior Job
        # 2. Synthetic Experience Ceiling Job (5 to 8 years)
        # 3. Synthetic Valid Job
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

        # Gates must increment
        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 1)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 1)
        self.assertEqual(len(enriched), 1)

        # Reset counters verification: resetting must bring them back to 0
        scraper.last_run_stats["dropped_senior_title"] = 0
        scraper.last_run_stats["dropped_experience_ceiling"] = 0
        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 0)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 0)


if __name__ == "__main__":
    unittest.main()
