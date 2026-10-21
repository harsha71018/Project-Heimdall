import os
import csv
import hashlib
import tempfile
import unittest
from datetime import datetime, timedelta
import pandas as pd

from scraper import MultiPlatformJobScraper
from extractor import JobRequirementExtractor
from main import migrate_and_append_run_log, format_run_log_status, RUN_LOG_HEADERS


class TestURLAllowDeny(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        config = {
            "search": {
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad",
                "platforms": ["linkedin", "indeed"],
                "max_experience_years": 2
            },
            "modes": {
                "frugal": {"search_limit_per_role": 5, "deep_scrape": False}
            },
            "output": {
                "seen_jobs_file": os.path.join(self.temp_dir.name, "seen_jobs.json")
            }
        }
        self.scraper = MultiPlatformJobScraper(config, active_mode="frugal")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_positive_platforms_and_portals(self):
        valid_job_urls = [
            # Greenhouse
            "https://boards.greenhouse.io/acme/jobs/12345",
            # Lever
            "https://jobs.lever.co/acme/12345-abc-def",
            # Workday
            "https://acme.wd1.myworkdayjobs.com/en-US/Careers/job/Hyderabad/Data-Analyst_R12345",
            # SmartRecruiters (8+ digits)
            "https://jobs.smartrecruiters.com/AcmeCorp/12345678",
            # Taleo
            "https://acme.taleo.net/careersection/ex/jobdetail.ftl?job=12345",
            # BambooHR
            "https://acme.bamboohr.com/careers/123",
            "https://acme.bamboohr.com/jobs/view.php?id=123",
            # Ashby
            "https://jobs.ashbyhq.com/acme/12345-abc-def",
            # ICIMS
            "https://careers-acme.icims.com/jobs/12345/job",
            # Keka
            "https://acme.keka.com/careers/jobdetails/12345",
            # Workable
            "https://apply.workable.com/j/123456ABC",
            # ZipRecruiter / ZipApply
            "https://www.ziprecruiter.com/jobs//12345",
            "https://www.ziprecruiter.com/job/data-analyst/abc",
            "https://www.ziprecruiter.com/c/acme/job/data-analyst-123",
            "https://acme.zipapply.com/job/data-analyst-123",
            "https://acme.zipapply.com/apply/data-analyst-123",
            # LinkedIn
            "https://www.linkedin.com/jobs/view/1234567890",
            # Indeed (valid 16-hex jk token or direct slug)
            "https://www.indeed.com/viewjob?jk=0123456789abcdef",
            "https://www.indeed.com/job/data-analyst-0123456789abcdef",
            # Glassdoor
            "https://www.glassdoor.com/job-listing/data-analyst?jl=12345",
            # Naukri (requires 10-14 digit ID)
            "https://www.naukri.com/job-listings-data-analyst-acme-hyderabad-0-to-2-years-1234567890",
            # Enterprise Portals
            "https://careers.google.com/jobs/results/12345",
            "https://amazon.jobs/en/jobs/12345",
            "https://careers.microsoft.com/professionals/us/en/job/12345"
        ]

        for url in valid_job_urls:
            with self.subTest(url=url):
                self.assertTrue(self.scraper._is_direct_job_url(url), f"Expected direct job URL: {url}")

    def test_real_collected_urls_pass_validation(self):
        real_collected_urls = [
            # Real Naukri postings from seen_jobs ledger
            "https://www.naukri.com/job-listings-associate-data-analyst-talentas-technology-private-limited-hyderabad-1-to-2-years-041022500913",
            "https://www.naukri.com/job-listings-data-analyst-enterprise-relationship-mgmt-franklin-templeton-asset-management-india-pvt-hyderabad-0-to-4-years-050826505118",
            # Real LinkedIn postings from seen_jobs ledger
            "https://in.linkedin.com/jobs/view/data-analyst-at-ltimindtree-4397229737",
            "https://in.linkedin.com/jobs/view/data-analyst-at-cognizant-4399377250",
            # Real Indeed postings from seen_jobs ledger
            "http://in.indeed.com/job/quality-control-analyst-20fe4c9ec59827dd",
            "http://in.indeed.com/job/data-analyst-trainer-04d21293a5dcac78",
            # Real Workday postings from scraper trackers
            "https://crowe.wd12.myworkdayjobs.com/External_Careers/job/Hyderabad-Telangana-India/Power-Platform-Developer_R-51690-1",
            "https://nextracker.wd5.myworkdayjobs.com/en-US/nextpower_careers/job/Hyderabad-Telangana---IND/Senior-Supply-Chain-Analyst_NX-4740",
            "https://nextracker.wd5.myworkdayjobs.com/en-US/nextpower_careers/job/Hyderabad-Telangana---IND/Sourcing-Analyst_NX-4746",
            "https://statestreet.wd1.myworkdayjobs.com/en-US/Global/job/Hyderabad-India/Business-Analyst-Lead---Manager_R-798566",
            "https://wd1.myworkdaysite.com/recruiting/wf/WellsFargoJobs/job/Hyderabad-India/Quantitative-Analytics-Specialist_R-575935",
            "https://genpact.wd108.myworkdayjobs.com/External_Careers/job/1401-G-India-14-45-Uppal-LR-STPI-Hyderabad/Sr-Developer---Application-Development-S-4B_JR10027883",
            "https://aveva.wd3.myworkdayjobs.com/en-US/AVEVA_careers/job/Hyderabad-India/R-D-Member-of-Technical-Staff-II_R015273",
            "https://hitachi.wd1.myworkdayjobs.com/en-US/hitachi/job/Hyderabad-Telangana-India/Micro-Frontend-Engineer_R0134488",
            # Real Microsoft Careers posting from scraper trackers
            "https://apply.careers.microsoft.com/careers/job/1970393557004838?utm_source=indeed&domain=microsoft.com",
            # Real Taleo, Keka, and iCIMS postings from scraper trackers
            "https://uhg.taleo.net/careersection/10780/jobdetail.ftl?job=2391724&lang=en",
            "https://talentformula.keka.com/careers/jobdetails/142835?source=indeed%20job%20posts",
            "https://globalcareers-ice.icims.com/jobs/13669/job?utm_source=indeed_integration&iis=Job%20Board&iisn=Indeed&indeed-apply-token=dummy_redacted_token",
            # Real employer enterprise postings from tracker & ledger
            "https://accenture.com/in-en/careers/jobdetails?id=ATCI-5781704-S2070296_en",
            "https://accenture.com/in-en/careers/jobdetails?id=R00302548_en",
            "https://career.infosys.com/jobdesc?jobReferenceCode=INFSYS-INDEED1-252727&sourceId=4010",
            "https://usijobs.deloitte.com/en_US/careersUSI/JobDetail/USI-EH27-GCS-Growth-Analyst-L25/368701",
            "https://careers.wipro.com/job/BUSINESS-ANALYST-L1%28CONTRACT%29/205175-en_US",
            "https://careers.hcltech.com/job/Lead-Business-Analyst/172489-en_US",
            "https://apply.careers.hsbc.com/job/Hyderabad-Automation-Testing%2C-Functional-Testing%2C-Payments-Consultant-Specialist-TG-500032/1373312557/",
            "https://careers.micron.com/careers/job/44638923?utm_source=indeed&domain=micron.com",
            "https://jobs.trinet.com/careers/job/42900185",
            "https://pepsicojobs.com/jobs/461251/job?utm_source=indeed_integration&iis=Job%20Board&iisn=Indeed&indeed-apply-token=dummy_redacted_token",
            "https://careers.infor.com/postings/5d552481-e07f-4518-97f8-b1ff15d19489?utm_medium=job_board&utm_source=indeed",
            "https://fa-etvl-saasfaprod1.fa.ocs.oraclecloud.com/hcmUI/CandidateExperience/en/sites/CX_1/requisitions/preview/151532",
            "https://ness-usa.ttcportals.com/jobs/18279961-sr-net-full-stack-developer?tm_job=9382&tm_event=view&tm_company=48372",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp?page=jobdetails&clid=21001&JobID=J0926-2029&BRID=1336393&sbdid=936"
        ]

        for url in real_collected_urls:
            with self.subTest(url=url):
                self.assertTrue(self.scraper._is_direct_job_url(url), f"Expected real collected job URL to pass: {url}")

    def test_negative_unknown_hosts_and_lookalikes(self):
        denied_urls = [
            # Arbitrary / lookalike hosts
            "https://evil.com/job/12345",
            "https://fakegreenhouse.io/acme/jobs/123",
            "https://notzipapply.com/job/123",
            "https://greenhouse.io.attacker.com/jobs/123",
            "https://careers.evil.com/jobs/123",
            "https://evilcareers.google.com/jobs/results/123",
            "https://myworkdayjobs.com.attacker.com/job/123",
            "https://fake-naukri.com/job-listings-data-analyst-acme-hyderabad-0-to-2-years-1234567890",

            # Ad click-tracking wrappers and malformed ledger URLs (intentional exclusions)
            "https://click.appcast.io/t/Atu6ulFeDXmeKUHUVkXwqjfDDTyCcs6frThJnvYIdGM=",
            "https://click.appcast.io/t/TP1mcunz3rA04IyEJtcf8_bh83wLPxg2hGN1Jyy1lno=",
            "https://in.indeed.com/viewjob",
            "https://uhg.taleo.net/careersection/10780/jobdetail.ftl",

            # Non-job paths and malformed/missing query parameters on enterprise portals
            "https://accenture.com/in-en/careers",
            "https://accenture.com/in-en/careers/jobdetails",
            "https://accenture.com/in-en/careers/jobdetails?id=",
            "https://accenture.com/in-en/careers/jobdetails?id=invalid!@#$",
            "https://accenture.com/in-en/careers/jobdetails?jobId=ATCI-123",
            "https://accenture.com/in-en/careers/jobdetails/extra?id=ATCI-123",
            "https://career.infosys.com/",
            "https://career.infosys.com/jobdesc",
            "https://career.infosys.com/jobdesc?jobReferenceCode=",
            "https://career.infosys.com/jobdesc?jobReferenceCode=invalid!*&",
            "https://career.infosys.com/jobdesc?id=INFSYS-123",
            "https://career.infosys.com/jobdesc/extra?jobReferenceCode=INFSYS-123",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp?page=jobdetails",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp?page=search&JobID=J123",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp?page=jobdetails&JobID=",
            "https://indeed.njoyn.com/corp/xweb/XWeb.asp?page=jobdetails&JobID=invalid!@#",
            "https://usijobs.deloitte.com/",
            "https://usijobs.deloitte.com/en_US/careersUSI/SearchJobs",
            "https://careers.wipro.com/",
            "https://careers.wipro.com/job",
            "https://careers.hcltech.com/",
            "https://apply.careers.hsbc.com/",
            "https://careers.micron.com/",
            "https://pepsicojobs.com/",
            "https://pepsicojobs.com/jobs",
            "https://careers.infor.com/",
            "https://careers.infor.com/postings/not-a-uuid",
            "https://fa-etvl-saasfaprod1.fa.ocs.oraclecloud.com/",
            "https://ness-usa.ttcportals.com/",

            # ZipRecruiter search and root URLs carrying lvk tracking tokens
            "https://www.ziprecruiter.com/candidate/search?q=analyst&lvk=abc123",
            "https://www.ziprecruiter.com/jobs?lvk=abc123",
            "https://www.ziprecruiter.com/?lvk=abc123",
            "https://www.ziprecruiter.com/jobs-search?lvk=abc123",

            # Non-job paths on allowed hosts
            "https://boards.greenhouse.io/acme",
            "https://boards.greenhouse.io/",
            "https://jobs.lever.co/acme",
            "https://acme.wd1.myworkdayjobs.com/en-US/Careers",
            "https://jobs.smartrecruiters.com/AcmeCorp",
            "https://acme.bamboohr.com/careers",
            "https://acme.bamboohr.com/about",
            "https://acme.zipapply.com/",
            "https://acme.zipapply.com/about",
            "https://acme.zipapply.com/c/acme",
            "https://www.indeed.com/jobs?q=data+analyst",
            "https://www.indeed.com/viewjob",
            "https://www.linkedin.com/jobs/search?keywords=data",
            "https://www.linkedin.com/jobs/view/",
            "https://www.naukri.com/data-analyst-jobs",
            "https://careers.google.com/",
            "https://careers.microsoft.com/",
            "https://amazon.jobs/",

            # Partial path matches and trailing garbage suffixes
            "https://boards.greenhouse.io/acme/jobs/123garbage",
            "https://boards.greenhouse.io/acme/jobs/12345/extra",
            "https://jobs.lever.co/acme/12345678-abcd-ef01/extra",
            "https://jobs.lever.co/acme/123",
            "https://acme.wd1.myworkdayjobs.com/en-US/Careers/job",
            "https://acme.wd1.myworkdayjobs.com/en-US/Careers/job/Hyderabad/Data-Analyst_R12345/extra",
            "https://jobs.smartrecruiters.com/AcmeCorp/12345678garbage",
            "https://jobs.smartrecruiters.com/AcmeCorp/12345",
            "https://jobs.ashbyhq.com/acme/12345-abc-def/extra/garbage",
            "https://jobs.ashbyhq.com/acme/short",
            "https://acme.taleo.net/careersection/ex/jobdetail.ftl",
            "https://acme.taleo.net/careersection/ex/jobdetail.ftl?job=invalid",
            "https://acme.taleo.net/careersection/ex/jobdetail.ftlgarbage?job=12345",
            "https://careers-acme.icims.com/jobs/123garbage/job",
            "https://careers-acme.icims.com/jobs/12345/job/extra",
            "https://acme.keka.com/careers/jobdetails",
            "https://acme.keka.com/careers/jobdetails/12345/extra",
            "https://apply.workable.com/j/123",
            "https://apply.workable.com/j/123456ABC/extra",
            "https://careers.google.com/jobs/results/12345garbage",
            "https://careers.google.com/jobs/results/12345/extra",
            "https://amazon.jobs/en/jobs/12345garbage",
            "https://amazon.jobs/en/jobs/12345/data-analyst/extra",
            "https://careers.microsoft.com/professionals/us/en/job/12345garbage",
            "https://www.linkedin.com/jobs/view/1234567890garbage",
            "https://www.indeed.com/jobs?q=data+analyst&jk=0123456789abcdef",
            "https://www.indeed.com/viewjob?jk=invalidjk",
            "https://www.naukri.com/job-listings-data-analyst-acme-hyderabad-0-to-2-years-1234567890garbage",
            "https://www.naukri.com/job-listings-data-analyst-acme-hyderabad-0-to-2-years-12345",
            "https://acme.bamboohr.com/careers/123garbage",
            "https://acme.bamboohr.com/jobs/view.php?id=invalid"
        ]

        for url in denied_urls:
            with self.subTest(url=url):
                self.assertFalse(self.scraper._is_direct_job_url(url), f"Expected URL to be denied: {url}")


class TestExperienceCeiling(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        config = {
            "max_experience_years": 2,
            "output": {
                "seen_jobs_file": os.path.join(self.temp_dir.name, "seen_jobs.json")
            }
        }
        self.scraper = MultiPlatformJobScraper(config)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_bidirectional_mixed_ranges(self):
        # Forward: months to years
        self.assertTrue(self.scraper._within_ceiling("6 months to 1 year", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("36 months to 4 years", max_years=2))

        # Reverse: years to months
        self.assertTrue(self.scraper._within_ceiling("1 year to 36 months", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("4 years to 48 months", max_years=2))

    def test_decimal_and_standard_ranges(self):
        self.assertTrue(self.scraper._within_ceiling("1.5 years", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("2.5 years", max_years=2))
        self.assertTrue(self.scraper._within_ceiling("0 - 2 years", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("3 - 5 years", max_years=2))
        self.assertTrue(self.scraper._within_ceiling("6 - 12 months", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("36 - 48 months", max_years=2))
        self.assertTrue(self.scraper._within_ceiling("Fresher", max_years=2))
        self.assertFalse(self.scraper._within_ceiling("5+ years", max_years=2))


class TestContextAwareExtractor(unittest.TestCase):
    def setUp(self):
        self.extractor = JobRequirementExtractor([])

    def test_company_heritage_ignored(self):
        desc = "Established in 1995 with 25+ years of excellence. We are looking for entry level engineers."
        exp = self.extractor.extract_experience(desc)
        self.assertNotIn("25+", exp)
        self.assertIn("Entry / Fresher", exp)

    def test_negative_fresher_phrase_ignored(self):
        desc = "5+ years experience required for this role. Freshers need not apply."
        exp = self.extractor.extract_experience(desc)
        self.assertIn("5+", exp)
        self.assertNotIn("Fresher", exp)

    def test_requirement_verb_priority(self):
        desc = "Our firm brings 20 years of expertise. Candidate must possess 1 to 2 years of experience."
        exp = self.extractor.extract_experience(desc)
        self.assertEqual(exp, "1 To 2 Years")

    def test_endash_and_yoe_extracted(self):
        desc_endash = "Role requires 5–8 years of hands-on experience in data warehousing."
        self.assertEqual(self.extractor.extract_experience(desc_endash), "5–8 Years")
        desc_yoe = "Looking for candidates with 4+ YOE in SQL and Python."
        self.assertEqual(self.extractor.extract_experience(desc_yoe), "4+ Yoe")


class TestNaukriDateHandling(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.seen_jobs_file = os.path.join(self.temp_dir.name, "seen_jobs.json")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_date_decoder_calendar_and_bounds(self):
        # Stale date from 2021
        url_stale = "https://www.naukri.com/job-listings-analyst-1-to-2-years-041021500967"
        d_stale = MultiPlatformJobScraper._naukri_posted_date(url_stale)
        self.assertIsNotNone(d_stale)
        self.assertEqual(d_stale.year, 2021)
        self.assertEqual(d_stale.month, 10)
        self.assertEqual(d_stale.day, 4)

        # Pre-2020 valid date (e.g. 2018)
        url_pre2020 = "https://www.naukri.com/job-listings-analyst-1-to-2-years-041018500967"
        d_pre2020 = MultiPlatformJobScraper._naukri_posted_date(url_pre2020)
        self.assertIsNotNone(d_pre2020)
        self.assertEqual(d_pre2020.year, 2018)

        # Invalid calendar day (32)
        url_invalid = "https://www.naukri.com/job-listings-analyst-1-to-2-years-321024500967"
        d_invalid = MultiPlatformJobScraper._naukri_posted_date(url_invalid)
        self.assertIsNone(d_invalid)

    def test_enrich_jobs_date_filtering_and_funnel(self):
        config = {
            "max_experience_years": 2,
            "output": {"seen_jobs_file": self.seen_jobs_file}
        }
        scraper = MultiPlatformJobScraper(config)

        today = datetime.now().date()
        fresh_date = (today - timedelta(days=2)).strftime("%d%m%y")
        future_date = (today + timedelta(days=5)).strftime("%d%m%y")
        stale_date = "010121"  # Jan 1, 2021

        data = [
            # 1. Fresh Naukri job -> should be kept
            {
                "title": "Data Analyst",
                "company": "Company A",
                "job_url_direct": f"https://www.naukri.com/job-listings-analyst-0-to-2-years-{fresh_date}123456",
                "job_url": None,
                "description": "SQL and Python skills needed."
            },
            # 2. Stale Naukri job -> should be dropped
            {
                "title": "Data Analyst",
                "company": "Company B",
                "job_url_direct": f"https://www.naukri.com/job-listings-analyst-0-to-2-years-{stale_date}123456",
                "job_url": None,
                "description": "SQL skills needed."
            },
            # 3. Future Naukri date -> should be dropped
            {
                "title": "Data Analyst",
                "company": "Company C",
                "job_url_direct": f"https://www.naukri.com/job-listings-analyst-0-to-2-years-{future_date}123456",
                "job_url": None,
                "description": "SQL skills needed."
            },
            # 4. Unparseable Naukri date -> retained with "Unverified (Date Parse Failed)"
            {
                "title": "Data Analyst",
                "company": "Company D",
                "job_url_direct": "https://www.naukri.com/job-listings-analyst-0-to-2-years-999999123456",
                "job_url": None,
                "description": "SQL skills needed."
            },
            # 5. Senior Title -> should be dropped
            {
                "title": "Senior Data Analyst",
                "company": "Company E",
                "job_url_direct": f"https://www.naukri.com/job-listings-sr-analyst-0-to-2-years-{fresh_date}123456",
                "job_url": None,
                "description": "SQL skills needed."
            },
            # 6. Experience ceiling exceeded (e.g. 5 years) -> should be dropped
            {
                "title": "Data Analyst",
                "company": "Company F",
                "job_url_direct": f"https://www.naukri.com/job-listings-analyst-5-to-8-years-{fresh_date}123456",
                "job_url": None,
                "description": "5+ years required."
            }
        ]

        df = pd.DataFrame(data)
        enriched = scraper._enrich_jobs(df)

        # Expected retained: 2 (fresh Naukri job + unparseable Naukri job)
        self.assertEqual(len(enriched), 2)
        self.assertEqual(scraper.last_run_stats["dropped_stale_naukri"], 1)
        self.assertEqual(scraper.last_run_stats["dropped_future_naukri"], 1)
        self.assertEqual(scraper.last_run_stats["unparseable_naukri_dates"], 1)
        self.assertEqual(scraper.last_run_stats["dropped_senior_title"], 1)
        self.assertEqual(scraper.last_run_stats["dropped_experience_ceiling"], 1)

        # Verify unparseable date got the unverified marker
        unverified_rows = enriched[enriched["company"] == "Company D"]
        self.assertEqual(unverified_rows.iloc[0]["formatted_date"], "Unverified (Date Parse Failed)")


class TestRunLogMigrationAndSafety(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()
        self.log_file = os.path.join(self.test_dir.name, "run_log.csv")

    def tearDown(self):
        self.test_dir.cleanup()

    def test_legacy_7_column_migration(self):
        headers_7 = ["timestamp", "mode", "credits_used", "fc_candidates", "fc_retained", "jobspy_retained", "total_unique_jobs"]
        sample_row = ["2026-03-01 10:00", "balanced", "10", "4", "2", "5", "7"]
        with open(self.log_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(headers_7)
            writer.writerow(sample_row)

        new_data = {
            "timestamp": "2026-03-05 12:00",
            "mode": "frugal",
            "credits_used": "5",
            "fc_candidates": "2",
            "fc_markdown_hits": "1",
            "fc_retained": "1",
            "jobspy_retained": "3",
            "total_unique_jobs": "4",
            "dropped_stale_naukri": "0",
            "dropped_future_naukri": "0",
            "unparseable_naukri_dates": "0",
            "dropped_senior_title": "1",
            "dropped_experience_ceiling": "0",
        }

        result = migrate_and_append_run_log(self.log_file, new_data)
        self.assertTrue(result)

        with open(self.log_file, "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))

        self.assertEqual(reader[0], RUN_LOG_HEADERS)
        self.assertEqual(len(reader), 3)  # header + migrated row + new row
        migrated = reader[1]
        self.assertEqual(migrated[4], "N/A")  # fc_markdown_hits
        self.assertEqual(migrated[9:], ["N/A", "N/A", "N/A", "N/A", "N/A"])

    def test_legacy_8_column_migration(self):
        headers_8 = ["timestamp", "mode", "credits_used", "fc_candidates", "fc_markdown_hits", "fc_retained", "jobspy_retained", "total_unique_jobs"]
        sample_row = ["2026-03-02 10:00", "balanced", "12", "5", "3", "3", "6", "9"]
        with open(self.log_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(headers_8)
            writer.writerow(sample_row)

        new_data = {
            "timestamp": "2026-03-05 12:00",
            "mode": "frugal",
            "credits_used": "0",
            "fc_candidates": "0",
            "fc_markdown_hits": "0",
            "fc_retained": "0",
            "jobspy_retained": "2",
            "total_unique_jobs": "2",
            "dropped_stale_naukri": "0",
            "dropped_future_naukri": "0",
            "unparseable_naukri_dates": "0",
            "dropped_senior_title": "0",
            "dropped_experience_ceiling": "0",
        }

        result = migrate_and_append_run_log(self.log_file, new_data)
        self.assertTrue(result)

        with open(self.log_file, "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))

        self.assertEqual(reader[0], RUN_LOG_HEADERS)
        self.assertEqual(len(reader), 3)
        migrated = reader[1]
        self.assertEqual(migrated[4], "3")  # preserved
        self.assertEqual(migrated[9:], ["N/A", "N/A", "N/A", "N/A", "N/A"])

    def test_unrecognized_schema_remains_byte_identical(self):
        unrecognized_content = b"unknown_header_a,unknown_header_b,extra_col\nval1,val2,val3\nval4,val5,val6\n"
        with open(self.log_file, "wb") as f:
            f.write(unrecognized_content)

        with open(self.log_file, "rb") as f:
            hash_before = hashlib.sha256(f.read()).hexdigest()

        new_data = {
            "timestamp": "2026-03-05 12:00",
            "mode": "frugal",
            "total_unique_jobs": "5"
        }

        result = migrate_and_append_run_log(self.log_file, new_data)
        self.assertFalse(result)

        with open(self.log_file, "rb") as f:
            content_after = f.read()
            hash_after = hashlib.sha256(content_after).hexdigest()

        self.assertEqual(hash_before, hash_after)
        self.assertEqual(content_after, unrecognized_content)

    def test_run_log_status_reporting_and_append_flow(self):
        # 1. New file creation returns True
        row1 = {h: "val1" for h in RUN_LOG_HEADERS}
        self.assertTrue(migrate_and_append_run_log(self.log_file, row1))

        # 2. Append to matching schema returns True
        row2 = {h: "val2" for h in RUN_LOG_HEADERS}
        self.assertTrue(migrate_and_append_run_log(self.log_file, row2))

        with open(self.log_file, "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))
        self.assertEqual(len(reader), 3)  # header + row1 + row2

        # 3. Non-writable or invalid path handles gracefully and returns False
        invalid_path = os.path.join(self.test_dir.name, "nonexistent_dir", "sub", "log.csv")
        self.assertFalse(migrate_and_append_run_log(invalid_path, row1))

        # 4. Formatted status messages reflect success and skipped/failure outcomes accurately
        success_msg = format_run_log_status(True, self.log_file)
        self.assertTrue(success_msg.startswith("[+] Run Log Updated"))
        self.assertIn(os.path.abspath(self.log_file), success_msg)

        failure_msg = format_run_log_status(False, self.log_file)
        self.assertTrue(failure_msg.startswith("[!] Run Log Update Skipped / Failed"))
        self.assertIn(os.path.abspath(self.log_file), failure_msg)


    def test_load_config_empty_file_returns_dict(self):
        """Empty 0-byte config.yaml must return {} instead of None to prevent AttributeError."""
        import tempfile
        from main import load_config
        with tempfile.NamedTemporaryFile("w", delete=False) as f:
            temp_path = f.name
        try:
            cfg = load_config(temp_path)
            self.assertIsInstance(cfg, dict)
            self.assertEqual(cfg, {})
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)


if __name__ == "__main__":
    unittest.main()
