import unittest
from datetime import datetime
from extractor import JobRequirementExtractor
from scraper import MultiPlatformJobScraper

class TestExperienceAndNaukriDate(unittest.TestCase):
    def setUp(self):
        self.extractor = JobRequirementExtractor([])

    def test_assistant_manager_not_tagged_fresher(self):
        desc = "Assistant Manager position. 5+ years experience required. Freshers need not apply."
        exp = self.extractor.extract_experience(desc)
        self.assertIn("5+", exp)
        self.assertNotIn("Fresher", exp)

    def test_pure_fresher_tagged_fresher(self):
        desc = "Exciting opening for college graduates and freshers. No prior experience required."
        exp = self.extractor.extract_experience(desc)
        self.assertEqual(exp, "0 - 1 Yrs (Entry / Fresher)")

    def test_naukri_date_decoder(self):
        # 041021 = 04-10-21 (Oct 4, 2021)
        url = "https://www.naukri.com/job-listings-data-annotator-navayuga-infotech-pvt-ltd-hyderabad-1-to-6-years-041021500967"
        d = MultiPlatformJobScraper._naukri_posted_date(url)
        self.assertIsNotNone(d)
        self.assertEqual(d.year, 2021)
        self.assertEqual(d.month, 10)
        self.assertEqual(d.day, 4)

        # Confirm age is over 30 days
        age_days = (datetime.now() - d).days
        self.assertGreater(age_days, 30)

if __name__ == "__main__":
    unittest.main()
