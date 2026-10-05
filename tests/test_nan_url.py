import unittest
import pandas as pd
from scraper import MultiPlatformJobScraper

class TestNaNURLFilter(unittest.TestCase):
    def test_mixed_dataframe_url_filter(self):
        config = {
            "search": {
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad",
                "platforms": ["linkedin", "indeed"],
                "max_experience_years": 2
            },
            "modes": {
                "frugal": {"search_limit_per_role": 10, "deep_scrape": False}
            }
        }
        scraper = MultiPlatformJobScraper(config, active_mode="frugal")

        # Mixed DataFrame: Row 1 has direct URL; Row 2 has NaN direct URL and valid LinkedIn job_url
        mixed_data = [
            {
                "title": "Data Analyst",
                "company": "Company A",
                "location": "Hyderabad",
                "job_url_direct": "https://careers.google.com/jobs/results/12345",
                "job_url": None
            },
            {
                "title": "Data Analyst",
                "company": "Company B",
                "location": "Hyderabad",
                "job_url_direct": None,
                "job_url": "https://www.linkedin.com/jobs/view/4471309557"
            }
        ]
        df = pd.DataFrame(mixed_data)

        # In pandas, None in mixed text column evaluates to float('nan')
        self.assertTrue(pd.isna(df.iloc[1]["job_url_direct"]))

        # Both rows must survive the filter
        valid_mask = df.apply(scraper._row_has_valid_direct_url, axis=1)
        survived_df = df[valid_mask].copy()

        self.assertEqual(len(survived_df), 2, f"Expected 2 rows to survive, got {len(survived_df)}")

if __name__ == "__main__":
    unittest.main()
