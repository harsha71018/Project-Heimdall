"""
Dedicated Runner for IT Support & ServiceNow Roles (Hyderabad & Bengaluru)
Completely isolated from your primary Data Analyst pipeline.
Does NOT modify or touch your personal trackers, seen_jobs.json, applications.json, or run_log.csv.
"""

import os
import sys
import time
import argparse
from datetime import datetime
import yaml
import pandas as pd
from dotenv import load_dotenv

# Ensure UTF-8 output encoding on Windows consoles
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from scraper import MultiPlatformJobScraper
from excel_generator import ExcelDashboardGenerator
from main import migrate_and_append_run_log, format_run_log_status

# Ensure .env is loaded
load_dotenv()


class ITSupportScraper(MultiPlatformJobScraper):
    """
    Subclass of MultiPlatformJobScraper customized for IT Support:
    - Uses its own isolated seen_jobs_it_support.json ledger
    - Natively searches multiple target locations (e.g., Hyderabad + Bengaluru)
    """

    def __init__(self, config: dict, active_mode: str = "frugal"):
        # Ensure seen_jobs_file is isolated in output config BEFORE super().__init__() loads it
        search_cfg = config.get("search", {})
        output_cfg = config.setdefault("output", {})
        isolated_seen = (
            output_cfg.get("seen_jobs_file")
            or search_cfg.get("seen_jobs_file")
            or "seen_jobs_it_support.json"
        )
        output_cfg["seen_jobs_file"] = isolated_seen
        output_cfg.setdefault("leads_dir", "leads_it_support")
        output_cfg.setdefault("stable_filename", "Tracker_Latest_IT_Support.xlsx")
        output_cfg.setdefault("applications_file", "applications_it_support.json")

        super().__init__(config, active_mode=active_mode)

    def scrape_all_roles_multi_location(self, locations: list) -> pd.DataFrame:
        """Runs the search across each target location and aggregates results."""
        from jobspy import scrape_jobs

        job_titles = self.config.get("job_titles", ["Service Desk"])
        platforms = self.config.get("platforms", ["linkedin", "indeed"])
        results_wanted = self.config.get("results_wanted_per_role", 10)
        hours_old = self.config.get("hours_old", 72)
        country_indeed = self.config.get("indeed_country", "india")

        firecrawl_enabled = self.firecrawl_config.get("enabled", True)
        use_firecrawl = firecrawl_enabled and bool(self.firecrawl_api_key)

        # Reset counters
        self.last_run_stats["fc_candidates"] = 0
        self.last_run_stats["fc_markdown_hits"] = 0
        self.last_run_stats["fc_retained"] = 0
        self.last_run_stats["fc_errors"] = 0
        self.last_run_stats["jobspy_candidates"] = 0
        self.last_run_stats["jobspy_retained"] = 0
        self.last_run_stats["jobspy_errors"] = 0
        self.last_run_stats["queries_attempted"] = 0
        self.last_run_stats["dropped_stale_naukri"] = 0
        self.last_run_stats["dropped_future_naukri"] = 0
        self.last_run_stats["unparseable_naukri_dates"] = 0
        self.last_run_stats["dropped_senior_title"] = 0
        self.last_run_stats["dropped_experience_ceiling"] = 0

        all_results = []

        # 1. Firecrawl Search (Naukri) across all locations
        if use_firecrawl:
            print(f"\n[*] [Firecrawl Engine] Mode: {self.active_mode.upper()}...")
            for loc in locations:
                for title in job_titles:
                    fc_df = self._scrape_with_firecrawl(title, loc, limit=self.search_limit_per_role)
                    if not fc_df.empty:
                        fc_df["_source_engine"] = "firecrawl"
                        all_results.append(fc_df)
                    time.sleep(12.5)

        # 2. JobSpy Search (LinkedIn + Indeed) across all locations
        platforms_str = ", ".join(p.capitalize() for p in platforms)
        print(f"\n[*] [JobSpy Engine] Querying multi-board direct feeds ({platforms_str})...")
        for loc in locations:
            for title in job_titles:
                print(f"[*] [Scraping] Searching for '{title}' in '{loc}' across {platforms}...")
                try:
                    self.last_run_stats["queries_attempted"] += 1
                    jobs_df = scrape_jobs(
                        site_name=platforms,
                        search_term=title,
                        location=loc,
                        results_wanted=results_wanted,
                        hours_old=hours_old,
                        country_indeed=country_indeed,
                        fetch_description=True,
                        verbose=0
                    )

                    if jobs_df is not None and not jobs_df.empty:
                        self.last_run_stats["jobspy_candidates"] += len(jobs_df)
                        valid_mask = jobs_df.apply(self._row_has_valid_direct_url, axis=1)
                        clean_df = jobs_df[valid_mask].copy()
                        if not clean_df.empty:
                            clean_df["search_query"] = title
                            clean_df["_source_engine"] = "jobspy"
                            all_results.append(clean_df)
                            print(f"   -> Found {len(clean_df)} direct postings for '{title}' ({loc}).")
                    else:
                        print(f"   -> No listings found for '{title}' ({loc}).")

                except Exception as e:
                    self.last_run_stats["jobspy_errors"] += 1
                    print(f"   [!] Warning: Error while scraping '{title}' ({loc}): {str(e)}")
                    continue

                time.sleep(2)

        if not all_results:
            print("[!] No job results collected from any platform.")
            return pd.DataFrame()

        combined_df = pd.concat(all_results, ignore_index=True)
        print(f"\n[*] Total raw direct jobs collected: {len(combined_df)}")
        deduped_df = self._deduplicate_jobs(combined_df)
        print(f"[*] Total unique verified jobs after deduplication: {len(deduped_df)}")

        enriched_df = self._enrich_jobs(deduped_df)
        print(f"[*] Final candidates within experience ceiling (<= {self.max_experience_years} YOE): {len(enriched_df)}")

        if "_source_engine" in enriched_df.columns:
            self.last_run_stats["fc_retained"] = int((enriched_df["_source_engine"] == "firecrawl").sum())
            self.last_run_stats["jobspy_retained"] = int((enriched_df["_source_engine"] == "jobspy").sum())
            enriched_df.drop(columns=["_source_engine"], inplace=True)

        return enriched_df


def main():
    parser = argparse.ArgumentParser(description="IT Support & ServiceNow Dedicated Job Hunter")
    parser.add_argument("--mode", type=str, choices=["frugal", "balanced", "exhaustive"], help="Scraping mode")
    parser.add_argument("--config", type=str, default="config_it_support.yaml", help="Path to config file")
    parser.add_argument("--output", type=str, default=None, help="Custom output Excel file path")
    parser.add_argument("--dry-run", action="store_true", help="Simulate run without making network calls")
    args = parser.parse_args()

    if not os.path.exists(args.config):
        print(f"[!] Error: Configuration file '{args.config}' not found.")
        sys.exit(1)

    with open(args.config, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    active_mode = args.mode or config.get("default_mode", "frugal")
    scraper = ITSupportScraper(config, active_mode=active_mode)

    roles = config["search"].get("job_titles", [])
    locations = config["search"].get("locations", ["Hyderabad", "Bengaluru"])
    platforms = config["search"].get("platforms", ["linkedin", "indeed"])

    print("=" * 70)
    print("  [*] IT SUPPORT & SERVICENOW JOB DISCOVERY ENGINE")
    print("=" * 70)
    print(f"[+] ACTIVE MODE    : {active_mode.upper()}")
    print(f"[+] TARGET ROLES   : {', '.join(roles)}")
    print(f"[+] LOCATION(S)    : {', '.join(locations)}")
    print(f"[+] PLATFORMS      : {', '.join(platforms)}")
    print(f"[+] SEEN LEDGER    : {scraper.seen_jobs_file}")
    print("-" * 70)

    if args.dry_run:
        print("[*] DRY-RUN MODE: Configuration validated successfully. Exiting without network calls.")
        print("=" * 70)
        return

    start_credits = scraper.get_remaining_credits()
    if start_credits is not None:
        print(f"[*] Firecrawl Live Balance : {start_credits:,} credits available.")

    enriched_jobs_df = scraper.scrape_all_roles_multi_location(locations)

    end_credits = scraper.get_remaining_credits()
    if isinstance(start_credits, int) and isinstance(end_credits, int):
        credits_used = max(0, start_credits - end_credits)
        credit_audit_str = f"{start_credits:,} before -> {end_credits:,} after ({credits_used} credits used this run)"
    else:
        credits_used = "unknown"
        credit_audit_str = "Delta calculation unavailable"

    stats = scraper.last_run_stats
    total_unique = len(enriched_jobs_df)

    log_file = config.get("output", {}).get("run_log", "run_log_it_support.csv")
    row_data = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "mode": active_mode,
        "credits_used": credits_used,
        "fc_candidates": stats.get("fc_candidates", 0),
        "fc_markdown_hits": stats.get("fc_markdown_hits", 0),
        "fc_retained": stats.get("fc_retained", 0),
        "jobspy_candidates": stats.get("jobspy_candidates", 0),
        "jobspy_retained": stats.get("jobspy_retained", 0),
        "total_unique_jobs": total_unique,
        "dropped_stale_naukri": stats.get("dropped_stale_naukri", 0),
        "dropped_future_naukri": stats.get("dropped_future_naukri", 0),
        "unparseable_naukri_dates": stats.get("unparseable_naukri_dates", 0),
        "dropped_senior_title": stats.get("dropped_senior_title", 0),
        "dropped_experience_ceiling": stats.get("dropped_experience_ceiling", 0),
    }

    log_updated = migrate_and_append_run_log(log_file, row_data)
    log_status_str = format_run_log_status(log_updated, log_file)

    # Instantiate generator early so applications_it_support.json syncs even on 0-job runs
    output_filename = args.output if args.output else None
    dashboard_gen = ExcelDashboardGenerator(config)

    if enriched_jobs_df.empty:
        total_raw_candidates = stats.get("fc_candidates", 0) + stats.get("jobspy_candidates", 0)
        fc_errors = stats.get("fc_errors", 0)
        jobspy_errors = stats.get("jobspy_errors", 0)
        use_firecrawl = config.get("firecrawl", {}).get("enabled", True) and bool(os.getenv("FIRECRAWL_API_KEY"))
        all_sources_failed = (
            total_raw_candidates == 0 and
            jobspy_errors > 0 and
            (not use_firecrawl or fc_errors > 0)
        )

        print("\n" + "=" * 70)
        if all_sources_failed:
            print("[!] ERROR: All attempted IT Support job scraping sources failed due to network or API connection errors.")
            print("    No platforms could be reached. Please check your internet connection or credentials.")
        else:
            print("[!] No jobs were returned for IT Support roles matching criteria today.")
        print(f"[+] Firecrawl Credit Audit              : {credit_audit_str}")
        print(log_status_str)
        print("=" * 70)
        if all_sources_failed:
            sys.exit(1)
        sys.exit(0)

    generated_file = dashboard_gen.generate(enriched_jobs_df, output_path=output_filename, run_stats=row_data)

    # Commit seen jobs ledger only after Excel write succeeds
    scraper._save_seen_jobs()

    print("\n" + "=" * 70)
    print(f"[+] SUCCESS! Total Active Opportunities Aggregated: {total_unique}")
    print(f"[+] Firecrawl Credit Audit              : {credit_audit_str}")
    print(f"[+] Output File Created                 : {os.path.abspath(generated_file)}")
    print(log_status_str)
    print("=" * 70)


if __name__ == "__main__":
    main()
