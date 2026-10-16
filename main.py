import os
import sys
import csv
import uuid
import yaml
import argparse
from datetime import datetime

# Ensure UTF-8 output encoding on Windows consoles
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from scraper import MultiPlatformJobScraper
from excel_generator import ExcelDashboardGenerator

RUN_LOG_HEADERS = [
    "timestamp", "mode", "credits_used", "fc_candidates", "fc_markdown_hits",
    "fc_retained", "jobspy_retained", "total_unique_jobs",
    "dropped_stale_naukri", "dropped_future_naukri", "unparseable_naukri_dates",
    "dropped_senior_title", "dropped_experience_ceiling"
]


def migrate_and_append_run_log(log_file: str, row_data: dict) -> bool:
    """
    Logs run analytics to run_log.csv with backward-compatible schema migration.
    - If file does not exist or is empty, creates it with RUN_LOG_HEADERS and writes the row.
    - If existing headers match RUN_LOG_HEADERS, appends the new row in append mode.
    - If existing headers match recognized 7-col or 8-col legacy schemas, performs an atomic
      migration by copying transformed rows to a process-scoped temporary file, appending
      the new row, and replacing the original file via os.replace.
    - If existing headers are unrecognized, preserves the file byte-for-byte without
      modifying or appending, prints a warning, and returns False.
    Returns True if the log was successfully updated or migrated; False otherwise.
    """
    file_exists = os.path.exists(log_file)
    if not file_exists:
        try:
            with open(log_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(RUN_LOG_HEADERS)
                writer.writerow([row_data.get(h, "") for h in RUN_LOG_HEADERS])
            return True
        except Exception as e:
            print(f"[!] Warning: Could not write to {log_file}: {e}")
            return False

    try:
        with open(log_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            first_line = next(reader, None)

        if first_line is None:
            with open(log_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(RUN_LOG_HEADERS)
                writer.writerow([row_data.get(h, "") for h in RUN_LOG_HEADERS])
            return True

        if first_line == RUN_LOG_HEADERS:
            with open(log_file, "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([row_data.get(h, "") for h in RUN_LOG_HEADERS])
            return True

        recognized_8 = [
            "timestamp", "mode", "credits_used", "fc_candidates", "fc_markdown_hits",
            "fc_retained", "jobspy_retained", "total_unique_jobs"
        ]
        recognized_7 = [
            "timestamp", "mode", "credits_used", "fc_candidates",
            "fc_retained", "jobspy_retained", "total_unique_jobs"
        ]

        if first_line != recognized_8 and first_line != recognized_7:
            print(f"[!] Warning: Unrecognized {log_file} schema. Leaving file untouched and skipping log append.")
            return False

        migrated_rows = []
        with open(log_file, "r", encoding="utf-8") as f:
            dict_reader = csv.DictReader(f)
            for r in dict_reader:
                new_row = [
                    r.get("timestamp", ""),
                    r.get("mode", ""),
                    r.get("credits_used", ""),
                    r.get("fc_candidates", ""),
                    r.get("fc_markdown_hits") if r.get("fc_markdown_hits") is not None else "N/A",
                    r.get("fc_retained", ""),
                    r.get("jobspy_retained", ""),
                    r.get("total_unique_jobs", ""),
                    "N/A",  # dropped_stale_naukri
                    "N/A",  # dropped_future_naukri
                    "N/A",  # unparseable_naukri_dates
                    "N/A",  # dropped_senior_title
                    "N/A",  # dropped_experience_ceiling
                ]
                migrated_rows.append(new_row)

        temp_file = f"{log_file}.{os.getpid()}_{uuid.uuid4().hex[:8]}.tmp"
        try:
            with open(temp_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(RUN_LOG_HEADERS)
                writer.writerows(migrated_rows)
                writer.writerow([row_data.get(h, "") for h in RUN_LOG_HEADERS])
            os.replace(temp_file, log_file)
            return True
        finally:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except OSError:
                    pass
    except Exception as e:
        print(f"[!] Warning: Could not migrate or append to {log_file}: {e}")
        return False


def format_run_log_status(log_updated: bool, log_file: str) -> str:
    """Formats the console status message reflecting whether run_log.csv was updated."""
    abs_path = os.path.abspath(log_file)
    if log_updated:
        return f"[+] Run Log Updated                     : {abs_path}"
    return f"[!] Run Log Update Skipped / Failed     : {abs_path}"


def load_config(config_path: str = "config.yaml") -> dict:
    """Loads configuration from YAML file."""
    if not os.path.exists(config_path):
        print(f"[!] Config file '{config_path}' not found, using default configuration.")
        return {
            "search": {
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad",
                "hours_old": 24,
                "results_wanted_per_role": 20,
                "platforms": ["linkedin", "indeed", "glassdoor"],
                "indeed_country": "india"
            },
            "default_mode": "balanced",
            "modes": {
                "balanced": {"search_limit_per_role": 4, "deep_scrape": True, "timeout": 60}
            },
            "output": {
                "filename": "Job_Applications_Tracker.xlsx",
                "sheet_name": "Active Job Leads"
            }
        }
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def parse_args():
    """Parses command-line arguments."""
    parser = argparse.ArgumentParser(description="Automated Multi-Platform Job Hunter & Excel Generator")
    parser.add_argument("--mode", type=str, choices=["frugal", "balanced", "exhaustive"], help="Scraping intensity mode")
    parser.add_argument("--role", type=str, help="Specific job role to search for (overrides config)")
    parser.add_argument("--location", type=str, help="Location to search for (overrides config)")
    parser.add_argument("--hours", type=int, help="Max job age in hours (e.g. 24, 72)")
    parser.add_argument("--output", type=str, help="Output Excel filename")
    parser.add_argument("--config", type=str, default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--dry-run", action="store_true", help="Validate configuration and simulate run without scraping")
    return parser.parse_args()


def main():
    args = parse_args()
    config = load_config(args.config)

    # Apply CLI overrides BEFORE initializing scraper
    config.setdefault("search", {})
    config.setdefault("output", {})
    if args.role:
        config["search"]["job_titles"] = [args.role]
    if args.location:
        config["search"]["location"] = args.location
    if args.hours:
        config["search"]["hours_old"] = args.hours
    if args.output:
        config["output"]["filename"] = args.output

    # Determine Active Mode (CLI flag > config default_mode > "balanced")
    active_mode = args.mode or config.get("default_mode", "balanced")

    # Initialize scraper with active mode and overridden config
    scraper = MultiPlatformJobScraper(config, active_mode=active_mode)

    roles = config["search"].get("job_titles", [])
    location = config["search"].get("location", "Hyderabad")
    platforms = config["search"].get("platforms", [])

    print("=" * 70)
    print("  [*] AUTOMATED JOB DISCOVERY & APPLICATION TRACKER ENGINE")
    print("=" * 70)
    print(f"[+] ACTIVE MODE    : {active_mode.upper()}")
    if scraper.deep_scrape:
        print(f"[+] SCRAPE DEPTH   : Full Markdown extraction on top {scraper.search_limit_per_role} listings per role")
    else:
        print(f"[+] SCRAPE DEPTH   : Fast SERP snippet search (no page scraping)")
    print(f"[+] TARGET ROLES   : {', '.join(roles)}")
    print(f"[+] LOCATION       : {location}")
    print(f"[+] PLATFORMS      : {', '.join(platforms)}")
    print("-" * 70)

    if args.dry_run:
        print("[*] DRY-RUN MODE: Configuration validated successfully. Exiting without network calls.")
        print("=" * 70)
        return

    # Pre-flight Credit Check
    start_credits = scraper.get_remaining_credits()
    if start_credits is not None:
        print(f"[*] Firecrawl Live Balance : {start_credits:,} credits available.")
        if active_mode == "exhaustive" and start_credits < 50:
            print("[!] Warning: Less than 50 credits remaining. Consider using 'balanced' or 'frugal' mode.")
    else:
        print("[*] Firecrawl Balance Check: No active key or endpoint unreachable (continuing gracefully).")

    # 1. Scrape and Enrich Jobs
    enriched_jobs_df = scraper.scrape_all_roles()

    # Post-flight Credit Check & Delta Computation
    end_credits = scraper.get_remaining_credits()
    if isinstance(start_credits, int) and isinstance(end_credits, int):
        credits_used = max(0, start_credits - end_credits)
        credit_audit_str = f"{start_credits:,} before -> {end_credits:,} after ({credits_used} credits used this run)"
    else:
        credits_used = "unknown"
        credit_audit_str = "Delta calculation unavailable"

    # Read per-engine and funnel stats
    stats = scraper.last_run_stats
    fc_candidates = stats.get("fc_candidates", 0)
    fc_markdown_hits = stats.get("fc_markdown_hits", 0)
    fc_retained = stats.get("fc_retained", 0)
    jobspy_candidates = stats.get("jobspy_candidates", 0)
    jobspy_retained = stats.get("jobspy_retained", 0)
    dropped_stale_naukri = stats.get("dropped_stale_naukri", 0)
    dropped_future_naukri = stats.get("dropped_future_naukri", 0)
    unparseable_naukri_dates = stats.get("unparseable_naukri_dates", 0)
    dropped_senior_title = stats.get("dropped_senior_title", 0)
    dropped_experience_ceiling = stats.get("dropped_experience_ceiling", 0)
    total_unique = len(enriched_jobs_df)

    log_file = "run_log.csv"
    row_data = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "mode": active_mode,
        "credits_used": credits_used,
        "fc_candidates": fc_candidates,
        "fc_markdown_hits": fc_markdown_hits,
        "fc_retained": fc_retained,
        "jobspy_candidates": jobspy_candidates,
        "jobspy_retained": jobspy_retained,
        "total_unique_jobs": total_unique,
        "dropped_stale_naukri": dropped_stale_naukri,
        "dropped_future_naukri": dropped_future_naukri,
        "unparseable_naukri_dates": unparseable_naukri_dates,
        "dropped_senior_title": dropped_senior_title,
        "dropped_experience_ceiling": dropped_experience_ceiling,
    }

    log_updated = migrate_and_append_run_log(log_file, row_data)
    log_status_str = format_run_log_status(log_updated, log_file)

    # Instantiate generator early so applications.json syncs even on 0-job runs
    output_filename = args.output if args.output else None
    dashboard_gen = ExcelDashboardGenerator(config)

    if enriched_jobs_df.empty:
        total_raw_candidates = fc_candidates + jobspy_candidates
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
            print("[!] ERROR: All attempted job scraping sources failed due to network or API connection errors.")
            print("    No platforms could be reached. Please check your internet connection or credentials.")
        else:
            print("[!] No jobs matching criteria were returned today (zero new leads or all filtered).")
        print(f"    - Funnel Filter Drops           : {dropped_senior_title} senior, {dropped_experience_ceiling} ceiling, {dropped_stale_naukri} stale Naukri, {dropped_future_naukri} future Naukri ({unparseable_naukri_dates} unparseable retained)")
        print(f"    - Firecrawl (Naukri / Verified) : {fc_retained} roles (from {fc_candidates} SERP candidates, {fc_markdown_hits} markdown hits)")
        print(f"    - JobSpy (LinkedIn / Indeed)    : {jobspy_retained} roles (from {jobspy_candidates} candidates)")
        print(f"[+] Firecrawl Credit Audit              : {credit_audit_str}")
        print(log_status_str)
        print("=" * 70)
        if all_sources_failed:
            sys.exit(1)
        sys.exit(0)

    # 2. Build Formatted Excel Tracker
    generated_file = dashboard_gen.generate(enriched_jobs_df, output_path=output_filename, run_stats=row_data)

    # 3. Commit seen jobs ledger only after Excel write succeeds
    scraper._save_seen_jobs()

    # 4. Print Summary Box
    print("\n" + "=" * 70)
    print(f"[+] SUCCESS! Total Active Opportunities Aggregated: {total_unique}")
    print(f"    - Firecrawl (Naukri / Verified) : {fc_retained} roles (from {fc_candidates} SERP candidates, {fc_markdown_hits} markdown hits)")
    print(f"    - JobSpy (LinkedIn / Indeed)    : {jobspy_retained} roles (from {jobspy_candidates} candidates)")
    print(f"    - Funnel Filter Drops           : {dropped_senior_title} senior, {dropped_experience_ceiling} ceiling, {dropped_stale_naukri} stale Naukri, {dropped_future_naukri} future Naukri ({unparseable_naukri_dates} unparseable retained)")
    print(f"[+] Firecrawl Credit Audit              : {credit_audit_str}")
    print(f"[+] Output File                         : {os.path.abspath(generated_file)}")
    print(log_status_str)
    print("=" * 70)
    print("\nNext Steps:")
    print(" 1. Open the generated Excel file.")
    print(" 2. Review the top extracted skills in column F.")
    print(" 3. Click the 'Apply Now' button in column J to jump straight to the post/careers page.")
    print(" 4. Update the 'Application Status' dropdown as you apply in column K.")
    print("=" * 70)


if __name__ == "__main__":
    main()
