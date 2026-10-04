import os
import sys
import csv
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

def load_config(config_path: str = "config.yaml") -> dict:
    """Loads configuration from YAML file."""
    if not os.path.exists(config_path):
        print(f"[!] Config file '{config_path}' not found, using default configuration.")
        return {
            "search": {
                "job_titles": ["Data Analyst"],
                "location": "Hyderabad",
                "hours_old": 72,
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

    # Determine Active Mode (CLI flag > config default_mode > "balanced")
    active_mode = args.mode or config.get("default_mode", "balanced")
    
    # Initialize scraper with active mode (single source of truth)
    scraper = MultiPlatformJobScraper(config, active_mode=active_mode)

    # Apply CLI overrides if provided
    if args.role:
        config["search"]["job_titles"] = [args.role]
    if args.location:
        config["search"]["location"] = args.location
    if args.hours:
        config["search"]["hours_old"] = args.hours
    if args.output:
        config["output"]["filename"] = args.output

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

    if enriched_jobs_df.empty:
        print("\n[!] No jobs were returned. Check your internet connection or search parameters.")
        sys.exit(0)

    # Post-flight Credit Check & Delta Computation
    end_credits = scraper.get_remaining_credits()
    if isinstance(start_credits, int) and isinstance(end_credits, int):
        credits_used = max(0, start_credits - end_credits)
        credit_audit_str = f"{start_credits:,} before -> {end_credits:,} after ({credits_used} credits used this run)"
    else:
        credits_used = "unknown"
        credit_audit_str = "Delta calculation unavailable"

    # Read per-engine stats including candidates and markdown hits
    fc_candidates = scraper.last_run_stats.get("fc_candidates", 0)
    fc_markdown_hits = scraper.last_run_stats.get("fc_markdown_hits", 0)
    fc_retained = scraper.last_run_stats.get("fc_retained", 0)
    jobspy_retained = scraper.last_run_stats.get("jobspy_retained", 0)
    total_unique = len(enriched_jobs_df)

    # 2. Build Formatted Excel Tracker
    output_filename = config.get("output", {}).get("filename", "Job_Applications_Tracker.xlsx")
    dashboard_gen = ExcelDashboardGenerator(config)
    generated_file = dashboard_gen.generate(enriched_jobs_df, output_path=output_filename)

    # 3. Log Analytics to run_log.csv with fc_markdown_hits (non-destructive migration)
    log_file = "run_log.csv"
    headers = ["timestamp", "mode", "credits_used", "fc_candidates", "fc_markdown_hits", "fc_retained", "jobspy_retained", "total_unique_jobs"]
    file_exists = os.path.exists(log_file)
    try:
        migrated_rows = []
        needs_migration = not file_exists
        if file_exists:
            with open(log_file, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                old_headers = next(reader, None)
                if old_headers and "fc_markdown_hits" not in old_headers:
                    needs_migration = True
                    for r in reader:
                        if not r:
                            continue
                        # If old format had fc_candidates (7 columns), insert N/A for fc_markdown_hits at index 4
                        if len(r) == 7:
                            r.insert(4, "N/A")
                        migrated_rows.append(r)

        if needs_migration:
            with open(log_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(migrated_rows)

        with open(log_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                active_mode,
                credits_used,
                fc_candidates,
                fc_markdown_hits,
                fc_retained,
                jobspy_retained,
                total_unique
            ])
    except Exception as e:
        print(f"[!] Warning: Could not write to run_log.csv: {e}")

    # 4. Print Summary Box
    print("\n" + "=" * 70)
    print(f"[+] SUCCESS! Total Active Opportunities Aggregated: {total_unique}")
    print(f"    - Firecrawl (Naukri / Verified) : {fc_retained} roles (from {fc_candidates} SERP candidates, {fc_markdown_hits} markdown hits)")
    print(f"    - JobSpy (LinkedIn / Indeed)    : {jobspy_retained} roles")
    print(f"[+] Firecrawl Credit Audit              : {credit_audit_str}")
    print(f"[+] Output File                         : {os.path.abspath(generated_file)}")
    print(f"[+] Run Log Updated                     : {os.path.abspath(log_file)}")
    print("=" * 70)
    print("\nNext Steps:")
    print(" 1. Open the generated Excel file.")
    print(" 2. Review the top extracted skills in column F.")
    print(" 3. Click the 'Apply Now' button in column J to jump straight to the post/careers page.")
    print(" 4. Update the 'Application Status' dropdown as you apply in column K.")
    print("=" * 70)

if __name__ == "__main__":
    main()
