import os
import re
import json
import time
import logging
import requests
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional
from urllib.parse import urlsplit, parse_qs
from dotenv import load_dotenv
from jobspy import scrape_jobs
from extractor import JobRequirementExtractor

# Load environment variables (.env)
load_dotenv()

logger = logging.getLogger("JobScraper")

class MultiPlatformJobScraper:
    """
    Scrapes jobs across LinkedIn, Indeed, Glassdoor, and ZipRecruiter using JobSpy,
    and extends to Naukri, LinkedIn & regional portals using Firecrawl Search v2.
    Features:
    - Experience ceiling gate (drops senior roles > max_experience_years)
    - Naukri URL slug ground-truth experience parsing
    - Cross-run freshness engine via seen_jobs.json (🟢 NEW vs ⚪ SEEN)
    - Ground-truth Firecrawl credit tracking and per-engine attribution
    """

    KNOWN_CITIES = {
        'hyderabad', 'secunderabad', 'bengaluru', 'bangalore', 'pune', 
        'mumbai', 'delhi', 'gurugram', 'gurgaon', 'noida', 'chennai', 
        'kolkata', 'ahmedabad', 'india', 'remote', 'hybrid'
    }

    ROLE_KEYWORDS = {
        'data', 'analyst', 'annotator', 'sql', 'developer', 'power', 
        'bi', 'fresher', 'junior', 'sr', 'senior', 'lead', 'intern', 
        'engineer', 'specialist', 'associate', 'trainee', 'consultant'
    }

    def __init__(self, config: Dict[str, Any], active_mode: str = "balanced"):
        self.config = config.get("search", {})
        self.firecrawl_config = config.get("firecrawl", {})
        self.modes_config = config.get("modes", {})
        self.target_skills = config.get("skills_filter", {}).get("target_skills", [])
        self.extractor = JobRequirementExtractor(self.target_skills)
        self.firecrawl_api_key = os.getenv("FIRECRAWL_API_KEY", "").strip()
        self.active_mode = active_mode

        # Selected mode parameters
        mode_settings = self.modes_config.get(self.active_mode, {})
        self.search_limit_per_role = mode_settings.get("search_limit_per_role", 4)
        self.deep_scrape = mode_settings.get("deep_scrape", True)
        self.timeout = mode_settings.get("timeout", 60 if self.deep_scrape else 20)
        self.max_experience_years = int(self.config.get("max_experience_years", 2))

        # Per-engine retention stats, funnel counters, and candidate count for analytics
        self.last_run_stats = {
            "fc_candidates": 0,
            "fc_markdown_hits": 0,
            "fc_retained": 0,
            "jobspy_retained": 0,
            "dropped_stale_naukri": 0,
            "dropped_future_naukri": 0,
            "unparseable_naukri_dates": 0,
            "dropped_senior_title": 0,
            "dropped_experience_ceiling": 0,
        }

        # Persistent seen jobs ledger
        self.seen_jobs_file = "seen_jobs.json"
        self.seen_jobs = self._load_seen_jobs()

    def _load_seen_jobs(self, max_days: int = 14) -> Dict[str, str]:
        """Loads previously seen job URLs and prunes entries older than max_days."""
        if not os.path.exists(self.seen_jobs_file):
            return {}
        try:
            with open(self.seen_jobs_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            if not isinstance(data, dict):
                return {}

            cutoff = datetime.now()
            pruned = {}
            for url, date_str in data.items():
                try:
                    entry_date = datetime.strptime(str(date_str), "%Y-%m-%d")
                    if (cutoff - entry_date).days <= max_days:
                        pruned[url] = str(date_str)
                except Exception:
                    # Keep if date format unparseable to avoid premature data loss
                    pruned[url] = str(date_str)
            return pruned
        except Exception:
            return {}

    def _save_seen_jobs(self):
        """Saves updated seen job URLs to persistent JSON file."""
        try:
            with open(self.seen_jobs_file, "w", encoding="utf-8") as f:
                json.dump(self.seen_jobs, f, indent=2)
        except Exception as e:
            print(f"[!] Warning: Could not save seen_jobs.json: {e}")

    def get_remaining_credits(self) -> Optional[int]:
        """
        Queries Firecrawl API for live remaining credit balance.
        Defensively handles both v2 camelCase (remainingCredits) and v1 snake_case (remaining_credits).
        Returns integer if successful, or None if unavailable.
        """
        if not self.firecrawl_api_key:
            return None
        
        headers = {"Authorization": f"Bearer {self.firecrawl_api_key}"}
        endpoints = [
            "https://api.firecrawl.dev/v2/team/credit-usage",
            "https://api.firecrawl.dev/v1/team/credit-usage"
        ]

        for url in endpoints:
            try:
                res = requests.get(url, headers=headers, timeout=10)
                if res.status_code == 200:
                    data = res.json().get("data", {})
                    if isinstance(data, dict):
                        val = data.get("remainingCredits", data.get("remaining_credits"))
                        if val is not None and isinstance(val, (int, float)):
                            return int(val)
            except Exception:
                continue

        return None

    def scrape_all_roles(self) -> pd.DataFrame:
        """
        Aggregates verified direct job postings from both Firecrawl v2 and JobSpy,
        tagging engine origins, enforcing experience ceiling, and marking cross-run freshness.
        """
        job_titles = self.config.get("job_titles", ["Data Analyst"])
        location = self.config.get("location", "Hyderabad")
        platforms = self.config.get("platforms", ["linkedin", "indeed", "glassdoor", "zip_recruiter"])
        results_wanted = self.config.get("results_wanted_per_role", 20)
        hours_old = self.config.get("hours_old", 72)
        country_indeed = self.config.get("indeed_country", "india")
        
        firecrawl_enabled = self.firecrawl_config.get("enabled", True)
        use_firecrawl = firecrawl_enabled and bool(self.firecrawl_api_key)

        if firecrawl_enabled and not self.firecrawl_api_key:
            print("\n[!] WARNING: Firecrawl is enabled in config.yaml, but FIRECRAWL_API_KEY is missing in .env.")
            print("    Skipping Firecrawl engine (Naukri/regional search) and continuing with JobSpy.\n")

        # Reset candidate, retention, and funnel counters at the start of each run
        self.last_run_stats["fc_candidates"] = 0
        self.last_run_stats["fc_markdown_hits"] = 0
        self.last_run_stats["fc_retained"] = 0
        self.last_run_stats["jobspy_retained"] = 0
        self.last_run_stats["dropped_stale_naukri"] = 0
        self.last_run_stats["dropped_future_naukri"] = 0
        self.last_run_stats["unparseable_naukri_dates"] = 0
        self.last_run_stats["dropped_senior_title"] = 0
        self.last_run_stats["dropped_experience_ceiling"] = 0

        all_results = []

        # -------------------------------------------------------------
        # 1. Firecrawl v2 Direct Job Search (Naukri, LinkedIn)
        # -------------------------------------------------------------
        if use_firecrawl:
            print(f"\n[*] [Firecrawl v2 Engine] Mode: {self.active_mode.upper()} (Deep Scrape: {self.deep_scrape}, Limit: {self.search_limit_per_role}/role)...")
            for title in job_titles:
                fc_df = self._scrape_with_firecrawl(title, location, limit=self.search_limit_per_role)
                if not fc_df.empty:
                    fc_df["_source_engine"] = "firecrawl"
                    all_results.append(fc_df)
                time.sleep(1)

        # -------------------------------------------------------------
        # 2. Local JobSpy Engine
        # -------------------------------------------------------------
        platforms_str = ", ".join(p.capitalize() for p in platforms)
        print(f"\n[*] [JobSpy Engine] Querying multi-board direct feeds ({platforms_str})...")
        for idx, title in enumerate(job_titles):
            print(f"[*] [Scraping] Searching for '{title}' in '{location}' across {platforms}...")
            try:
                jobs_df = scrape_jobs(
                    site_name=platforms,
                    search_term=title,
                    location=location,
                    results_wanted=results_wanted,
                    hours_old=hours_old,
                    country_indeed=country_indeed,
                    linkedin_fetch_description=True,
                    verbose=0
                )

                if jobs_df is not None and not jobs_df.empty:
                    valid_mask = jobs_df.apply(self._row_has_valid_direct_url, axis=1)
                    clean_df = jobs_df[valid_mask].copy()
                    if not clean_df.empty:
                        clean_df["search_query"] = title
                        clean_df["_source_engine"] = "jobspy"
                        all_results.append(clean_df)
                        print(f"   -> Found {len(clean_df)} direct postings for '{title}'.")
                else:
                    print(f"   -> No listings found for '{title}'.")

            except Exception as e:
                print(f"   [!] Warning: Error while scraping '{title}': {str(e)}")
                continue

            if idx < len(job_titles) - 1:
                time.sleep(2)

        if not all_results:
            print("[!] No job results collected from any platform.")
            return pd.DataFrame()

        combined_df = pd.concat(all_results, ignore_index=True)
        
        # Deduplicate based on company name, title, and location
        print(f"\n[*] Total raw direct jobs collected: {len(combined_df)}")
        deduped_df = self._deduplicate_jobs(combined_df)
        print(f"[*] Total unique verified jobs after deduplication: {len(deduped_df)}")

        # Enrich with skills, slug ground-truth experience, ceiling gate, and freshness tags
        enriched_df = self._enrich_jobs(deduped_df)
        print(f"[*] Final candidates within experience ceiling (<= {self.max_experience_years} YOE): {len(enriched_df)}")

        # Record exact per-engine retention statistics
        if "_source_engine" in enriched_df.columns:
            self.last_run_stats["fc_retained"] = int((enriched_df["_source_engine"] == "firecrawl").sum())
            self.last_run_stats["jobspy_retained"] = int((enriched_df["_source_engine"] == "jobspy").sum())
            enriched_df.drop(columns=["_source_engine"], inplace=True)

        # Persist newly seen jobs
        self._save_seen_jobs()

        return enriched_df

    def _scrape_with_firecrawl(self, title: str, location: str, limit: int = 4) -> pd.DataFrame:
        """
        Uses Firecrawl Search v2 API with optional full markdown deep scraping.
        Unshackled query targeted to direct job listing paths.
        """
        headers = {
            "Authorization": f"Bearer {self.firecrawl_api_key}",
            "Content-Type": "application/json"
        }
        
        query = f'"{title}" "{location}" site:naukri.com/job-listings'
        
        payload: Dict[str, Any] = {
            "query": query,
            "limit": limit,
            "tbs": "qdr:w"
        }

        if self.deep_scrape:
            payload["scrapeOptions"] = {"formats": ["markdown"]}

        try:
            print(f"   [Firecrawl v2] Querying '{title}' (\"{location}\")...")
            res = requests.post(
                "https://api.firecrawl.dev/v2/search",
                headers=headers,
                json=payload,
                timeout=self.timeout
            )
            
            if res.status_code == 402:
                print("   [!] Firecrawl credits depleted (HTTP 402 Payment Required). Continuing with local JobSpy.")
                return pd.DataFrame()

            if res.status_code != 200:
                print(f"   [!] Firecrawl v2 returned status {res.status_code}: {res.text[:100]}")
                return pd.DataFrame()

            raw_data = res.json().get("data", {})
            items = raw_data.get("web", []) if isinstance(raw_data, dict) else (raw_data if isinstance(raw_data, list) else [])
            
            self.last_run_stats["fc_candidates"] += len(items)
            self.last_run_stats["fc_markdown_hits"] += sum(
                1 for item in items if (item.get("markdown") or "").strip()
            )

            if not items:
                return pd.DataFrame()

            rows = []
            for item in items:
                url = item.get("url", "")
                title_raw = item.get("title", "")
                desc = item.get("description", "")
                
                if not self._is_direct_job_url(url):
                    continue

                if "naukri.com" in url:
                    site = "Naukri"
                elif "linkedin.com" in url:
                    site = "LinkedIn"
                elif "indeed.com" in url:
                    site = "Indeed"
                else:
                    site = "Direct Careers"

                role, company = self._extract_clean_metadata(title_raw, url, site)

                if self._is_invalid_company(company):
                    continue

                markdown = (item.get("markdown") or "").strip()
                body = markdown if (self.deep_scrape and markdown) else desc

                rows.append({
                    "company": company,
                    "title": role,
                    "location": location,
                    "site": site,
                    "job_url": url,
                    "job_url_direct": url,
                    "description": f"{title_raw}\n\n{body}",
                    "date_posted": "Recent",
                    "search_query": title
                })

            print(f"   -> [Firecrawl v2] Found {len(items)} candidates, retained {len(rows)} verified direct jobs for '{title}'.")
            return pd.DataFrame(rows)

        except requests.exceptions.Timeout:
            print(f"   [!] Firecrawl request timed out after {self.timeout}s for '{title}'.")
            return pd.DataFrame()
        except Exception as e:
            print(f"   [!] Firecrawl Search Error: {str(e)}")
            return pd.DataFrame()

    @staticmethod
    def _matches_host(host: str, target_domain: str) -> bool:
        """Strict domain boundary check: matches target_domain or its subdomains."""
        return host == target_domain or host.endswith("." + target_domain)

    @staticmethod
    def _clean_http_url(val: Any) -> Optional[str]:
        """Returns string URL if valid non-empty http(s) string with valid hostname, else None."""
        if not isinstance(val, str):
            return None
        val = val.strip()
        if not val:
            return None
        try:
            parsed = urlsplit(val)
            if parsed.scheme.lower() in ("http", "https") and parsed.hostname:
                return val
        except Exception:
            return None
        return None

    # Dedicated ATS platforms with required individual job page path patterns
    KNOWN_ATS_JOB_PATTERNS = {
        "greenhouse.io": r'/[a-zA-Z0-9_-]+/jobs/\d+',
        "lever.co": r'/[a-zA-Z0-9_-]+/[0-9a-fA-F-]{8,}(?:/apply)?',
        "myworkdayjobs.com": r'(?:/[a-zA-Z0-9_.-]+){1,3}/job/(?:[^/]+/)?(?:[a-zA-Z0-9_.-]+_)?[a-zA-Z0-9_-]+',
        "myworkdaysite.com": r'(?:/[a-zA-Z0-9_.-]+){1,3}/job/(?:[^/]+/)?(?:[a-zA-Z0-9_.-]+_)?[a-zA-Z0-9_-]+',
        "smartrecruiters.com": r'/[a-zA-Z0-9_-]+/\d{8,}(?:-[a-zA-Z0-9_-]+)?',
        "ashbyhq.com": r'/[a-zA-Z0-9_-]+/[a-zA-Z0-9_-]{8,}(?:/application)?',
        "taleo.net": r'/careersection/[^/]+/jobdetail\.ftl',
        "icims.com": r'/jobs/\d+/(?:[a-zA-Z0-9_-]+/)?job',
        "keka.com": r'/careers/jobdetails/[a-zA-Z0-9_-]+',
        "workable.com": r'(?:/[a-zA-Z0-9_-]+)?/j/[0-9a-zA-Z]{6,}',
    }

    # Enterprise career portals with required individual job page patterns
    KNOWN_ENTERPRISE_PORTALS = {
        "careers.google.com": r'/jobs/results/\d+(?:-[a-zA-Z0-9_-]+)?',
        "amazon.jobs": r'(?:/[a-zA-Z]{2})?/jobs/\d+(?:/[a-zA-Z0-9_-]+)?',
        "careers.microsoft.com": r'(?:/professionals|/careers)?(?:/[a-zA-Z]{2}(?:/[a-zA-Z]{2})?)?/job/\d+(?:/[a-zA-Z0-9_-]+)?',
        "wipro.com": r'/job/[^/]+/\d+(?:-[a-zA-Z_]+)?',
        "hcltech.com": r'/job/[^/]+/\d+(?:-[a-zA-Z_]+)?',
        "hsbc.com": r'/job/[^/]+/\d+',
        "micron.com": r'(?:/careers)?/job/\d+',
        "trinet.com": r'(?:/careers)?/job/\d+',
        "pepsicojobs.com": r'/jobs/\d+/(?:[a-zA-Z0-9_-]+/)?job',
        "infor.com": r'/postings/[0-9a-fA-F-]{36}',
        "oraclecloud.com": r'/hcmUI/CandidateExperience/[a-zA-Z]{2}/sites/[^/]+/requisitions/preview/\d+',
        "deloitte.com": r'/(?:[a-zA-Z]{2}_[a-zA-Z]{2}/)?careersUSI/JobDetail/[^/]+/\d+',
        "ttcportals.com": r'/jobs/\d+-[a-zA-Z0-9_-]+',
    }

    def _row_has_valid_direct_url(self, r: Any) -> bool:
        """Checks if either job_url_direct or job_url is a valid direct job listing."""
        u_direct = self._clean_http_url(r.get("job_url_direct"))
        u_job = self._clean_http_url(r.get("job_url"))
        return bool(
            (u_direct and self._is_direct_job_url(u_direct)) or 
            (u_job and self._is_direct_job_url(u_job))
        )

    def _is_direct_job_url(self, url: str) -> bool:
        """
        Validates whether a URL points directly to an individual job application page.
        Fail-closed default: rejects unknown hosts, search pages, non-job paths, and trailing garbage.
        """
        clean_url = self._clean_http_url(url)
        if not clean_url:
            return False

        try:
            parsed = urlsplit(clean_url)
            query_params = parse_qs(parsed.query)
        except Exception:
            return False

        host = (parsed.hostname or "").lower()
        path = parsed.path.rstrip('/')
        u_lower = clean_url.lower()

        disallowed_patterns = [
            "-jobs-in-", "-vacancies-in-", "/jobs-in-", "job-vacancies", 
            "/jobs/search", "/jobs/collections", "/jobs?q=", "/cmp/", 
            "tcs-power-bi-jobs", "analyst-jobs", "developer-jobs",
            "/jobs-search", "/jobs?search="
        ]
        if any(bad in u_lower for bad in disallowed_patterns):
            return False

        # 1. Naukri Direct Job (requires numeric ID at end of path)
        if self._matches_host(host, "naukri.com"):
            return bool(re.fullmatch(r'/job-listings-[a-zA-Z0-9_-]+-\d{10,14}', path))

        # 2. LinkedIn Direct Job
        if self._matches_host(host, "linkedin.com"):
            return bool(re.fullmatch(r'/jobs/view/(?:[a-zA-Z0-9_-]+-)?\d+', path))

        # 3. Indeed Direct Job (validates full hex token on viewjob path or direct hash slug)
        if self._matches_host(host, "indeed.com"):
            jk_vals = query_params.get("jk", [])
            has_valid_jk = bool(jk_vals and re.fullmatch(r'[a-f0-9]{16}', jk_vals[0], re.I))
            is_viewjob_path = bool(re.fullmatch(r'/(?:m/)?viewjob|/rc/clk', path))
            is_direct_slug = bool(re.fullmatch(r'/job/[a-zA-Z0-9_-]+-[a-f0-9]{16}', path))
            return (has_valid_jk and is_viewjob_path) or is_direct_slug

        # 4. Glassdoor Direct Job
        if self._matches_host(host, "glassdoor.com") or self._matches_host(host, "glassdoor.co.in"):
            jl_vals = query_params.get("jl", [])
            has_valid_jl = bool(jl_vals and re.fullmatch(r'\d+', jl_vals[0]))
            is_listing_path = bool(re.fullmatch(r'/job-listing/[a-zA-Z0-9_-]+', path))
            is_htm_path = bool(re.fullmatch(r'/job/[a-zA-Z0-9_-]+\.htm', path, re.I))
            return (has_valid_jl and is_listing_path) or is_htm_path

        # 5. ZipRecruiter & ZipApply Direct Job
        if self._matches_host(host, "ziprecruiter.com"):
            return bool(
                re.fullmatch(r'/jobs/(?:[a-zA-Z0-9_-]*/)?\d+', path) or
                re.fullmatch(r'/job/[a-zA-Z0-9_-]+/[a-zA-Z0-9_-]+', path) or
                re.fullmatch(r'/c/[a-zA-Z0-9_-]+/job/[a-zA-Z0-9_-]+', path) or
                re.fullmatch(r'/k/job/[a-zA-Z0-9_-]+', path)
            )
        if self._matches_host(host, "zipapply.com"):
            return bool(re.fullmatch(r'/(?:job|jobs|apply)/[a-zA-Z0-9_-]+', path))

        # 6. BambooHR (query id validation or numeric career path)
        if self._matches_host(host, "bamboohr.com"):
            if re.fullmatch(r'/careers/\d+', path):
                return True
            id_vals = query_params.get("id", [])
            if path == '/jobs/view.php' and id_vals and re.fullmatch(r'\d+', id_vals[0]):
                return True
            return False

        # 7. Accenture (requires job id query parameter)
        if self._matches_host(host, "accenture.com"):
            id_vals = query_params.get("id", [])
            return bool(
                id_vals and re.fullmatch(r'[a-zA-Z0-9_-]+', id_vals[0]) and
                re.fullmatch(r'/(?:[a-zA-Z]{2}-[a-zA-Z]{2}/)?careers/jobdetails', path, re.IGNORECASE)
            )

        # 8. Infosys (requires jobReferenceCode query parameter)
        if self._matches_host(host, "infosys.com"):
            ref_vals = query_params.get("jobReferenceCode", [])
            return bool(
                ref_vals and re.fullmatch(r'[a-zA-Z0-9_-]+', ref_vals[0]) and
                re.fullmatch(r'/jobdesc', path, re.IGNORECASE)
            )

        # 9. Njoyn ATS (requires page=jobdetails and JobID parameter)
        if self._matches_host(host, "njoyn.com"):
            page_vals = query_params.get("page", [])
            job_vals = query_params.get("JobID", [])
            return bool(
                page_vals == ["jobdetails"] and job_vals and
                re.fullmatch(r'[a-zA-Z0-9_-]+', job_vals[0]) and
                re.fullmatch(r'/corp/xweb/XWeb\.asp', path, re.IGNORECASE)
            )

        # 10. Dedicated ATS Providers
        for ats_domain, pattern in self.KNOWN_ATS_JOB_PATTERNS.items():
            if self._matches_host(host, ats_domain):
                if ats_domain == "taleo.net":
                    job_vals = query_params.get("job", [])
                    has_valid_job = bool(job_vals and re.fullmatch(r'\d+', job_vals[0]))
                    return has_valid_job and bool(re.fullmatch(pattern, path, re.IGNORECASE))
                return bool(re.fullmatch(pattern, path, re.IGNORECASE))

        # 11. Enterprise Portals
        for portal_domain, pattern in self.KNOWN_ENTERPRISE_PORTALS.items():
            if self._matches_host(host, portal_domain):
                return bool(re.fullmatch(pattern, path, re.IGNORECASE))

        return False

    def _extract_clean_metadata(self, raw_title: str, url: str, site: str) -> Tuple[str, str]:
        """Extracts clean Role Title and actual Company Name from the URL and title text."""
        role = raw_title
        company = "Hiring Company"

        if site == "Naukri":
            path = urlsplit(url).path.rstrip('/')
            m = re.search(r'/job-listings-(.+?)-\d{10,14}$', path)
            if m:
                slug = m.group(1).lower()
                slug = re.sub(r'-\d+-to-\d+-years?', '', slug)
                tokens = slug.split('-')
                
                start_idx = 0
                while start_idx < len(tokens) and tokens[start_idx] in self.ROLE_KEYWORDS:
                    start_idx += 1
                
                end_idx = len(tokens)
                while end_idx > start_idx and (tokens[end_idx-1] in self.KNOWN_CITIES or tokens[end_idx-1] in ['hiring', 'opening', 'job', 'jobs']):
                    end_idx -= 1
                
                company_tokens = tokens[start_idx:end_idx]
                deduped_tokens = []
                for t in company_tokens:
                    if t not in deduped_tokens and t != 'hiring':
                        deduped_tokens.append(t)
                        
                if deduped_tokens:
                    candidate_co = ' '.join(deduped_tokens).title()
                    if candidate_co.lower() not in self.KNOWN_CITIES:
                        company = candidate_co

            cleaned = re.sub(r'\s*-\s*\d+\s*(?:to|-)\s*\d+\s*years?.*$', '', raw_title, flags=re.I).strip()
            cleaned = re.sub(r'\s*[-|]\s*Naukri\.com.*$', '', cleaned, flags=re.I).strip()
            parts = [p.strip() for p in cleaned.split(' - ')]
            if len(parts) >= 2 and parts[0].lower() in ["fresher", "intern", "junior", "sr"]:
                role = f"{parts[0]} {parts[1]}"
            elif parts:
                role = parts[0]

        elif site == "LinkedIn":
            m = re.search(r'/jobs/view/([^/?]+)', url)
            if m and "-at-" in m.group(1):
                slug_parts = m.group(1).split("-at-")
                role = slug_parts[0].replace("-", " ").title()
                co_part = re.sub(r'-\d+$', '', slug_parts[1]).replace("-", " ").title()
                company = co_part
            else:
                cleaned = re.sub(r'\s*[-|]\s*LinkedIn.*$', '', raw_title, flags=re.I).strip()
                if " at " in cleaned:
                    parts = cleaned.split(" at ", 1)
                    role = parts[0].strip()
                    company = parts[1].split(" - ")[0].strip()
                elif " hiring " in cleaned:
                    parts = cleaned.split(" hiring ", 1)
                    company = parts[0].strip()
                    role = parts[1].split(" in ")[0].strip()

            company = re.sub(r'\s*[-–]\s*(?:Greater\s+)?Hyderabad.*$', '', company, flags=re.I).strip()

        else:
            role = raw_title.split(" - ")[0].strip()

        return role, company

    def _is_invalid_company(self, company: str) -> bool:
        """Checks if company name is a placeholder, city, or blank."""
        if not company or len(company) < 2:
            return True
        c_lower = company.lower().strip()
        if c_lower in self.KNOWN_CITIES or c_lower in ["hiring company", "company", "unknown company"]:
            return True
        if any(f"{city}," in c_lower for city in self.KNOWN_CITIES):
            return True
        return False

    def _extract_naukri_slug_experience(self, url: str) -> Optional[str]:
        """
        Finding 1 Fix: Extracts authoritative ground-truth experience range from Naukri URL slug.
        e.g., ...careernet-hyderabad-5-to-8-years-120626011886 -> '5 - 8 Years'
        """
        try:
            parsed = urlsplit((url or "").strip())
            host = (parsed.hostname or "").lower()
            if not self._matches_host(host, "naukri.com"):
                return None
            path = parsed.path.rstrip('/')
            m = re.search(r'-(\d{1,2})-to-(\d{1,2})-years?-', path.lower())
            if m:
                return f"{m.group(1)} - {m.group(2)} Years"
        except Exception:
            return None
        return None

    @staticmethod
    def _naukri_posted_date(url: str) -> Optional[datetime]:
        """
        Decodes the DDMMYY creation date prefix embedded in Naukri trailing job IDs.
        e.g., job-listings-...-150225005334 -> 15-02-25 (Feb 15, 2025).
        Returns datetime object if valid calendar date, else None.
        """
        try:
            parsed = urlsplit(url.strip())
            host = (parsed.hostname or "").lower()
            if not (host == "naukri.com" or host.endswith(".naukri.com")):
                return None
            path = parsed.path.rstrip('/')
            m = re.search(r'-(\d{6})\d{4,8}$', path)
            if not m:
                return None
            return datetime.strptime(m.group(1), "%d%m%y")
        except Exception:
            return None

    def _within_ceiling(self, exp_str: str, max_years: int = 2) -> bool:
        """
        Finding 1 Fix: Gates rows against max experience ceiling.
        Returns False if minimum required experience > max_years.
        Leniently returns True for unparsable or fresher rows.
        """
        if not exp_str or exp_str in ["Not Specified", "General Requirements"]:
            return True

        exp_lower = exp_str.lower()
        if any(k in exp_lower for k in ["fresher", "intern", "entry", "trainee", "new grad", "graduate"]):
            return True

        # 1. Forward Mixed Unit Range: "6 months to 1 year" -> min = 6/12 = 0.5 yrs
        m_mixed = re.search(r'(\d+(?:\.\d+)?)\s*(?:months?|mos?)\s*(?:-|to)\s*(\d+(?:\.\d+)?)\s*(?:years?|yrs?)', exp_lower)
        if m_mixed:
            return (float(m_mixed.group(1)) / 12.0) <= max_years

        # 2. Reverse Mixed Unit Range: "1 year to 36 months" -> min = 1.0 yrs
        m_rev_mixed = re.search(r'(\d+(?:\.\d+)?)\s*(?:years?|yrs?)\s*(?:-|to)\s*(\d+(?:\.\d+)?)\s*(?:months?|mos?)', exp_lower)
        if m_rev_mixed:
            return float(m_rev_mixed.group(1)) <= max_years

        # 3. Month Range: "3 - 6 months" -> min = 3/12 = 0.25 yrs
        m_mrange = re.search(r'(\d+(?:\.\d+)?)\s*(?:-|to)\s*(\d+(?:\.\d+)?)\s*(?:months?|mos?)', exp_lower)
        if m_mrange:
            return (float(m_mrange.group(1)) / 12.0) <= max_years

        # 4. Single / Plus Months: "6 months", "6+ mos" -> min = 6/12 = 0.5 yrs
        m_msingle = re.search(r'(\d+(?:\.\d+)?)\+?\s*(?:months?|mos?)', exp_lower)
        if m_msingle:
            return (float(m_msingle.group(1)) / 12.0) <= max_years

        # 5. Year Range: "0.5 - 2 Years", "2 to 4 years"
        m_range = re.search(r'(\d+(?:\.\d+)?)\s*(?:-|to)\s*(\d+(?:\.\d+)?)\s*(?:years?|yrs?)', exp_lower)
        if m_range:
            return float(m_range.group(1)) <= max_years

        # 6. Single / Plus Years: "1.5+ Years", "5+ yrs"
        m_plus = re.search(r'(\d+(?:\.\d+)?)\+?\s*(?:years?|yrs?)', exp_lower)
        if m_plus:
            return float(m_plus.group(1)) <= max_years

        return True

    SENIOR_TITLE_PATTERN = re.compile(
        r'\b(?:sr\.?|senior|lead|principal|architect|manager|head|director|vp)\b',
        re.IGNORECASE
    )

    def _is_senior_title(self, title: str, url: str = "") -> bool:
        """Detects senior, lead, or management postings in title or URL slug."""
        text = f"{title} {url}".replace("-", " ").replace("_", " ").lower()
        return bool(self.SENIOR_TITLE_PATTERN.search(text))

    def _deduplicate_jobs(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove duplicate postings while preserving distinct openings."""
        if df.empty:
            return df

        df['clean_company'] = df['company'].fillna('').astype(str).str.strip().str.lower()
        df['clean_title'] = df['title'].fillna('').astype(str).str.strip().str.lower()
        df['clean_location'] = df['location'].fillna('').astype(str).str.strip().str.lower()
        
        deduped = df.drop_duplicates(subset=['clean_company', 'clean_title', 'clean_location'], keep='first').copy()
        deduped.drop(columns=['clean_company', 'clean_title', 'clean_location'], inplace=True)
        return deduped

    def _enrich_jobs(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Parse descriptions to extract skills, experience, and clean apply links.
        Applies:
        - Senior title & slug heuristic gate (<= 2 YOE ceiling)
        - Naukri URL slug ground-truth experience override
        - Post-fetch experience ceiling gate (<= max_experience_years)
        - Cross-run freshness tracker (🟢 NEW vs ⚪ SEEN)
        """
        enriched_rows = []
        today = datetime.now().date()
        today_str = today.strftime("%Y-%m-%d")

        for _, row in df.iterrows():
            job_dict = row.to_dict()
            extracted_info = self.extractor.process_job(job_dict)
            
            u_direct = self._clean_http_url(row.get("job_url_direct"))
            u_job = self._clean_http_url(row.get("job_url"))
            if u_direct and self._is_direct_job_url(u_direct):
                final_url_str = u_direct
            elif u_job and self._is_direct_job_url(u_job):
                final_url_str = u_job
            else:
                final_url_str = u_direct or u_job or ""

            # Finding: Drop stale Naukri postings older than 30 days
            naukri_date = None
            date_status = None
            final_host = (urlsplit(final_url_str).hostname or "").lower() if final_url_str else ""
            if self._matches_host(final_host, "naukri.com"):
                naukri_date = self._naukri_posted_date(final_url_str)
                if naukri_date is not None:
                    post_date = naukri_date.date()
                    if post_date > today:
                        self.last_run_stats["dropped_future_naukri"] += 1
                        continue
                    age_days = (today - post_date).days
                    if age_days > 30:
                        self.last_run_stats["dropped_stale_naukri"] += 1
                        continue
                    date_status = "parsed"
                else:
                    self.last_run_stats["unparseable_naukri_dates"] += 1
                    date_status = "unparseable"

            # Amendment 1 Gate: Drop senior/lead/management roles before slug/ceiling checks
            if self.max_experience_years <= 2 and self._is_senior_title(str(job_dict.get("title", "")), final_url_str):
                self.last_run_stats["dropped_senior_title"] += 1
                continue

            # Finding 1: Check authoritative Naukri slug ground-truth first
            slug_exp = self._extract_naukri_slug_experience(final_url_str)
            final_exp = slug_exp if slug_exp else extracted_info["experience"]

            # Finding 1 Gate: Drop if minimum experience requirement exceeds ceiling
            if not self._within_ceiling(final_exp, max_years=self.max_experience_years):
                self.last_run_stats["dropped_experience_ceiling"] += 1
                continue

            # Feature 2: Cross-run freshness tagging
            norm_url = re.sub(r'\?.*$', '', final_url_str).strip().lower()
            if norm_url in self.seen_jobs:
                lead_status = "⚪ SEEN"
            else:
                lead_status = "🟢 NEW"
                if norm_url:
                    self.seen_jobs[norm_url] = today_str

            job_dict["lead_status"] = lead_status
            job_dict["extracted_skills"] = extracted_info["skills"]
            job_dict["extracted_experience"] = final_exp
            job_dict["work_mode"] = extracted_info["work_mode"]
            job_dict["final_apply_url"] = final_url_str

            date_val = row.get("date_posted")
            if pd.notna(date_val) and str(date_val).strip() not in ["", "Recent"]:
                job_dict["formatted_date"] = str(date_val)
            elif naukri_date:
                job_dict["formatted_date"] = naukri_date.strftime("%Y-%m-%d")
            elif date_status == "unparseable":
                job_dict["formatted_date"] = "Unverified (Date Parse Failed)"
            else:
                job_dict["formatted_date"] = "Recent"

            enriched_rows.append(job_dict)

        return pd.DataFrame(enriched_rows)
