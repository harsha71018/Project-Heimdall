import os
import re
import json
import time
import logging
import requests
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional
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

        # Per-engine retention stats and candidate count for analytics
        self.last_run_stats = {
            "fc_candidates": 0,
            "fc_markdown_hits": 0,
            "fc_retained": 0,
            "jobspy_retained": 0
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

        # Reset candidate and retention counters at the start of each run
        self.last_run_stats["fc_candidates"] = 0
        self.last_run_stats["fc_markdown_hits"] = 0
        self.last_run_stats["fc_retained"] = 0
        self.last_run_stats["jobspy_retained"] = 0

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
        # 2. Local JobSpy Engine (LinkedIn, Indeed, Glassdoor, ZipRecruiter)
        # -------------------------------------------------------------
        print("\n[*] [JobSpy Engine] Querying multi-board direct feeds (LinkedIn, Indeed, Glassdoor, ZipRecruiter)...")
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
                    verbose=0
                )

                if jobs_df is not None and not jobs_df.empty:
                    valid_mask = jobs_df.apply(
                        lambda r: self._is_direct_job_url(r.get("job_url_direct") or r.get("job_url")), 
                        axis=1
                    )
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
        
        query = f'"{title}" "{location}" site:naukri.com/job-listings OR site:linkedin.com/jobs/view'
        
        payload: Dict[str, Any] = {
            "query": query,
            "limit": limit
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

    def _is_direct_job_url(self, url: str) -> bool:
        """
        Validates whether a URL points directly to an individual job application page.
        """
        if not url or not isinstance(url, str) or not url.startswith("http"):
            return False

        disallowed_patterns = [
            "-jobs-in-", "-vacancies-in-", "/jobs-in-", "job-vacancies", 
            "/jobs/search", "/jobs/collections", "/jobs?q=", "/cmp/", 
            "tcs-power-bi-jobs", "analyst-jobs", "developer-jobs",
            "/jobs-search", "/jobs?search="
        ]
        if any(bad in url for bad in disallowed_patterns):
            return False

        # 1. Naukri Direct Job (requires numeric job ID at end)
        if "naukri.com" in url:
            return bool(re.search(r'/job-listings-.+?-\d{10,14}$', url))

        # 2. LinkedIn Direct Job
        if "linkedin.com" in url:
            return bool(re.search(r'/jobs/view/(?:.+?-)?\d+', url))

        # 3. Indeed Direct Job
        if "indeed.com" in url:
            return ('/viewjob' in url or bool(re.search(r'/job/[^/?]+-[a-f0-9]{16}', url)))

        # 4. Glassdoor Direct Job
        if "glassdoor." in url:
            return any(k in url.lower() for k in ["joblisting", "job-listing", "jl=", "/job/"])

        # 5. ZipRecruiter Direct Job
        if "ziprecruiter." in url or "zipapply." in url:
            u_lower = url.lower()
            return ("lvk=" in u_lower or "/jobs//" in u_lower or "/job/" in u_lower or "/c/" in u_lower or "zipapply.com" in u_lower)

        # 6. Direct ATS / Company Career URLs
        ats_keywords = [
            'workday', 'taleo', 'icims', 'greenhouse', 'lever', 
            'keka', 'smartrecruiters', 'careers.', 'jobs.', '/job/', '/careers/'
        ]
        if any(ats in url for ats in ats_keywords):
            return True

        return False

    def _extract_clean_metadata(self, raw_title: str, url: str, site: str) -> Tuple[str, str]:
        """Extracts clean Role Title and actual Company Name from the URL and title text."""
        role = raw_title
        company = "Hiring Company"

        if site == "Naukri":
            m = re.search(r'/job-listings-(.+?)-\d{10,14}$', url)
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
        if "naukri.com" not in url:
            return None
        m = re.search(r'-(\d{1,2})-to-(\d{1,2})-years?-', url.lower())
        if m:
            return f"{m.group(1)} - {m.group(2)} Years"
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

        # Match ranges like "5 - 8 Years", "3 to 5 years"
        m_range = re.search(r'(\d+)\s*(?:-|to)\s*(\d+)', exp_lower)
        if m_range:
            return int(m_range.group(1)) <= max_years

        # Match plus forms like "5+ Years", "6+ yrs", "11+ Years"
        m_plus = re.search(r'(\d+)\+?\s*(?:years?|yrs?)', exp_lower)
        if m_plus:
            return int(m_plus.group(1)) <= max_years

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
        today_str = datetime.now().strftime("%Y-%m-%d")

        for _, row in df.iterrows():
            job_dict = row.to_dict()
            extracted_info = self.extractor.process_job(job_dict)
            
            direct_url = row.get("job_url_direct")
            general_url = row.get("job_url")
            final_url = direct_url if (isinstance(direct_url, str) and direct_url.startswith("http")) else general_url
            final_url_str = final_url if isinstance(final_url, str) else ""

            # Amendment 1 Gate: Drop senior/lead/management roles before slug/ceiling checks
            if self.max_experience_years <= 2 and self._is_senior_title(str(job_dict.get("title", "")), final_url_str):
                continue

            # Finding 1: Check authoritative Naukri slug ground-truth first
            slug_exp = self._extract_naukri_slug_experience(final_url_str)
            final_exp = slug_exp if slug_exp else extracted_info["experience"]

            # Finding 1 Gate: Drop if minimum experience requirement exceeds ceiling
            if not self._within_ceiling(final_exp, max_years=self.max_experience_years):
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
            job_dict["formatted_date"] = str(date_val) if pd.notna(date_val) else "Recent"

            enriched_rows.append(job_dict)

        return pd.DataFrame(enriched_rows)
