"""Shared URL normalisation helpers.

Single source of truth for the canonical job-URL key used by BOTH ledgers:
  * seen_jobs.json    (freshness tagging)   -> scraper.py
  * applications.json (application status)  -> excel_generator.py

These two MUST agree on the key format. If they diverge, every application
status silently detaches from its job. Never re-implement canonicalize_job_url
or IDENTITY_QUERY_PARAMS anywhere else.
"""

import re
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

# Query parameters that uniquely identify a job posting on aggregator/ATS platforms.
IDENTITY_QUERY_PARAMS = {
    "jk", "jl", "id", "jobreferencecode", "jobid", "page", "job"
}


def canonicalize_job_url(url: str) -> str:
    """
    Normalizes a job URL while preserving platform-required identity query parameters
    (Indeed ?jk=, Glassdoor ?jl=, BambooHR/Accenture ?id=, Infosys ?jobReferenceCode=, etc.).
    """
    if not url or not isinstance(url, str):
        return ""
    raw = url.strip()
    if not raw.lower().startswith(("http://", "https://")):
        return ""
    try:
        parsed = urlsplit(raw)
        if not parsed.scheme or not parsed.hostname:
            return ""
        kept_params = [
            (k.lower(), v.strip().lower())
            for k, v in parse_qsl(parsed.query, keep_blank_values=False)
            if k.lower() in IDENTITY_QUERY_PARAMS
        ]
        kept_params.sort()
        clean_path = parsed.path.rstrip("/").lower()
        return urlunsplit((
            parsed.scheme.lower(),
            parsed.hostname.lower(),
            clean_path,
            urlencode(kept_params),
            ""
        ))
    except Exception:
        return re.sub(r'\?.*$', '', raw).strip().lower()
