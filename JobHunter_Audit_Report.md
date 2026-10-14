# JobHunter — Engineering Audit Report

**Audit date:** 14 October 2026  
**Scope:** Uploaded `JobHunter.zip`; reviewed `main.py`, `scraper.py`, `extractor.py`, `excel_generator.py`, `config.yaml`, Windows launchers, requirements, README, and tests.  
**Approach:** Static review, compile check, existing test suite, and targeted runtime probes. No live platform scraping was performed.

## Executive assessment

**Overall: 7.5/10 as a personal automation MVP; approximately 6.5/10 for dependable unattended operation.** The architecture is sensible and the Excel persistence code is better than a typical one-off scraper. The highest-value improvements concern the *truthfulness and completeness of the data*, rather than a wholesale rewrite.

The pipeline has two ingestion branches, direct-URL checks, canonicalization/deduplication, title/experience/date filters, persistent “seen” state, skills/work-mode extraction, an Excel tracker with application-status persistence, and CSV run analytics.

### Verification performed

- `python -m compileall -q .`: passed.
- `python -m pytest -q`: could not collect tests in the audit environment because `python-jobspy` is not installed.
- Re-ran tests with a temporary `jobspy` import stub outside the project: **30 tests passed, 151 subtests passed**.
- Targeted probes confirmed the ATS metadata rejection, case-folding URL identity collision, and ignored configured output filename described below.
- No real LinkedIn/Indeed/Firecrawl request was made. The passing tests therefore validate local logic only, not live source integration or rate-limit behavior.

## Pipeline map

1. `main.py` loads YAML and applies CLI overrides.
2. `MultiPlatformJobScraper` loads a 45-day URL ledger.
3. If enabled and keyed, Firecrawl Search v2 is queried; JobSpy is queried for each configured role.
4. Rows without a URL accepted by `_is_direct_job_url()` are removed.
5. Results are concatenated, canonical URL duplicates are removed, and company/title/location duplicates across hosts are heuristically collapsed.
6. Descriptions are parsed for skills, experience, and work mode; Naukri URL slugs supply experience/date hints; senior titles and experience above the ceiling are filtered.
7. Accepted URLs get NEW/SEEN tags and are persisted to `seen_jobs.json`.
8. `main.py` appends/migrates `run_log.csv`; the Excel generator synchronizes historical workbooks with `applications.json` and builds the tracker plus Run Summary tab.

The division of responsibility is broadly good. `scraper.py` is, however, doing too much: URL allowlisting, metadata parsing, date parsing, filtering, deduplication, persistence, API calls, and stats all live in the same class. That is maintainable at this size, but the pipeline needs clearer per-stage outcomes before more source types are added.

## Confirmed / high-value findings

### P1 — Firecrawl query coverage is much narrower than the project description

**Location:** `scraper.py:347` (query construction), `scraper.py:262–270` (engine orchestration), README sections “How It Works” and “Dedicated IT Support Pipeline”.

The Firecrawl query is hardcoded to:

```python
query = f'"{title}" "{location}" site:naukri.com/job-listings'
```

So this branch only searches Naukri job-listing paths. It does not search a list of ATS or corporate career domains, despite comments/docs describing LinkedIn and regional/corporate career-board coverage. LinkedIn/Indeed still have a separate JobSpy path, but Firecrawl's advertised incremental coverage is not implemented.

**Patch:** move source domains into config, use Firecrawl Search's `includeDomains` field or explicitly build per-domain queries, and track candidate/retained counts per domain. Do not broaden the query alone until the metadata blocker below is fixed.

### P1 — Valid non-Naukri direct-career results are currently dropped

**Location:** `scraper.py:395–407`, `scraper.py:609–672`, `scraper.py:674–683`.

For hosts that are neither Naukri, LinkedIn, Indeed, nor Indeed India, Firecrawl sets `site = "Direct Careers"`. `_extract_clean_metadata()` then takes the generic `else` branch, changes the title but leaves `company = "Hiring Company"`. `_is_invalid_company()` rejects that placeholder. I reproduced this with a valid Greenhouse job URL: `_is_direct_job_url()` accepted it, but the company extractor returned the placeholder and the row was rejected.

**Impact:** adding ATS/corporate domains to search will still yield zero retained records for those domains unless company identification is implemented.

**Patch:** create a source-aware company resolver. Use the result title/metadata and URL path/domain as evidence; if the company cannot be inferred confidently, retain an explicit `Company Unverified` value with a confidence flag rather than silently deleting the listing.

### P1 — The “strict experience ceiling” is not strict for unknown experience

**Location:** `scraper.py:723–771`; call site `scraper.py:921–928`.

`_within_ceiling()` returns `True` for `Not Specified`, `General Requirements`, an empty value, or a description from which no numeric requirement was recognized. This means jobs with unknown requirements pass the same gate as verified 0–2 YOE jobs. The behavior may be a sensible recall-first decision, but it does not match the “strict ceiling” language in the code and README.

**Patch:** model the outcome as three states: `eligible`, `ineligible`, and `unverified`. Keep unverified listings only if configured to do so, and mark them clearly in the tracker. Add tests for “unknown but retained”, “unknown and excluded”, and numeric cases; don't make missing data indistinguishable from verified junior eligibility.

### P1 — URL canonicalization can merge distinct job identities that differ by case

**Location:** `scraper.py:26–51`; duplicated implementation in `excel_generator.py:22–51`.

The canonicalizer lowercases both the URL path and identity-query values:

```python
(k.lower(), v.strip().lower())
clean_path = parsed.path.rstrip("/").lower()
```

I confirmed that `/Job/ABC` and `/job/abc` canonicalize to the same key, as do `jobReferenceCode=AbC-123` and `jobReferenceCode=abc-123`. Host names are case-insensitive, but URL paths and arbitrary identifier values are not universally case-insensitive.

**Impact:** possible false deduplication, incorrect NEW/SEEN labels, and cross-contamination of saved application status/notes for case-sensitive ATS identifiers.

**Patch:** lowercase only scheme/hostname and query parameter *names*. Preserve path case and identity parameter values unless a specific provider's documented canonicalization says otherwise. Keep the two canonicalizer implementations in one shared helper so they cannot drift. Add regression tests for case-sensitive identifiers.

### P2 — Firecrawl can label jobs with the wrong location

**Location:** `scraper.py:412–421`.

Every Firecrawl result is assigned `"location": location`, where `location` is the requested search location. The code does not extract the actual work location from the result or description. Search relevance does not prove that every result is in that city, so the exported tracker can state Hyderabad even if the listing describes another city, multiple locations, or remote work.

**Patch:** distinguish `search_location` from `job_location`; parse the page's actual location if available, otherwise use `Location unverified` or the query city with a clearly marked inferred value. Never present the query term as verified listing metadata.

### P2 — Freshness controls are not consistent between engines

**Location:** `config.yaml:20–23`; `scraper.py:347–353`, `scraper.py:704–720`, `scraper.py:896–914`.

`hours_old` is passed to JobSpy, but Firecrawl always uses `tbs: qdr:w` (past week). Naukri dates are then filtered with a fixed 14-day cutoff rather than the configured age. Thus a 24-hour setting does not mean a 24-hour freshness window across the whole pipeline.

**Patch:** apply the configured freshness policy to Firecrawl results too. Where a reliable posting timestamp exists, filter by the configured age. Where date data is unavailable, keep the record as date-unverified or make that retention policy explicit. Derive `tbs` from the configuration where practical, but don't depend on a search-engine time filter as the only date validation.

### P2 — A zero-result run leaves the previous workbook looking current

**Location:** `main.py`, the `if enriched_jobs_df.empty` branch after run logging; `scraper.py:311–314`.

If the result frame is empty, `main.py` prints a message and exits before `dashboard_gen.generate(...)` is called. `Tracker_Latest.xlsx` is therefore not refreshed, and an opened tracker may still display old leads after a no-result run. The scraper also catches per-role errors and can return empty data, so a source failure can look similar to a legitimate zero-match run. The process exits with status 0 in this branch.

**Patch:** track source request successes/failures separately from the number of jobs retained. For a *successful search with zero eligible leads*, generate an empty tracker/Run Summary that states zero matches. If all sources failed, preserve the last good tracker, record a failed run, and return a non-zero process exit code. Don't overwrite a good snapshot based on a network outage.

### P2 — README describes an IT Support pipeline that is missing from this ZIP

**Location:** `README.md` sections on sharing, Quick Start, IT Support pipeline, and architecture.

The archive does not contain `config_it_support.yaml`, `main_it_support.py`, or `run_it_support.bat`, although the README tells users to run them and lists them as project files. For someone following the README literally, that workflow is not runnable from this package.

**Patch:** include the missing files in the release archive, or remove/update the claims and links so the README describes only the delivered pipeline.

## Additional defects and worthwhile refinements

| Priority | Finding | Location | Recommended action |
|---|---|---|---|
| P2 | Firecrawl marks unknown posting dates as `Recent` for non-Naukri results. | `scraper.py:412–421`, `scraper.py:945–953` | Represent actual date, inferred date, and unknown date separately. Do not use “Recent” as a fallback without evidence. |
| P2 | The `output.filename` YAML setting is not used as the default write destination. | `main.py:272–287`, `excel_generator.py:283–305` | Either remove/rename the setting to show it is only a legacy sync candidate, or implement the documented behavior without breaking dual-output generation. |
| P2 | Archive filenames have minute-only precision, so two default runs in the same minute can overwrite the same archive. | `excel_generator.py:296–303` | Include seconds or a unique suffix in the archive name. |
| P2 | Run-log “credits used” is a team balance delta, not an attribution guaranteed to belong only to this process. | `main.py` pre-/post-flight credit checks; Firecrawl response includes `creditsUsed` | Sum per-request `creditsUsed` where available and retain account balance delta as a cross-check. Concurrent use of the same API key can distort the current delta. |
| P2 | Run Summary doesn't account for all funnel losses. | `scraper.py:292–305`, `scraper.py:378–425`, `scraper.py:317–330`, `excel_generator.py:513–520` | Add counters for invalid URLs, invalid/missing company, URL duplicates, cross-platform duplicates, scrape errors and date-unverified results. Label counts by stage so totals reconcile. |
| P2 | No explicit retry/backoff for transient Firecrawl errors or 429/5xx responses. | `scraper.py:358–432`; JobSpy exceptions at `scraper.py:280–309` | Add bounded retries with exponential backoff for retryable errors only; honor `Retry-After`; report permanent auth/credit failures distinctly. |
| P3 | `--hours 0` is ignored because the override uses `if args.hours:`. | `main.py:181–183` | Change to `if args.hours is not None:` and validate negative values. |
| P3 | `--dry-run` says “configuration validated successfully” but does not perform schema/type validation. | `main.py:165–215` | Validate required keys, list types, mode membership, nonnegative age/limits and `max_experience_years` before claiming validation. |
| P3 | Work mode is inferred from any occurrence of `remote`, `wfh` or `hybrid` in a long description. | `extractor.py:118–129` | Prefer explicit work-arrangement phrases and distinguish inferred from verified work mode. |
| P3 | Senior-role detection can over-filter titles containing terms like “Lead Generation” even when “lead” isn't a seniority marker. | `scraper.py:773–790` | Test common non-seniority contexts and use title-aware phrases/metadata instead of an unconditional token gate. |
| P3 | `FIRECRAWL_API_KEY=your_firecrawl_api_key_here` is a non-empty placeholder and is treated as a real key. | `.env.example`, `scraper.py:115`, `scraper.py:240–245` | Detect the placeholder and treat it as missing, or leave the template value blank. |
| P3 | Multiple parallel/manual/scheduled runs can race on shared JSON ledgers and the CSV log. Atomic replacement helps avoid partial files but does not prevent lost updates. | `scraper.py:175–197`, `excel_generator.py:101–123`, `main.py:20–103` | Add a process lock or explicitly prevent overlapping runs. This matters if manual runs can overlap Task Scheduler. |

## What is already strong

- The code has a clean top-level pipeline and isolates extraction, scraping, and Excel output into different modules.
- URL allowlisting is fail-closed for unknown domains and uses host-boundary checks rather than simple suffix substring checks.
- Cross-run state and application-ledger writes use temporary files plus `os.replace`.
- The application tracker supports status/note carryover, historical-file precedence, and reverting a status; existing tests exercise these behaviors.
- Spreadsheet cell values are sanitized against basic formula injection; apply hyperlinks are only assigned for HTTP(S) URLs.
- The run-log schema migration preserves unknown schemas rather than silently rewriting them.
- Scraping modes make cost/coverage trade-offs explicit. Firecrawl's current public pricing documentation supports the stated estimate of 2 credits per 10 search results plus 1 basic scrape credit per page, so the 8 / 24 / 48 credit estimates are internally consistent for 4 roles and limits of 10 / 4 / 10 respectively. Actual account usage should still be logged per request.

## Proposed patch order

### Patch 1 — Correct identity and eligibility semantics

1. Preserve case in canonical paths and ID values; share one canonicalizer implementation.
2. Replace boolean experience gating with explicit eligible/ineligible/unverified states.
3. Parse actual Firecrawl company/location/date metadata; never silently substitute the query location.

### Patch 2 — Make the end-to-end result trustworthy

1. Add source-domain config and implement Firecrawl coverage as advertised.
2. Fix company resolution before enabling ATS/enterprise domains.
3. Apply one freshness policy across both engines.
4. Differentiate a valid zero-result run from all-source failure; prevent a failed run from masquerading as a current empty result.

### Patch 3 — Improve operational robustness

1. Reconcile funnel metrics by stage and capture API errors/retries.
2. Use per-request Firecrawl `creditsUsed` alongside the account balance delta.
3. Add run locking, seconds-resolution archive names, and config validation.
4. Bring README and packaged files back into agreement.

### Patch 4 — Regression tests worth adding

- Case-sensitive URLs and identity query values do not collide.
- Valid Greenhouse/Workday/corporate URLs can acquire a company name and survive ingestion.
- Search location is not written as verified job location when the listing disagrees.
- Unknown experience is marked unverified rather than silently treated as verified eligible.
- A 24-hour setting cannot silently accept week-old Firecrawl listings when dates are available.
- A zero-match successful run updates the current-run summary; all-source failure preserves the last good tracker and exits nonzero.
- Two runs within one minute produce distinct archive files.
- `--hours 0` is applied and invalid negative ages are rejected.
- Invalid API key, HTTP 429, HTTP 500, and depleted credits are distinct in run analytics.

## External API verification

The current Firecrawl Search endpoint documentation lists `POST https://api.firecrawl.dev/v2/search`, supports `limit`, `tbs`, `includeDomains`, and `scrapeOptions`, and documents results under `data.web`: https://docs.firecrawl.dev/api-reference/endpoint/search

The current pricing page says Search costs 2 credits per 10 results, rounded up, with standard scrape credits charged per page when `scrapeOptions` requests page content: https://www.firecrawl.dev/pricing

The code's use of the v2 search endpoint and `scrapeOptions: {"formats": ["markdown"]}` matches the documented interface. The bigger Firecrawl problems are source-domain scope, downstream metadata rejection, freshness semantics, and metrics attribution—not the basic request shape.

## Final verdict

**Keep the architecture; patch the data semantics before adding features.** There is no need for a rewrite. The highest payoff comes from ensuring the pipeline never confuses *unknown* with *verified*, never invents listing metadata from the search query, preserves distinct job identities, and does not leave a stale workbook looking current after a failed run.
