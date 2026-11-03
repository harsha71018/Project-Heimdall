"""
AI Decision Gatekeeper (Tier-2 Semantic Validation)
Evaluates candidate leads using OpenRouter's free tier (e.g. nvidia/nemotron-3-ultra-550b-a55b:free).
Fails open: if no key is set or API is unreachable, it smoothly logs a notice and returns the original DataFrame.
"""

import os
import json
import re
from typing import Dict, Any, List
import requests
import pandas as pd


def parse_json_loose(text: str) -> Any:
    """Parses JSON from model responses, stripping markdown wrappers if present."""
    if not text:
        return None
    cleaned = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.M).strip()
    cleaned = re.sub(r"```$", "", cleaned.strip(), flags=re.M).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        m = re.search(r"(\[.*\]|\{.*\})", text, flags=re.S)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                return None
    return None


class AIGatekeeper:
    """
    Tier-2 Semantic Validator.
    Takes candidate DataFrame and runs a single-batch classification call
    evaluating true fresher/entry-level eligibility and domain relevance.
    """

    def __init__(self, config: Dict[str, Any]):
        self.ai_cfg = config.get("ai_gatekeeper", {})
        self.enabled = self.ai_cfg.get("enabled", True)
        self.api_key = os.getenv("OPENROUTER_API_KEY", "").strip() or os.getenv("GEMINI_API_KEY", "").strip()
        self.provider = self.ai_cfg.get("provider", "openrouter").lower()
        self.model = self.ai_cfg.get("model", "nvidia/nemotron-3-ultra-550b-a55b:free")
        self.fallback_model = self.ai_cfg.get("fallback_model", "openrouter/free")
        self.min_fit_score = float(self.ai_cfg.get("min_fit_score", 6.0))
        self.target_titles = config.get("search", {}).get("job_titles", ["Data Analyst", "SQL Developer"])
        self.max_experience_years = config.get("search", {}).get("max_experience_years", 2)

    @property
    def is_active(self) -> bool:
        return bool(self.enabled and self.api_key)

    def filter_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df

        if not self.enabled:
            print("[*] [AI Gatekeeper] Disabled in config.yaml. Skipping Tier-2 semantic validation.")
            return df

        if not self.api_key:
            print("[*] [AI Gatekeeper] No OPENROUTER_API_KEY in .env. Tier-1 deterministic filters applied.")
            return df

        print(f"\n[*] [AI Gatekeeper] Evaluating {len(df)} candidate leads with {self.model}...")

        # Prepare batch prompt payload
        leads_payload = []
        for idx, row in df.iterrows():
            leads_payload.append({
                "id": int(idx),
                "title": str(row.get("title", "")),
                "company": str(row.get("company", "")),
                "experience": str(row.get("extracted_experience", "Not Specified")),
                "location": str(row.get("location", "")),
                "snippet": str(row.get("description", ""))[:400]
            })

        roles_str = ", ".join(self.target_titles)
        system_prompt = (
            "You are an expert technical recruiter specializing in entry-level (0-2 YOE) staffing in India.\n"
            f"Target roles wanted: {roles_str}.\n"
            f"Candidate ceiling: strictly <= {self.max_experience_years} years of experience.\n"
            "Evaluate each job listing:\n"
            "- is_fresher_eligible (bool): true if a fresher/junior (0-2 years) is genuinely eligible. false if the role asks for >2 years, is a Senior/Lead/Architect role, SME, or lateral hire.\n"
            f"- is_domain_match (bool): true if this role genuinely matches one of the target domains ({roles_str}). false if non-technical, sales, telecalling, customer support, or off-domain.\n"
            "- fit_score (float 1-10): rate match against entry-level target role.\n"
            "- reason (str): short 1-sentence explanation.\n\n"
            "Respond ONLY with a valid JSON object matching this schema:\n"
            '{"decisions": [{"id": 0, "is_fresher_eligible": true, "is_domain_match": true, "fit_score": 8.5, "reason": "..."}]}'
        )

        user_prompt = f"Jobs to evaluate:\n{json.dumps(leads_payload, indent=2)}"

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/Antigravity",
            "X-Title": "Project Heimdall Job Hunter"
        }

        # Try primary model, fallback if error
        models_to_try = [self.model]
        if self.fallback_model and self.fallback_model != self.model:
            models_to_try.append(self.fallback_model)

        raw_response_text = None
        for current_model in models_to_try:
            try:
                payload = {
                    "model": current_model,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    "temperature": 0.1,
                    "response_format": {"type": "json_object"}
                }
                res = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=35
                )
                if res.status_code == 200:
                    data = res.json()
                    choices = data.get("choices", [])
                    if choices:
                        raw_response_text = choices[0].get("message", {}).get("content", "")
                        break
                else:
                    print(f"   [!] OpenRouter HTTP {res.status_code} for {current_model}: {res.text[:120]}")
            except Exception as e:
                print(f"   [!] Error querying {current_model}: {e}")

        if not raw_response_text:
            print("[!] Warning: AI Gatekeeper failed all attempts. Gracefully keeping Tier-1 filtered leads.")
            return df

        parsed = parse_json_loose(raw_response_text)
        if not parsed or "decisions" not in parsed:
            print("[!] Warning: Could not parse AI Gatekeeper JSON response. Keeping Tier-1 leads.")
            return df

        decisions_map = {d["id"]: d for d in parsed.get("decisions", []) if "id" in d}

        keep_indices = []
        dropped_count = 0
        for idx in df.index:
            dec = decisions_map.get(int(idx))
            if not dec:
                # If specific ID was missing from response, fail-safe keep
                keep_indices.append(idx)
                continue

            is_fresher = dec.get("is_fresher_eligible", True)
            is_domain = dec.get("is_domain_match", True)
            score = float(dec.get("fit_score", 7.0))

            if is_fresher and is_domain and score >= self.min_fit_score:
                keep_indices.append(idx)
            else:
                dropped_count += 1

        retained_df = df.loc[keep_indices].copy()
        print(f"[+] [AI Gatekeeper] Retained {len(retained_df)} verified leads (Dropped {dropped_count} experienced/off-domain leads).")
        return retained_df
