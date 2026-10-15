import re
from typing import List, Dict, Any


class JobRequirementExtractor:
    """
    Extracts key skills, years of experience, and employment characteristics
    from raw job descriptions.
    """

    def __init__(self, target_skills: List[str] = None):
        # Default taxonomy if none provided
        self.target_skills = target_skills or [
            "SQL", "Power BI", "Python", "Excel", "Advanced Excel", "Tableau",
            "Data Annotation", "Data Cleaning", "ETL", "PostgreSQL", "MySQL",
            "Pandas", "NumPy", "Statistics", "Data Modeling", "DAX", "Power Query",
            "Machine Learning", "AI", "Azure", "AWS", "Git", "Snowflake", "BigQuery"
        ]

        # Build regex patterns for accurate matching
        self.compiled_skills = []
        for skill in self.target_skills:
            escaped = re.escape(skill)
            if skill in ["C++", "C#", ".NET"]:
                pattern = re.compile(rf'(?:^|[\s,;/]){escaped}(?:[\s,;/]|$)', re.IGNORECASE)
            elif skill.strip().lower() == "excel":
                # Avoid matching English verb usage: "to excel", "will excel in", "excel at", etc.
                # Matches uppercase/titlecase Excel/EXCEL or lowercase excel when not followed by prepositions
                pattern = re.compile(
                    r'(?<!\bto\s)(?<!\bwill\s)(?<!\bcan\s)(?<!\bwho\s)(?<!\byou\s)(?<!\bthey\s)(?<!\bmust\s)(?<!\bshould\s)'
                    r'\b(?:Excel\b|EXCEL\b|excel\b(?!\s+(?:in|at|within|under|when|as|by)\b))'
                )
            elif " " in skill:
                # Multi-word skills like "Power BI", "Data Annotation"
                pattern = re.compile(rf'\b{escaped}\b', re.IGNORECASE)
            else:
                pattern = re.compile(rf'\b{escaped}\b', re.IGNORECASE)
            self.compiled_skills.append((skill, pattern))

    def extract_skills(self, description: str) -> List[str]:
        """Extract matching skill keywords from description text."""
        if not description or not isinstance(description, str):
            return []

        matched_skills = []
        for original_name, pattern in self.compiled_skills:
            if pattern.search(description):
                matched_skills.append(original_name)

        return matched_skills

    COMPANY_SUBJECT_MARKERS = [
        "we have", "our company", "with over", "with more than",
        "history of", "legacy of", "celebrating", "founded in", "established in",
        "in business for", "operating for", "in the last", "in the past",
        "over the last", "over the past", "for the last", "for the past",
        "contract of", "bond of", "duration of", "tenure of", "notice period of"
    ]
    COMPANY_POST_MARKERS = [
        "in the industry", "in industry", "in business", "of excellence", "excellence",
        "serving clients", "track record", "market leader"
    ]

    NON_EXPERIENCE_POST_RE = re.compile(
        r'^\s*(?:of\s+)?(?:bachelor|master|degree|b\.?\s*tech|m\.?\s*tech|b\.?\s*e\b|b\.?\s*sc|bca|mca|mba|'
        r'diploma|graduation|graduate\s+program|college|university|education|'
        r'contract|bond|service\s+agreement|internship|training|probation|notice\s+period|duration|warranty)\b',
        re.IGNORECASE
    )

    EXPLICIT_EXP_POST_RE = re.compile(
        r'^\s*(?:[\'’]s?\s*)?(?:of\s+)?(?:relevant\s+|hands[-\s]on\s+|work\s+|industry\s+|professional\s+|total\s+|prior\s+|minimum\s+)?exp(?:erience)?\b',
        re.IGNORECASE
    )

    NEGATIVE_FRESHER_PATTERN = re.compile(
        r'\b(?:no\s+freshers?|freshers?\s+(?:need\s+not|not|do\s+not|should\s+not)\s+apply|freshers?\s+please\s+excuse)\b',
        re.IGNORECASE
    )

    def extract_experience(self, description: str) -> str:
        """Extract experience requirement (e.g. '0-2 years', '3–5 YOE', 'Fresher', 'Entry-Level')."""
        if not description or not isinstance(description, str):
            return "Not Specified"

        desc_lower = description.lower()
        req_verbs = (
            r'(?:minimum|min|at least|requires?|requiring|requirement|must have|must possess|should have|'
            r'looking for|seeking|position requires|role requires|experience\s*[:\-–—])'
        )
        dash_or_to = r'(?:[-–—]|to)'
        yr_unit = r'(?:years?|yrs?|yoe)\b'
        mo_unit = r'(?:months?|mos?)\b'
        any_unit = r'(?:years?|yrs?|yoe|months?|mos?)\b'

        def is_non_experience_context(start_idx: int, end_idx: int, matched_text: str = "") -> bool:
            pre = desc_lower[max(0, start_idx - 35):start_idx]
            post = desc_lower[end_idx:min(len(desc_lower), end_idx + 35)]
            if any(s in pre for s in self.COMPANY_SUBJECT_MARKERS) or any(s in post for s in self.COMPANY_POST_MARKERS):
                return True
            if "exp" not in matched_text and not self.EXPLICIT_EXP_POST_RE.search(post):
                if self.NON_EXPERIENCE_POST_RE.search(post):
                    return True
            return False

        # Tier 1: Explicit candidate requirement expressions (prefixed by requirement verbs)
        req_prefixed_patterns = [
            rf'(?:{req_verbs}\s+(?:a\s+)?)(\d+(?:\.\d+)?\+?\s*{mo_unit}\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{yr_unit})',
            rf'(?:{req_verbs}\s+(?:a\s+)?)(\d+(?:\.\d+)?\+?\s*{yr_unit}\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{mo_unit})',
            rf'(?:{req_verbs}\s+(?:a\s+)?)(\d+(?:\.\d+)?\+?\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{any_unit})',
            rf'(?:{req_verbs}\s+(?:a\s+)?)(\d+(?:\.\d+)?\+?\s*{any_unit}(?:\s+(?:of\s+)?(?:relevant\s+)?exp(?:erience)?)?)',
        ]
        for p in req_prefixed_patterns:
            for m in re.finditer(p, desc_lower):
                if not is_non_experience_context(m.start(1), m.end(1), m.group(1)):
                    return m.group(1).strip().title()

        # Tier 2: Unprefixed ranges (ranges are candidate requirements unless attached to bond/contract/heritage)
        range_patterns = [
            rf'(\d+(?:\.\d+)?\+?\s*{mo_unit}\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{yr_unit})',
            rf'(\d+(?:\.\d+)?\+?\s*{yr_unit}\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{mo_unit})',
            rf'(\d+(?:\.\d+)?\+?\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{yr_unit})',
            rf'(\d+(?:\.\d+)?\+?\s*{dash_or_to}\s*\d+(?:\.\d+)?\+?\s*{mo_unit})',
        ]
        for p in range_patterns:
            for m in re.finditer(p, desc_lower):
                if not is_non_experience_context(m.start(), m.end(), m.group(1)):
                    return m.group(1).strip().title()

        # Tier 3: Unprefixed single/plus numbers (requires explicit exp/experience, '+' sign, or 'yoe')
        numeric_patterns = [
            rf'\b(\d+(?:\.\d+)?\+?\s*(?:years?|yrs?)\s+(?:in|of|with)\b)',
            rf'\b(\d+(?:\.\d+)?\+?\s*{any_unit}\s*(?:[\'’]s?\s*)?(?:of\s+)?(?:relevant\s+|hands[-\s]on\s+|work\s+|industry\s+|professional\s+|total\s+)?exp(?:erience)?)\b',
            rf'\b(\d+(?:\.\d+)?\+\s*{any_unit})',
            rf'\b(\d+(?:\.\d+)?\+?\s*yoe)\b',
        ]
        for p in numeric_patterns:
            for m in re.finditer(p, desc_lower):
                if not is_non_experience_context(m.start(), m.end(), m.group(1)):
                    return m.group(1).strip().title()

        # Tier 4: Fresher / Entry keywords (only when no valid numeric requirement was found)
        if not self.NEGATIVE_FRESHER_PATTERN.search(desc_lower):
            if re.search(r'\b(fresher|freshers|entry[-\s]?level|new grad|college graduate|intern|internship|trainee)\b', desc_lower):
                return "0 - 1 Yrs (Entry / Fresher)"

        return "Not Specified"

    def extract_work_mode(self, location: str, description: str) -> str:
        """Determine if job is Remote, Hybrid, or On-site."""
        combined_text = f"{location or ''} {description or ''}".lower()

        if "remote" in combined_text or "work from home" in combined_text or "wfh" in combined_text:
            return "Remote"
        elif "hybrid" in combined_text:
            return "Hybrid"
        elif location:
            return f"On-site ({location})"
        else:
            return "On-site"

    def process_job(self, job_data: Dict[str, Any]) -> Dict[str, Any]:
        """Enrich a single raw job dictionary with extracted metrics."""
        description = job_data.get("description", "")
        location = job_data.get("location", "")

        extracted_skills = self.extract_skills(description)
        experience = self.extract_experience(description)
        work_mode = self.extract_work_mode(location, description)

        return {
            "skills": ", ".join(extracted_skills) if extracted_skills else "General Technical Skills",
            "skills_count": len(extracted_skills),
            "experience": experience,
            "work_mode": work_mode
        }
