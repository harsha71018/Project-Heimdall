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
        "history of", "legacy of", "celebrating", "founded in",
        "in business for", "operating for"
    ]
    COMPANY_POST_MARKERS = [
        "in the industry", "in industry", "in business", "of excellence",
        "serving clients", "track record", "market leader"
    ]

    NEGATIVE_FRESHER_PATTERN = re.compile(
        r'\b(?:no\s+freshers?|freshers?\s+(?:need\s+not|not|do\s+not|should\s+not)\s+apply|freshers?\s+please\s+excuse)\b',
        re.IGNORECASE
    )

    def extract_experience(self, description: str) -> str:
        """Extract experience requirement (e.g. '0-2 years', 'Fresher', 'Entry-Level')."""
        if not description or not isinstance(description, str):
            return "Not Specified"

        desc_lower = description.lower()

        # Tier 1: Explicit candidate requirement expressions (prefixed by requirement verbs)
        req_prefixed_patterns = [
            r'(?:(?:minimum|min|at least|requires?|requiring|requirement|must have|should have|looking for|seeking|position requires|role requires)\s+)(\d+(?:\.\d+)?\s*(?:months?|mos?)\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:years?|yrs?))',
            r'(?:(?:minimum|min|at least|requires?|requiring|requirement|must have|should have|looking for|seeking|position requires|role requires)\s+)(\d+(?:\.\d+)?\s*(?:years?|yrs?)\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:months?|mos?))',
            r'(?:(?:minimum|min|at least|requires?|requiring|requirement|must have|should have|looking for|seeking|position requires|role requires)\s+)(\d+(?:\.\d+)?\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:years?|yrs?|months?|mos?))',
            r'(?:(?:minimum|min|at least|requires?|requiring|requirement|must have|should have|looking for|seeking|position requires|role requires)\s+)(\d+(?:\.\d+)?\+?\s*(?:years?|yrs?|months?|mos?)(?:\s+(?:of\s+)?(?:relevant\s+)?exp(?:erience)?)?)',
        ]
        for p in req_prefixed_patterns:
            m = re.search(p, desc_lower)
            if m:
                return m.group(1).strip().title()

        def is_company_heritage(start_idx: int, end_idx: int) -> bool:
            pre = desc_lower[max(0, start_idx - 35):start_idx]
            post = desc_lower[end_idx:min(len(desc_lower), end_idx + 35)]
            return any(s in pre for s in self.COMPANY_SUBJECT_MARKERS) or any(s in post for s in self.COMPANY_POST_MARKERS)

        # Tier 2: Unprefixed ranges (ranges are candidate requirements, not company age)
        range_patterns = [
            r'(\d+(?:\.\d+)?\s*(?:months?|mos?)\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:years?|yrs?))',
            r'(\d+(?:\.\d+)?\s*(?:years?|yrs?)\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:months?|mos?))',
            r'(\d+(?:\.\d+)?\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:years?|yrs?))',
            r'(\d+(?:\.\d+)?\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:months?|mos?))',
        ]
        for p in range_patterns:
            for m in re.finditer(p, desc_lower):
                if not is_company_heritage(m.start(), m.end()):
                    return m.group(1).strip().title()

        # Tier 3: Unprefixed single/plus numbers (evaluated BEFORE fresher keywords)
        numeric_patterns = [
            r'(\d+(?:\.\d+)?\+?\s*(?:years?|yrs?|months?|mos?)\s*(?:[\'’]s?\s*)?(?:of\s+)?(?:relevant\s+)?exp(?:erience)?)',
            r'(\d+(?:\.\d+)?\+?\s*(?:years?|yrs?|months?|mos?))'
        ]
        for p in numeric_patterns:
            for m in re.finditer(p, desc_lower):
                if not is_company_heritage(m.start(), m.end()):
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
            "skills": ", ".join(extracted_skills) if extracted_skills else "Data Analysis, Problem Solving",
            "skills_count": len(extracted_skills),
            "experience": experience,
            "work_mode": work_mode
        }
