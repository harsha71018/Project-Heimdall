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

    def extract_experience(self, description: str) -> str:
        """Extract experience requirement (e.g. '0-2 years', 'Fresher', 'Entry-Level')."""
        if not description or not isinstance(description, str):
            return "Not Specified"

        desc_lower = description.lower()

        # Check for fresher/intern terms
        if re.search(r'\b(fresher|freshers|entry[-\s]?level|new grad|college graduate|intern|internship|trainee)\b', desc_lower):
            return "0 - 1 Yrs (Entry / Fresher)"

        # Common experience patterns (e.g., "0 to 2 years", "1-3 yrs", "0-2 yrs", "minimum 1 year")
        patterns = [
            r'(\d+\s*(?:-|to)\s*\d+\s*(?:years?|yrs?))',
            r'(\d+\+?\s*(?:years?|yrs?)\s*(?:of\s*)?(?:relevant\s*)?experience)',
            r'((?:minimum|min|at least)\s*\d+\s*(?:years?|yrs?))',
            r'(\d+\+?\s*(?:years?|yrs?))'
        ]

        for p in patterns:
            match = re.search(p, desc_lower)
            if match:
                exp_str = match.group(1).strip()
                return exp_str.title()

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
