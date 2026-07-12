import re
import pdfplumber
from docx import Document
from pathlib import Path
from collections import Counter, defaultdict
import json
import datetime

# ------------------------------
# Resume Text Extraction
# ------------------------------
def extract_text(file_path):
    """
    Extract text from PDF or DOCX resume.
    Returns lowercase string.
    """
    text = ""
    if file_path.suffix.lower() == ".pdf":
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    elif file_path.suffix.lower() == ".docx":
        doc = Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    return text.lower()

# ------------------------------
# Keyword Extraction
# ------------------------------
def extract_keywords(text, min_len=3):
    """
    Extract all words with min_len or more characters.
    """
    return [w for w in re.findall(r'\w+', text) if len(w) >= min_len]

# ------------------------------
# Section Detection
# ------------------------------
SECTION_HEADERS = {
    "education": ["education", "degree", "university", "college", "school"],
    "experience": ["experience", "work history", "employment", "career", "role"],
    "projects": ["project", "research", "internship", "portfolio"],
    "skills": ["skills", "technical skills", "competencies"],
    "certifications": ["certification", "certificate", "course"],
    "achievements": ["achievement", "award", "honor"],
    "summary": ["summary", "profile", "objective"],
    "hobbies": ["hobby", "interest", "extracurricular"]
}

def detect_sections(resume_text):
    """
    Returns detected sections with boolean flags.
    """
    sections_found = {}
    for sec, keywords in SECTION_HEADERS.items():
        sections_found[sec] = any(kw in resume_text for kw in keywords)
    return sections_found

# ------------------------------
# Skill Matching
# ------------------------------
TECHNICAL_SKILLS = [
    "python", "java", "c++", "sql", "excel", "aws", "tableau", "powerbi",
    "r", "tensorflow", "pandas", "numpy", "react", "django", "flask"
]

SOFT_SKILLS = [
    "communication", "teamwork", "leadership", "problem-solving", 
    "creativity", "adaptability", "critical thinking"
]

def score_skills(resume_text):
    """
    Returns technical and soft skills match.
    """
    tech_count = sum(1 for skill in TECHNICAL_SKILLS if skill in resume_text)
    soft_count = sum(1 for skill in SOFT_SKILLS if skill in resume_text)
    return {
        "technical": tech_count,
        "soft": soft_count,
        "technical_score": min(100, tech_count * 10),
        "soft_score": min(100, soft_count * 10)
    }

# ------------------------------
# Experience Analysis
# ------------------------------
def analyze_experience(resume_text):
    """
    Simple heuristic: Detect years of experience.
    """
    years = []
    pattern = r'(\d+)\s*(?:years|yrs|year)'
    matches = re.findall(pattern, resume_text)
    for m in matches:
        years.append(int(m))
    total_years = sum(years)
    exp_score = min(100, total_years * 10)  # Max 100%
    return {
        "years": total_years,
        "score": exp_score
    }

# ------------------------------
# Keyword Matching
# ------------------------------
def match_keywords(job_desc_text, resume_text):
    """
    Returns keyword match percentage per word.
    """
    job_keywords = extract_keywords(job_desc_text)
    resume_words = extract_keywords(resume_text)
    keyword_count = Counter(resume_words)
    match = {}
    for kw in set(job_keywords):
        count = keyword_count.get(kw, 0)
        match_pct = min(100, int((count / job_keywords.count(kw)) * 100)) \
                    if job_keywords.count(kw) else 0
        match[kw] = match_pct
    return match

# ------------------------------
# Formatting & Readability
# ------------------------------
def readability_score(resume_text):
    """
    Scores resume based on formatting heuristics.
    """
    lines = [line.strip() for line in resume_text.splitlines() if line.strip()]
    avg_line_len = sum(len(line) for line in lines) / max(1, len(lines))
    bullets = sum(1 for line in lines if line.startswith(("-", "*", "•")))
    headers = sum(1 for line in lines if line.isupper() and len(line.split()) < 6)
    
    score = 0
    if avg_line_len <= 120:
        score += 40
    if bullets >= 5:
        score += 30
    if headers >= 3:
        score += 30
    return min(100, score)

# ------------------------------
# Feedback Generation
# ------------------------------
def generate_feedback(sections_detected, skills_score, exp_score, keywords_match):
    feedback = []
    for sec, present in sections_detected.items():
        if not present:
            feedback.append(f"Add your {sec.capitalize()} section for better clarity.")
    
    if skills_score["technical_score"] < 50:
        feedback.append("Consider adding more technical skills.")
    if skills_score["soft_score"] < 50:
        feedback.append("Consider adding relevant soft skills.")
    if exp_score["score"] < 50:
        feedback.append("Experience section seems short; try to elaborate on roles.")
    
    for kw, pct in keywords_match.items():
        if pct < 50:
            feedback.append(f"Keyword '{kw}' appears less than 50% in your resume.")
    
    return feedback

# ------------------------------
# Overall Resume Analysis
# ------------------------------
def analyze_resume(resume_path, job_desc_text):
    """
    Full resume analysis returning:
    - overall score
    - section-wise scores
    - keyword matches
    - feedback
    """
    resume_text = extract_text(resume_path)
    
    # Sections
    sections_detected = detect_sections(resume_text)
    section_scores = {}
    for sec, present in sections_detected.items():
        section_scores[sec] = 100 if present else 0
    
    # Skills
    skills_score = score_skills(resume_text)
    section_scores["technical_skills"] = skills_score["technical_score"]
    section_scores["soft_skills"] = skills_score["soft_score"]
    
    # Experience
    exp_score = analyze_experience(resume_text)
    section_scores["experience_years"] = exp_score["score"]
    
    # Keywords
    keywords_match = match_keywords(job_desc_text.lower(), resume_text)
    
    # Readability
    section_scores["formatting"] = readability_score(resume_text)
    
    # Overall score
    overall_score = int(sum(section_scores.values()) / len(section_scores))
    
    # Feedback
    feedback = generate_feedback(sections_detected, skills_score, exp_score, keywords_match)
    
    # Return full structured analysis
    return {
        "score": overall_score,
        "section_scores": section_scores,
        "keywords": keywords_match,
        "feedback": feedback
    }

# ------------------------------
# Optional: Save feedback to JSON
# ------------------------------
def save_feedback_json(feedback_data, output_path=None):
    if not output_path:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path(f"resume_feedback_{timestamp}.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(feedback_data, f, indent=4)
    return output_path

# ------------------------------
# Example usage
# ------------------------------
if __name__ == "__main__":
    from pathlib import Path
    resume_file = Path("sample_resume.pdf")  # Replace with test file
    job_desc = "Data Engineering Python SQL AWS Tableau"
    result = analyze_resume(resume_file, job_desc)
    print(json.dumps(result, indent=4))
