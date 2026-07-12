import re
import pdfplumber
from docx import Document
from pathlib import Path
from collections import defaultdict

# ------------------------------
# Resume Text Extraction
# ------------------------------
def extract_text(file_path):
    """
    Extracts text from PDF or DOCX file and returns as lowercase string.
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
# Extract Personal Information
# ------------------------------
EMAIL_REGEX = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}'
PHONE_REGEX = r'(\+?\d{1,3}[\s-]?)?(\(?\d{3}\)?[\s-]?)?\d{3}[\s-]?\d{4}'

def extract_contact_info(resume_text):
    email_matches = re.findall(EMAIL_REGEX, resume_text)
    phone_matches = re.findall(PHONE_REGEX, resume_text)
    phones = ["".join(filter(str.isdigit, p)) for p in phone_matches if any(c.isdigit() for c in p)]
    return {
        "emails": list(set(email_matches)),
        "phones": list(set(phones))
    }

# ------------------------------
# Extract Name (heuristic)
# ------------------------------
def extract_name(resume_text):
    """
    Tries to extract candidate name from the first lines of resume.
    """
    lines = [line.strip() for line in resume_text.splitlines() if line.strip()]
    for line in lines[:5]:  # First 5 lines likely contain name
        if all(x.isalpha() or x.isspace() for x in line) and 2 <= len(line.split()) <= 4:
            return line.title()
    return "Not Found"

# ------------------------------
# Extract Education
# ------------------------------
EDU_KEYWORDS = ["bachelor", "master", "phd", "b.sc", "m.sc", "mba", "degree", "university", "college"]

def extract_education(resume_text):
    edu_lines = []
    for line in resume_text.splitlines():
        if any(keyword in line for keyword in EDU_KEYWORDS):
            edu_lines.append(line.strip())
    return edu_lines

# ------------------------------
# Extract Work Experience
# ------------------------------
EXP_KEYWORDS = ["experience", "worked", "role", "employment", "internship", "project"]

def extract_experience(resume_text):
    exp_lines = []
    for line in resume_text.splitlines():
        if any(keyword in line for keyword in EXP_KEYWORDS):
            exp_lines.append(line.strip())
    return exp_lines

# ------------------------------
# Extract Skills
# ------------------------------
TECH_SKILLS = [
    "python", "java", "c++", "sql", "excel", "aws", "tableau", "powerbi",
    "r", "tensorflow", "pandas", "numpy", "react", "django", "flask",
    "javascript", "html", "css"
]

SOFT_SKILLS = [
    "communication", "teamwork", "leadership", "problem-solving",
    "creativity", "adaptability", "critical thinking"
]

def extract_skills(resume_text):
    skills_found = {"technical": [], "soft": []}
    for skill in TECH_SKILLS:
        if skill.lower() in resume_text:
            skills_found["technical"].append(skill)
    for skill in SOFT_SKILLS:
        if skill.lower() in resume_text:
            skills_found["soft"].append(skill)
    return skills_found

# ------------------------------
# Extract Projects (heuristic)
# ------------------------------
PROJECT_KEYWORDS = ["project", "research", "portfolio", "capstone", "thesis"]

def extract_projects(resume_text):
    project_lines = []
    for line in resume_text.splitlines():
        if any(keyword in line for keyword in PROJECT_KEYWORDS):
            project_lines.append(line.strip())
    return project_lines

# ------------------------------
# Parse Resume
# ------------------------------
def parse_resume(resume_path):
    """
    Returns structured information from resume:
    - name
    - contact info
    - education
    - experience
    - skills
    - projects
    """
    resume_text = extract_text(resume_path)
    
    parsed_data = {
        "name": extract_name(resume_text),
        "contact": extract_contact_info(resume_text),
        "education": extract_education(resume_text),
        "experience": extract_experience(resume_text),
        "skills": extract_skills(resume_text),
        "projects": extract_projects(resume_text)
    }
    
    return parsed_data

# ------------------------------
# Example usage
# ------------------------------
if __name__ == "__main__":
    resume_file = Path("sample_resume.pdf")  # Replace with your resume
    parsed = parse_resume(resume_file)
    import json
    print(json.dumps(parsed, indent=4))
