# resume_insight_pro_plus

Unified project: shared core/ logic, backend FastAPI service, and app.py Streamlit demo frontend.

## 🚀 About the Project
**Resume Insight Pro Plus** is an intelligent, full-stack application designed to act like an Applicant Tracking System (ATS). It parses professional resumes, compares them against industry standards or job requirements, and provides actionable, data-driven feedback along with an optimization score.

### Key Features
* **Multi-Format Parsing:** Seamlessly processes both `.pdf` and `.docx` resumes.
* **ATS Scoring Engine:** Evaluates resume structure, key sections, and critical formatting parameters to output a definitive score.
* **Keyword & Semantic Analytics:** Analyzes skills and industry keywords, generating a visual `keyword_chart.png` to map keyword density.
* **Streamlit Demo Frontend:** Interactive, user-friendly UI for uploading documents and visualizing real-time metrics dynamically.
* **FastAPI Backend Service:** High-performance API endpoints managing parsing logic, data validation, and historical logging records.

---

## 🛠️ Setup & Installation

1. Create virtualenv and activate:
   ```bash
   python -m venv venv
   source venv/bin/activate   # mac/linux
   venv\Scripts\activate      # windows