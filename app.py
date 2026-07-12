import streamlit as st
from pathlib import Path
from docx import Document
import PyPDF2
import re
from collections import Counter
try:
    from textstat import flesch_reading_ease
    _TEXTSTAT_AVAILABLE = True
except Exception:
    _TEXTSTAT_AVAILABLE = False
    # Fallback implementation of Flesch Reading Ease (approximate)
    import string

    def _count_syllables(word: str) -> int:
        word = word.lower().strip().strip(string.punctuation)
        if len(word) <= 3:
            return 1
        vowels = "aeiouy"
        count = 0
        prev_was_vowel = False
        for ch in word:
            is_vowel = ch in vowels
            if is_vowel and not prev_was_vowel:
                count += 1
            prev_was_vowel = is_vowel
        # Subtract silent 'e'
        if word.endswith("e") and count > 1:
            count -= 1
        return max(1, count)

    def flesch_reading_ease(text: str) -> float:
        # Very small approximate implementation to avoid hard dependency on textstat
        if not text or not text.strip():
            return 0.0

        sentences = re.split(r'[.!?]+', text)
        sentences = [s.strip() for s in sentences if s.strip()]
        words = re.findall(r"\b\w+\b", text)
        if not words or not sentences:
            return 0.0

        total_words = len(words)
        total_sentences = max(1, len(sentences))
        total_syllables = sum(_count_syllables(w) for w in words)

        # Flesch Reading Ease formula
        asl = total_words / total_sentences
        asw = total_syllables / total_words
        score = 206.835 - (1.015 * asl) - (84.6 * asw)
        try:
            return float(score)
        except Exception:
            return 0.0
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import base64
import json
import time
import datetime
import io 
import tempfile
import os 

# Text-to-speech: try pyttsx3; fallback handled at runtime
_PYTTSX3_AVAILABLE = False
_GTTS_AVAILABLE = False

try:
    import pyttsx3
    _PYTTSX3_AVAILABLE = True
except Exception:
    # Fallback to gTTS if pyttsx3 isn't available
    try:
        from gtts import gTTS
        _GTTS_AVAILABLE = True
    except Exception:
        pass

# ==================================================
# Config & Session State Initialization
# ==================================================
st.set_page_config(
    page_title="Resume Insight Pro+",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize Session State for persistent data (FIX for TTS re-run)
if 'metric_class' not in st.session_state:
    st.session_state['metric_class'] = '' 
if 'tts_audio_bytes' not in st.session_state:
    st.session_state['tts_audio_bytes'] = None
if 'tts_format' not in st.session_state:
    st.session_state['tts_format'] = 'audio/mp3'

# ==================================================
# Inject Custom CSS (Placeholder for external file content)
# ==================================================
def load_css(file_name="style.css"):
    # Mock CSS loading since 'style.css' is not provided
    css_code = """
    /* Main Layout */
    .section-scores { 
        padding: 15px; 
        border-radius: 8px; 
        background-color: #f0f2f6; 
        margin-bottom: 20px;
    }
    .stSpinner > div {
        text-align: center;
    }

    /* Animated Score Chart CSS */
    .chart-container {
        display: flex;
        justify-content: center;
        align-items: center;
        overflow: hidden; /* Important for the sliding effect */
    }
    
    /* 1. Metric Flash Animation */
    @keyframes metric-flash {
        0% { background-color: transparent; }
        50% { background-color: #d1e7dd; } /* Light success color */
        100% { background-color: transparent; }
    }
    .metric-container {
        padding: 10px;
        border-radius: 8px;
        transition: background-color 0.3s ease;
    }
    .metric-flash {
        animation: metric-flash 0.8s ease-out 2; /* Flash twice */
    }

    /* 2. Custom Animated Progress Bar */
    .progress-bar-container {
        width: 100%;
        background-color: #e0e0e0;
        border-radius: 4px;
        margin-bottom: 10px;
        overflow: hidden;
    }
    .progress-bar-fill {
        height: 15px;
        background-color: #4CAF50; /* Default green */
        width: 0;
        transition: width 1.5s ease-out; /* Animation effect */
        border-radius: 4px;
    }
    .progress-bar-fill.warning {
        background-color: #FFC107;
    }
    .progress-bar-fill.danger {
        background-color: #F44336;
    }

    /* 3. Resume Preview Styling */
    .resume-preview-container {
        background: white;
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        padding: 30px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        max-width: 100%;
        font-family: 'Calibri', 'Arial', sans-serif;
        line-height: 1.5;
        color: #333;
    }

    .resume-preview-container h1 {
        font-size: 24px;
        margin-bottom: 5px;
        color: #2C3E50;
        font-weight: bold;
    }

    .resume-preview-container h2 {
        font-size: 14px;
        margin-top: 15px;
        margin-bottom: 8px;
        color: #2C3E50;
        border-bottom: 2px solid #2C3E50;
        padding-bottom: 3px;
        font-weight: bold;
    }

    .resume-preview-container h3 {
        font-size: 12px;
        margin-top: 8px;
        margin-bottom: 3px;
        color: #34495E;
        font-weight: bold;
    }

    .resume-preview-container p {
        margin: 3px 0;
        font-size: 12px;
    }

    .resume-preview-container ul {
        margin: 5px 0;
        padding-left: 20px;
        font-size: 12px;
    }

    .resume-preview-container li {
        margin: 2px 0;
    }

    .resume-content-text {
        white-space: pre-wrap;
        word-wrap: break-word;
        font-size: 12px;
        line-height: 1.6;
        font-family: 'Calibri', 'Arial', sans-serif;
    }

    .resume-preview-pdf {
        border: 1px solid #ddd;
        border-radius: 8px;
        overflow: hidden;
    }

    .preview-label {
        font-weight: bold;
        color: #2C3E50;
        margin-top: 15px;
        font-size: 14px;
    }
    """
    st.markdown(f"<style>{css_code}</style>", unsafe_allow_html=True)

load_css()

# ==================================================
# Paths
# ==================================================
temp_dir = Path("temp_uploads")
temp_dir.mkdir(exist_ok=True)
history_file = temp_dir / "history.json"
if not history_file.exists():
    history_file.write_text(json.dumps([]))

# ==================================================
# History Management Functions (FIXED: Robust Loading)
# ==================================================
def load_history():
    """Loads analysis history from the JSON file, handling corruption."""
    try:
        with open(history_file, 'r') as f:
            data = json.load(f)
            # Ensure the root object is a list
            if isinstance(data, list):
                return data
            return []
    except (json.JSONDecodeError, FileNotFoundError, IOError):
        # If file doesn't exist or is invalid JSON, return empty list
        return []
    except Exception:
        return []

def save_history(new_entry):
    """Adds a new entry to the history file."""
    history = load_history()
    # Limit history size to prevent file bloat
    if len(history) >= 20: 
        history = history[-19:]
    
    history.append(new_entry)
    
    try:
        with open(history_file, 'w') as f:
            json.dump(history, f, indent=4)
    except Exception as e:
        print(f"Error saving history: {e}")

# ==================================================
# Sidebar Navigation (Improved)
# ==================================================
st.sidebar.markdown("## 📌 Navigation")

page_selection = st.sidebar.radio(
    "",
    ["🏠 Home", "🤖 AI Analyzer", "📜 History"]
)

# Normalize page value (strip icons for logic)
if "Home" in page_selection:
    page = "Home"
elif "AI Analyzer" in page_selection:
    page = "AI Analyzer"
elif "History" in page_selection:
    page = "History"
else:
    page = "Home" # Default

# ==================================================
# Animated Score Chart
# ==================================================
def animated_score_chart(score: int, size=(2.8, 2.8), font_size=18, thickness=0.4):
    placeholder = st.empty()

    def score_color(val):
        if val < 50:
            return "#E74C3C"  # red
        elif val < 75:
            return "#F39C12"  # orange
        else:
            return "#2ECC71"  # green

    # Build sliding animation by drawing the pie to an image and moving it in via CSS transform
    steps = [i for i in range(0, score, max(1, score // 20))]
    steps.append(score)

    # Starting offset percentage (slide from left)
    start_offset = -40  # pixels
    for idx, i in enumerate(steps):
        fig, ax = plt.subplots(figsize=size)
        color = score_color(i)
        ax.pie(
            [i, 100 - i],
            colors=[color, "#ECECEC"],
            startangle=90,
            wedgeprops={"width": thickness, "edgecolor": "white"},
            counterclock=False
        )
        ax.text(
            0, 0, f"{i}%", ha="center", va="center",
            fontsize=font_size, fontweight="bold", color="#2C3E50"
        )
        ax.axis("equal")

        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', transparent=True)
        plt.close(fig)
        buf.seek(0)
        img_b64 = base64.b64encode(buf.read()).decode('utf-8')

        # compute current offset easing from start_offset to 0
        progress = (idx + 1) / len(steps)
        offset_px = int(start_offset * (1 - progress))

        html = f"""
        <div class='chart-container' style='width:{int(size[0]*100)}px;'>
            <img src='data:image/png;base64,{img_b64}' style='transform: translateX({offset_px}px); transition: transform 0.18s ease-out; max-width:100%; height:auto;' />
        </div>
        """
        placeholder.markdown(html, unsafe_allow_html=True)
        time.sleep(0.02)

    # Final frame larger display
    fig, ax = plt.subplots(figsize=(4, 4))
    color = score_color(score)
    ax.pie(
        [score, 100 - score],
        colors=[color, "#ECECEC"],
        startangle=90,
        wedgeprops={"width": 0.35, "edgecolor": "white"},
        counterclock=False
    )
    ax.text(
        0, 0, f"{score}%", ha="center", va="center",
        fontsize=26, fontweight="bold", color="#2C3E50"
    )
    ax.axis("equal")

    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode('utf-8')
    final_html = f"""
    <div class='chart-container' style='width:400px;'>
        <img src='data:image/png;base64,{img_b64}' style='transform: translateX(0px); transition: transform 0.35s ease-out; max-width:100%; height:auto;' />
    </div>
    """
    placeholder.markdown(final_html, unsafe_allow_html=True)


# ==================================================
# Custom Animated Progress Bar Function
# ==================================================
def animated_progress_bar(label, value):
    """Displays an animated progress bar using custom HTML/CSS."""
    # Determine color class based on value
    if value < 50:
        color_class = "danger"
    elif value < 75:
        color_class = "warning"
    else:
        color_class = "" # Default green
        
    html = f"""
    <div style="font-weight: bold; margin-top: 15px;">{label} ({value}%)</div>
    <div class="progress-bar-container">
        <div class="progress-bar-fill {color_class}" style="width: {value}%;"></div>
    </div>
    """
    # The animation relies on the CSS transition property
    st.markdown(html, unsafe_allow_html=True)

# ==================================================
# Text Extraction and Document Display
# ==================================================
@st.cache_data
def extract_text(uploaded_file): 
    """Extracts text from a PDF or DOCX file."""
    
    ext = uploaded_file.name.split(".")[-1].lower() 
    content = ""
    
    uploaded_file.seek(0) 
    file_bytes = uploaded_file.read() 
    
    try:
        file_buffer = io.BytesIO(file_bytes)
        
        if ext == "pdf":
            pdf = PyPDF2.PdfReader(file_buffer)
            content = "\n".join([p.extract_text() or "" for p in pdf.pages])
        elif ext == "docx":
            doc = Document(file_buffer) 
            content = "\n".join([p.text for p in doc.paragraphs])
        else:
            st.error("Unsupported file type.")
            return ""
            
    except Exception as e:
        st.error(f"Failed to parse file. It might be corrupt or password-protected: {e}")
        return ""
        
    return content.lower()


# ==================================================
# Text-to-Speech Function
# ==================================================
def _speak_with_modulation(text, rate: int = 150, volume: float = 1.0, voice_id: str = None):
    """Generate TTS audio bytes with modulation. Uses pyttsx3 when available, otherwise gTTS fallback.

    Returns bytes (WAV or MP3) or None on failure.
    """
    if not text:
        return None

    # Try pyttsx3 first (allows rate/volume/voice alterations and WAV output on most platforms)
    if _PYTTSX3_AVAILABLE:
        try:
            engine = pyttsx3.init()
            
            try:
                engine.setProperty('rate', int(rate))
                engine.setProperty('volume', float(volume))
            except Exception:
                pass
            if voice_id:
                try:
                    engine.setProperty('voice', voice_id)
                except Exception:
                    pass

            fd, tmp_path = tempfile.mkstemp(suffix='.wav', prefix='resume_tts_')
            os.close(fd)
            audio_bytes = None
            try:
                engine.save_to_file(text, tmp_path)
                engine.runAndWait() # This blocks until the audio file is written
                
                if os.path.exists(tmp_path): 
                    with open(tmp_path, 'rb') as f:
                        audio_bytes = f.read()
            finally:
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
            
            if audio_bytes:
                return audio_bytes
        except Exception as e:
            print(f'pyttsx3 error: {e}')

    if _GTTS_AVAILABLE:
        try:
            from gtts import gTTS # Re-import locally for clarity
            tts = gTTS(text)
            buf = io.BytesIO()
            tts.write_to_fp(buf)
            buf.seek(0)
            return buf.read()
        except Exception as e:
            print(f'gTTS error: {e}')

    return None

# Function to display the file
def display_resume_preview(uploaded_file, resume_text=""):
    """Display resume with styled formatting for a professional appearance."""
    ext = uploaded_file.name.split(".")[-1].lower()
    
    uploaded_file.seek(0)
    
    if ext == "pdf":
        # PDF display with styled container
        st.markdown("### 📄 PDF Resume Preview")
        base64_pdf = base64.b64encode(uploaded_file.read()).decode('utf-8')
        pdf_display = f'<div class="resume-preview-pdf"><iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe></div>'
        st.markdown(pdf_display, unsafe_allow_html=True)
        
    elif ext == "docx":
        # DOCX display with styled text formatting
        st.markdown("### 📄 DOCX Resume Preview")
        if resume_text:
            # Create a styled text display
            st.markdown(f'<div class="resume-preview-container"><div class="resume-content-text">{resume_text[:3000]}</div></div>', 
                       unsafe_allow_html=True)
        else:
            st.write("DOCX file uploaded. Extracting content...")
    else:
        st.warning("Preview not available for this file type.")

def display_file(uploaded_file):
    ext = uploaded_file.name.split(".")[-1].lower()
    
    uploaded_file.seek(0)
    
    if ext == "pdf":
        base64_pdf = base64.b64encode(uploaded_file.read()).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
    elif ext == "docx":
        st.write("DOCX file uploaded. Displaying as text preview as a full document preview is not supported.")
        uploaded_file.seek(0)
        preview_text = extract_text(uploaded_file)
        st.info("DOCX file content: " + preview_text[:1000] + "...")
    else:
        st.warning("Preview not available for this file type.")


# ==================================================
# Analyzer Logic & Constants
# ==================================================
# Centralized Constants
CONSTANTS = {
    "SKILLS": ["python", "java", "sql", "excel", "aws", "tableau", "spark", "pandas", "numpy", "hadoop", "c++", "javascript", "react", "docker"],
    "KEYWORDS": ["machine learning", "data analysis", "communication", "leadership", "teamwork", "project management", "problem solving", "adaptability"],
    "DEGREES": ["bachelor", "master", "phd", "mba"],
    "ACTION_VERBS": ["managed", "developed", "led", "created", "designed", "implemented", "analyzed", "optimized", "collaborated", "achieved"]
}


def analyze(text, job_description_text=""):
    if not text:
        return {}
        
    # --- Structural Analysis ---
    
    # Check for contact info
    contact_info = {
        "email": bool(re.search(r'\S+@\S+', text)),
        "phone": bool(re.search(r'\+?\d{1,4}?[-.\s]?\(?\d{1,3}?\)?[-.\s]?\d{1,4}[-.\s]?\d{1,4}[-.\s]?\d{1,9}', text)),
        "linkedin": bool(re.search(r'linkedin\.com\/in\/', text)),
        "github": bool(re.search(r'github\.com\/', text))
    }
    
    # Section Score Calculation
    sections = {
        "Experience": 100 if "experience" in text or "project" in text else 50,
        "Education": 100 if any(d in text for d in CONSTANTS["DEGREES"]) else 50,
        # Score based on number of technical skills found
        "Skills": min(100, int(sum(1 for s in CONSTANTS["SKILLS"] if s in text) * (100 / len(CONSTANTS["SKILLS"])))), 
        "Projects": 100 if "project" in text else 50
    }
    core_score = int(sum(sections.values()) / len(sections))

    # Action Verbs and Quantifiable Metrics
    num_metrics = len(re.findall(r'(\d+%|(\$|€|£)\d+|(\d+\s+million)|(\d+\s+thousand))', text))
    action_verb_count = sum(1 for verb in CONSTANTS["ACTION_VERBS"] if verb in text)
    
    # --- Keyword Matching ---
    match_score = 0
    jd_keywords_found = {}
    if job_description_text:
        jd_text_lower = job_description_text.lower()
        # Combine predefined skills/keywords with 5 most frequent non-stop words from JD
        all_jd_terms = set(CONSTANTS["SKILLS"] + CONSTANTS["KEYWORDS"])
        
        # Simple tokenization for JD: find words of length > 3
        words = re.findall(r'\b[a-z]{4,}\b', jd_text_lower)
        word_counts = Counter(words)
        
        # Add top 5 unique words from JD
        for word, count in word_counts.most_common(5):
            all_jd_terms.add(word)
             
        total_unique_jd_terms = len(all_jd_terms)
        matched_terms = 0
        
        for term in all_jd_terms:
            if term in text:
                matched_terms += 1
                jd_keywords_found[term] = True
            
        if total_unique_jd_terms > 0:
            match_score = int((matched_terms / total_unique_jd_terms) * 100)
        
        # Adjust overall score slightly by the match score (70% core, 30% match)
        final_score = int((core_score * 0.7) + (match_score * 0.3))
    else:
        final_score = core_score
        match_score = 0
    
    # --- Dynamic Feedback Generation ---
    feedback = []
    if not contact_info["email"] or not contact_info["phone"]:
        feedback.append("Ensure your resume has clear contact information (email & phone).")
    if not contact_info["linkedin"]:
        feedback.append("Add a link to your LinkedIn profile to showcase your professional network.")
    if num_metrics < 3:
        feedback.append("Use more quantifiable metrics (e.g., 'Increased sales by 15%', 'Managed a $1M budget').")
    if action_verb_count < 5:
        feedback.append("Start your bullet points with strong action verbs to show impact.")

    missing_skills = [s for s in CONSTANTS["SKILLS"] if s not in text]
    missing_keywords = [k for k in CONSTANTS["KEYWORDS"] if k not in text]
    
    if sections["Skills"] < 50 and missing_skills:
        feedback.append(f"Consider adding technical skills like: {', '.join(missing_skills[:5])}")
    if missing_keywords:
        feedback.append(f"Include soft skills and keywords relevant to your field: {', '.join(missing_keywords[:3])}")
    if sections["Education"] < 100:
        feedback.append("Clearly mention your degrees and institutions.")
    if sections["Experience"] < 100:
        feedback.append("Add more detail on work/project experience, highlighting achievements.")
        
    if job_description_text and match_score < 70:
        feedback.append(f"Your resume match is **{match_score}%**. You are likely missing job-specific keywords. Tailor your language!")
        
    if not feedback:
        feedback.append("Your resume looks great! It's well-structured and impactful.")


    exp_years = sum(int(y) for y in re.findall(r'(\d+)\s+years?', text))
    readability = int(flesch_reading_ease(text)) if text.strip() else 0
    projects_count = len(re.findall(r'project', text))
    skill_counts = Counter([s for s in CONSTANTS["SKILLS"] if s in text])

    return {
        "score": final_score,
        "core_score": core_score,
        "match_score": match_score, 
        "sections": sections,
        "feedback": feedback,
        "exp_years": exp_years,
        "readability": readability,
        "projects_count": projects_count,
        "skill_counts": skill_counts,
        "num_metrics": num_metrics,
        "action_verb_count": action_verb_count,
        "jd_keywords_found": jd_keywords_found 
    }

# --- AI Mock Function (Placeholder) ---
def mock_ai_analysis(resume_text, job_description_text, result):
    if not job_description_text:
        return "I need a job description to provide personalized AI insights."

    match = result.get('match_score', 0)
    
    ai_feedback = f"Based on your resume and the provided Job Description, here is a deeper AI-driven assessment:\n\n"
    ai_feedback += f"### Match Confidence: {'High' if match >= 80 else ('Medium' if match >= 60 else 'Low')}\n"
    
    if match >= 80:
        ai_feedback += f"**Excellent fit!** Your resume strongly reflects the required skills and experience. Ensure your quantified achievements are the first thing recruiters see."
    elif match >= 60:
        ai_feedback += f"**Good potential.** Your core score is solid, but the match to the JD (**{match}%**) can be improved. Focus on incorporating the most frequent terms from the job description directly into your bullet points, especially if they relate to project management or core tools."
    else:
        ai_feedback += f"**Moderate fit.** To increase your chances, a significant revision is needed. The system detected a low keyword match (**{match}%**). Revamp your Experience section to use the specific technologies and responsibilities listed in the job description."

    return ai_feedback

# ==================================================
# Home Page
# ==================================================
if page == "Home":
    st.markdown(
        """
        <div style="text-align:center; padding:30px; background:linear-gradient(120deg, #74ebd5, #ACB6E5);
                    border-radius:15px; color:white; margin-bottom:25px;">
            <h1 style="font-size:40px; margin-bottom:10px;">📄 Resume Insight Pro+</h1>
            <p style="font-size:18px; line-height:1.6;">
                Your AI-powered companion to analyze, score, and enhance your resume instantly 🚀
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("Welcome to **Resume Insight Pro+**! 🎯 Upload your resume and let AI guide you towards career success.")
    st.markdown("Use the sidebar to navigate to **AI Analyzer** and get detailed insights on your resume.")

    st.markdown("---")
    st.subheader("✨ Key Features")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("📊 **Resume Score Calculation** – Get an instant score with a dynamic pie chart.")
        st.markdown("🧠 **Skills & Keyword Analysis** – Detect missing technical and job-relevant keywords.")
        st.markdown("💬 **Soft Skills Feedback** – Improve communication, teamwork, and leadership impact.")

    with col2:
        st.markdown("🎓 **Experience & Education Metrics** – Validate your academic and career journey.")
        st.markdown("☁️ **Skill Word Cloud** – Visualize top skills mentioned in your resume.")
        st.markdown("📜 **Resume History** – Track past uploads and measure your improvements.")
        st.markdown("📄 **Live Resume Preview** – See a clear front view of your uploaded resume.")
        st.markdown("🎯 **Job Description Match Score** – Tailor your resume for a specific job.")
        st.markdown("🔊 **Voice Summary** - Hear an audio summary of your results with modulation.")

    st.markdown("---")
    st.subheader("💡 How It Works")
    st.markdown(
        """
        1. 📂 Upload your resume (PDF/DOCX).
        2. 📝 **(NEW)** Paste the **Job Description** you are applying for.
        3. 🤖 AI extracts text and evaluates your resume against the JD.
        4. 📈 See an **animated score chart** and improvement areas, including a **Match Score**.
        5. 📝 Review personalized **feedback & suggestions**.
        6. 📊 Explore **history of previous uploads** and progress over time.
        """
    )

    st.success("👉 Head over to the **AI Analyzer** in the sidebar and upload your resume now!")

# ==================================================
# Analyzer Page (Enhanced)
# ==================================================
elif page == "AI Analyzer":
    st.markdown("<h1 style='text-align:center;color:#2C3E50;'>🔍 AI Resume Analyzer</h1>", unsafe_allow_html=True)
    
    # --- File Upload & JD Input ---
    
    col_upload, col_jd = st.columns(2)
    
    with col_upload:
        uploaded_resume = st.file_uploader("1. Upload your resume (PDF or DOCX)", type=["pdf", "docx"], key="resume_uploader")
    
    with col_jd:
        job_description_text = st.text_area(
            "2. Paste the **Job Description** here (Optional but Recommended for Match Score)", 
            height=200,
            placeholder="e.g., Responsibilities include developing Python scripts, managing AWS infrastructure, and collaborating with the Data Science team...",
            key="jd_text_area"
        )

    resume_text = ""
    file_details = {}
    
    if uploaded_resume:
        uploaded_resume.seek(0) 
        
        file_details = {
            "Filename": uploaded_resume.name,
            "File Size (KB)": round(len(uploaded_resume.getbuffer()) / 1024, 2),
            "Upload Time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        st.json(file_details)

        # Show resume preview prominently
        uploaded_resume.seek(0)
        
        with st.spinner("Extracting text from your resume..."):
            resume_text = extract_text(uploaded_resume)
        
        # Display resume with styled preview
        st.markdown("---")
        st.markdown("### 👀 Resume Front Display")
        display_resume_preview(uploaded_resume, resume_text)
        st.markdown("---")

    st.markdown("---")
    
    if st.button("🚀 Analyze Resume", type="primary", key="analyze_button"):
        # Reset TTS audio bytes when running new analysis
        st.session_state['tts_audio_bytes'] = None 
        
        if not uploaded_resume:
            st.error("Please upload a valid resume first!")
            st.stop()
        elif not resume_text:
            st.error("Text extraction failed. Please check the file format (must be a text-based PDF or DOCX).")
            st.stop()
        else:
            # Set a session state flag to trigger result display
            st.session_state['analysis_performed'] = True
            st.session_state['result'] = analyze(resume_text, job_description_text)
            st.session_state['file_details'] = file_details
            # Flag to trigger metric animation classes
            st.session_state['metric_class'] = 'metric-flash'
            
            # Save History immediately after analysis
            result = st.session_state['result']
            new_history_entry = {
                "file_name": uploaded_resume.name,
                "timestamp": file_details["Upload Time"],
                "overall_score": result["score"],
                "match_score": result["match_score"],
                "readability": result["readability"],
                "exp_years": result["exp_years"],
                "top_skills": dict(result["skill_counts"].most_common(5)),
                "feedback_summary": result["feedback"][0]
            }
            save_history(new_history_entry)
            st.toast("✅ Analysis saved to History!", icon='📜')
            st.rerun() # Rerun to display results clearly

    # --- Results Display (Conditional on analysis being performed) ---
    if st.session_state.get('analysis_performed', False):
        result = st.session_state['result']
        file_details = st.session_state['file_details']
        
        st.subheader("📊 Analysis Results")
        
        col_score, col_match = st.columns(2)
        
        with col_score:
            st.markdown("#### Overall Weighted Score")
            st.write("Structural compliance, skills, and JD match combined.")
            animated_score_chart(result["score"])
            st.info(f"Structural Score (Without JD Match): **{result['core_score']}%**")
            
        with col_match:
            if job_description_text:
                st.markdown("#### Job Description Match Score 🎯")
                st.write("Percentage of JD keywords found in your resume.")
                animated_score_chart(result["match_score"], size=(2.8, 2.8), font_size=18, thickness=0.4) 
            else:
                st.warning("Paste a Job Description above to get a **Match Score**!")
                
        st.markdown("---")

        # --- AI/LLM MOCK FEEDBACK ---
        st.subheader("🤖 AI-Driven Strategy & Feedback")
        with st.container(border=True):
             ai_feedback = mock_ai_analysis(resume_text, job_description_text, result)
             st.markdown(ai_feedback)
        
        st.markdown("---")

        # --- Feedback & Suggestions ---
        st.subheader("💬 Structural Feedback & Quick Wins")
        with st.container(border=True):
            for f in result["feedback"]:
                if "Ensure" in f or "Add" in f or "Consider" in f or "Include" in f or "missing job-specific" in f or "Revamp" in f:
                    st.warning(f)
                else:
                    st.success(f)
        
        st.markdown("---")
        
        # --- Section Scores and Metrics ---
        st.subheader("📂 Section-wise Scores & Key Metrics")
        
        col_metrics, col_sections = st.columns([1.5, 2])
        
        # 1. Metric Flash Animation
        metric_class = st.session_state.get('metric_class', '')
        
        with col_metrics:
            st.markdown("##### Resume Insights")
            
            st.markdown(f'<div class="metric-container {metric_class}">', unsafe_allow_html=True)
            st.metric("Experience Years", result['exp_years'])
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown(f'<div class="metric-container {metric_class}">', unsafe_allow_html=True)
            st.metric("Projects", result['projects_count'])
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown(f'<div class="metric-container {metric_class}">', unsafe_allow_html=True)
            st.metric("Readability (Flesch)", result['readability'], help="Aim for 60-70 for professional reading ease.")
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown(f'<div class="metric-container {metric_class}">', unsafe_allow_html=True)
            st.metric("Action Verbs", result['action_verb_count'])
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown(f'<div class="metric-container {metric_class}">', unsafe_allow_html=True)
            st.metric("Quantifiable Metrics", result['num_metrics'], help="Number of times currency, percentages, or large numbers are used to show impact.")
            st.markdown('</div>', unsafe_allow_html=True)


        with col_sections:
            st.markdown("##### Structural Completion (Animated)")
            st.markdown("<div class='section-scores'>", unsafe_allow_html=True)
            # 2. Custom Animated Progress Bars
            for k, v in result["sections"].items():
                animated_progress_bar(k, v)
            st.markdown("</div>", unsafe_allow_html=True)

        # Clear the flash class after use
        st.session_state['metric_class'] = '' 

        st.markdown("---")
        
        # --- Word Cloud ---
        if result["skill_counts"]:
            st.subheader("☁️ Skill Cloud")
            st.write("This shows the frequency of detected technical skills in your resume.")
            word_freqs = {k: v for k, v in result["skill_counts"].items()}
            
            wc = WordCloud(width=600, height=300, background_color="white").generate_from_frequencies(word_freqs)
            plt.figure(figsize=(8, 4))
            plt.imshow(wc, interpolation="bilinear")
            plt.axis("off")
            st.pyplot(plt)
            plt.close() 
            
        # --- Voice Modulation / Audio Summary (FIXED: Uses Session State) ---
        st.markdown("---")
        st.subheader("🔊 Voice Summary (Modulation)")
        
        if not _PYTTSX3_AVAILABLE and not _GTTS_AVAILABLE:
            st.error('No TTS engine available. Install `pyttsx3` (for local modulation) or `gTTS` (for online fallback).')
        else:
            col_v1, col_v2, col_v3 = st.columns([2,1,1])
            
            tts_rate = st.slider("Speech Rate (Words/Min)", 80, 250, 150, key='tts_rate')
            tts_volume = st.slider("Volume", 0.1, 1.0, 1.0, step=0.1, key='tts_volume')
            
            selected_voice = None
            if _PYTTSX3_AVAILABLE:
                try:
                    engine_tmp = pyttsx3.init()
                    voices = engine_tmp.getProperty('voices')
                    voice_options = {v.id: getattr(v, 'name', v.id) for v in voices}
                    
                    with col_v3: 
                         selected_voice = st.selectbox('Voice', options=list(voice_options.keys()), format_func=lambda x: voice_options.get(x, x), key='tts_voice_select')
                    engine_tmp.stop()
                except Exception:
                    pass
            
            summary_text = st.text_area('Optional: Edit the audio summary text', 
                                        value=(f"Your resume scored {result['score']} out of 100. Readability is {result['readability']}. ATS match is {result['match_score']} percent."), 
                                        height=120, key='tts_summary_text')

            # --- TTS Button Logic: Store audio bytes in session state ---
            if st.button('▶️ Play Voice Summary', key='tts_button_play'):
                with st.spinner('Generating audio...'):
                    # The function returns audio bytes
                    audio_bytes = _speak_with_modulation(summary_text, rate=tts_rate, volume=tts_volume, voice_id=selected_voice)
                    
                    if audio_bytes:
                        # Store audio bytes and format in session state
                        st.session_state['tts_audio_bytes'] = audio_bytes
                        st.session_state['tts_format'] = 'audio/wav' if _PYTTSX3_AVAILABLE else 'audio/mp3'
                        st.success('Audio generated. Playing now.')
                        st.rerun() # Rerun to display the audio widget outside the button block
                    else:
                        st.error('Failed to generate audio. Check TTS installation or system configuration.')
                        st.session_state['tts_audio_bytes'] = None

            # --- TTS Audio Display: Persistent outside the button block ---
            if st.session_state['tts_audio_bytes'] is not None:
                st.audio(
                    st.session_state['tts_audio_bytes'], 
                    format=st.session_state['tts_format']
                )

# ==================================================
# History Page (FIXED: Robust Filtering)
# ==================================================
elif page == "History":
    st.markdown("<h1 style='text-align:center;color:#2C3E50;'>📜 Analysis History</h1>", unsafe_allow_html=True)
    
    history_data = load_history()
    
    # Define required keys for a valid history entry
    required_keys = ["timestamp", "overall_score", "match_score", "readability", "file_name", "top_skills", "feedback_summary"]
    
    # Filter out corrupted or incomplete entries (Fixes KeyError: 'timestamp')
    valid_history = [
        entry for entry in history_data 
        if all(key in entry for key in required_keys)
    ]
    
    if not history_data:
        st.info("No analysis history found. Upload and analyze a resume on the 'AI Analyzer' page to start tracking your progress.")
    elif not valid_history:
        st.warning("History file found, but no **valid** analysis entries could be processed for charting. Please upload a new resume to create a valid entry, or manually inspect/clear `temp_uploads/history.json`.")
    else:
        st.subheader("Progress Over Time")

        # Prepare data for line charts using *only* valid_history
        try:
            import pandas as pd
            
            timestamps = [entry["timestamp"] for entry in valid_history]
            overall_scores = [entry["overall_score"] for entry in valid_history]
            match_scores = [entry["match_score"] for entry in valid_history]
            readability_scores = [entry["readability"] for entry in valid_history]
            
            chart_data = pd.DataFrame({
                'Date': pd.to_datetime(timestamps),
                'Overall Score': overall_scores,
                'Match Score': match_scores,
                'Readability': readability_scores
            })

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.line_chart(chart_data, x='Date', y=['Overall Score', 'Match Score'], use_container_width=True)
            
            with col_chart2:
                st.line_chart(chart_data, x='Date', y='Readability', use_container_width=True, color="#F39C12")

        except ImportError:
            st.warning("Install `pandas` for advanced history charts.")
        except Exception as e:
            st.error(f"Error generating charts: {e}")

        st.markdown("---")
        st.subheader("Latest Analysis Entries")
        
        # Display history in reverse order (most recent first) using *only* valid_history
        for i, entry in enumerate(reversed(valid_history)):
            with st.expander(f"**{entry['timestamp']}** | File: **{entry['file_name']}** | Score: **{entry['overall_score']}%**", expanded=(i==0)):
                
                col_h1, col_h2, col_h3 = st.columns(3)
                
                with col_h1:
                    st.metric("Overall Score", f"{entry['overall_score']}%")
                with col_h2:
                    st.metric("Match Score", f"{entry['match_score']}%")
                with col_h3:
                    st.metric("Readability", entry['readability'])
                    
                st.caption(f"**Top Feedback:** {entry['feedback_summary']}")
                
                skills_display = entry['top_skills'].keys() if isinstance(entry['top_skills'], dict) else entry.get('top_skills', ['N/A'])
                st.text(f"Top Skills: {', '.join(skills_display)}")
        
        if st.button("Clear History", type="secondary"):
            history_file.write_text(json.dumps([]))
            st.success("History cleared!")
            st.rerun()