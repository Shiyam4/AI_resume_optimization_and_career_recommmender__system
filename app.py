import datetime as dt
import base64
import difflib
import html
import io
import json
import os
import re
import tempfile
from collections import Counter
from pathlib import Path

import matplotlib.pyplot as plt
import PyPDF2
import streamlit as st
from docx import Document
from wordcloud import WordCloud

try:
    import pyttsx3
    TTS_AVAILABLE = True
except Exception:
    TTS_AVAILABLE = False
try:
    from gtts import gTTS
    GTTS_AVAILABLE = True
except Exception:
    GTTS_AVAILABLE = False

try:
    from textstat import flesch_reading_ease
except Exception:
    def flesch_reading_ease(text):
        words = re.findall(r"\b\w+\b", text)
        sentences = [item for item in re.split(r"[.!?]+", text) if item.strip()]
        if not words or not sentences:
            return 0
        syllables = sum(max(1, len(re.findall(r"[aeiouy]+", word.lower()))) for word in words)
        return 206.835 - 1.015 * len(words) / len(sentences) - 84.6 * syllables / len(words)


st.set_page_config(page_title="Resume Insight", page_icon="✨", layout="wide", initial_sidebar_state="expanded")

SKILLS = ["python", "sql", "excel", "aws", "tableau", "power bi", "spark", "pandas", "numpy", "java", "c++", "javascript", "react", "docker", "git", "machine learning"]
ACTION_VERBS = ["achieved", "analyzed", "built", "collaborated", "created", "delivered", "designed", "developed", "improved", "implemented", "increased", "launched", "led", "managed", "optimized", "reduced"]
STOP_WORDS = {"about", "after", "against", "along", "also", "and", "are", "been", "business", "candidate", "candidates", "company", "data", "for", "from", "have", "into", "join", "looking", "our", "role", "skills", "team", "that", "the", "this", "through", "with", "will", "work", "years", "your"}

HISTORY_PATH = Path("temp_uploads") / "history.json"
HISTORY_PATH.parent.mkdir(exist_ok=True)
if not HISTORY_PATH.exists():
    HISTORY_PATH.write_text("[]", encoding="utf-8")


def inject_styles():
    st.markdown("""
    <style>
      .stApp { background: #f7f8fc; color: #172033; }
      [data-testid="stSidebar"] { background: #101a35; }
      [data-testid="stSidebar"] * { color: #f7f9ff !important; }
      [data-testid="stSidebar"] .stRadio label { border-radius: 10px; padding: 9px 10px; margin: 3px 0; font-weight: 700; opacity: 1 !important; }
      [data-testid="stSidebar"] .stRadio label:has(input:checked) { background: #3659b8; box-shadow: inset 3px 0 0 #78ebd8; }
      .block-container { max-width: 1280px; padding-top: 2rem; padding-bottom: 3rem; overflow: visible !important; }
      section.main > div, .stApp, .stApp > div, [data-testid="stAppViewContainer"], [data-testid="stVerticalBlock"], [data-testid="stHorizontalBlock"], [data-testid="stColumn"], [data-testid="stTabPanel"], [data-baseweb="tab-panel"], [data-baseweb="tab-list"], .stTabs, .stTabs > div { overflow: visible !important; }
      [data-testid="stVerticalBlock"], [data-testid="stHorizontalBlock"], [data-testid="stColumn"] { position: relative; z-index: 1; }
      h1, h2, h3 { letter-spacing: -0.035em; }
      div[data-testid="stMetric"] { background: white; border: 1px solid #e7eaf2; border-radius: 14px; padding: 14px 16px; box-shadow: 0 6px 22px rgba(25, 38, 73, .045); }
      .hero { padding: 38px 42px; border-radius: 22px; color: #fff; background: radial-gradient(circle at 88% 12%, #57d4c2 0, transparent 29%), linear-gradient(120deg,#172a5a,#3659b8); box-shadow: 0 16px 36px rgba(31, 57, 125, .22); }
      .hero h1 { font-size: 2.7rem; margin: 0 0 8px; color: white; }
      .hero p { font-size: 1.1rem; margin: 0; opacity: .88; }
      .eyebrow { color: #78ebd8; text-transform: uppercase; letter-spacing: .12em; font-weight: 700; font-size: .76rem; }
      .feature-card, .insight-card { background: #fff; border: 1px solid #e7eaf2; border-radius: 16px; padding: 20px; height: 100%; box-sizing: border-box; }
      .feature-card h3, .insight-card h3 { margin: 0 0 7px; font-size: 1.02rem; }
      .feature-card p, .insight-card p { color: #667085; margin: 0; line-height: 1.5; }
      .step { display:flex; gap:12px; align-items:center; padding: 14px 0; border-bottom: 1px solid #edf0f5; }
      .step:last-child { border-bottom:0; }
      .step-num { flex: 0 0 30px; height:30px; border-radius:50%; background:#e8edff; color:#3659b8; text-align:center; line-height:30px; font-weight:700; }
      .score-orb { width: 166px; height:166px; border-radius:50%; display:flex; flex-direction:column; justify-content:center; align-items:center; margin: 8px auto 14px; background: conic-gradient(var(--score-color) calc(var(--score) * 1%), #e9edf5 0); position:relative; }
      .score-orb:before { content:""; position:absolute; width:130px; height:130px; border-radius:50%; background:white; }
      .score-orb strong, .score-orb span { z-index:1; } .score-orb strong { font-size:2.25rem; } .score-orb span { color:#667085; font-size:.78rem; }
      .bar-label { display:flex; justify-content:space-between; margin:12px 0 6px; font-size:.88rem; font-weight:600; }
      .bar-track { background:#edf0f6; height:9px; border-radius:10px; overflow:hidden; } .bar-fill { height:100%; border-radius:10px; background:linear-gradient(90deg,#4667d9,#5ccbbd); animation: grow .7s ease-out; transform-origin:left; } @keyframes grow { from { transform:scaleX(0); } to { transform:scaleX(1); } }
      .tag { display:inline-block; background:#eef2ff; color:#3b59b8; padding:5px 9px; margin:2px; border-radius:99px; font-size:.82rem; font-weight:600; }
      .missing { background:#fff4e5; color:#a35b00; }
      .small-muted { color:#667085; font-size:.88rem; }
      .stButton > button { border-radius:10px; font-weight:650; min-height:2.7rem; }
      [data-testid="stFileUploader"] { background:#fff; padding:10px; border:1px dashed #9eace5; border-radius:14px; }
      textarea, input[type="text"], [data-testid="stTextArea"] textarea, [data-testid="stTextInput"] input, [data-baseweb="textarea"] textarea {
        color: #172033 !important; -webkit-text-fill-color: #172033 !important;
        caret-color: #172033 !important; background: #ffffff !important;
        border: 1px solid #cfd7e6 !important; border-radius: 12px !important;
        font-size: .96rem !important; line-height: 1.5 !important;
      }
      textarea::placeholder, input[type="text"]::placeholder, [data-testid="stTextArea"] textarea::placeholder, [data-testid="stTextInput"] input::placeholder {
        color: #7a8699 !important; -webkit-text-fill-color: #7a8699 !important; opacity: 1;
      }
      textarea:focus, input[type="text"]:focus, [data-testid="stTextArea"] textarea:focus, [data-testid="stTextInput"] input:focus {
        border-color: #4667d9 !important; box-shadow: 0 0 0 3px rgba(70, 103, 217, .14) !important;
      }
      [data-testid="stTextArea"] label, [data-testid="stFileUploader"] label { color: #25324b !important; font-weight: 700 !important; }
      [data-testid="stFileUploader"] button { color: #2449b8 !important; border-color: #aab8e8 !important; background: #f7f9ff !important; }
      [data-testid="stDownloadButton"] button { background: #eef2ff !important; border: 1px solid #9eace5 !important; color: #193f9d !important; }
      [data-testid="stDownloadButton"] button *, [data-testid="stDownloadButton"] button p { color: #193f9d !important; -webkit-text-fill-color: #193f9d !important; }
      [data-testid="stText"], [data-testid="stCaptionContainer"] { color: #53617a !important; }
      [data-testid="stText"] *, [data-testid="stCaptionContainer"] * { color: inherit !important; }
      [data-baseweb="tab-list"] { gap: 8px; border-bottom: 1px solid #e7eaf2; }
    [data-baseweb="tab-list"] [data-baseweb="tab"] { height: 42px; padding: 0 12px; color: #344054 !important; -webkit-text-fill-color: #344054 !important; font-weight: 700 !important; opacity: 1 !important; }
    [data-baseweb="tab-list"] [data-baseweb="tab"] p, [data-baseweb="tab-list"] [data-baseweb="tab"] span { color: #344054 !important; -webkit-text-fill-color: #344054 !important; opacity: 1 !important; }
    [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] { color: #c62828 !important; -webkit-text-fill-color: #c62828 !important; border-bottom: 3px solid #ef4444 !important; }
    [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] p, [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] span { color: #c62828 !important; -webkit-text-fill-color: #c62828 !important; }
    [data-baseweb="tab-highlight"] { background-color: #ef4444 !important; height: 3px !important; }
      .page-intro { background: #ffffff; border: 1px solid #e7eaf2; border-radius: 18px; padding: 22px 24px; margin: 0 0 22px; }
      .page-intro h1 { margin: 0 0 5px; color: #172a5a; font-size: 2rem; }
      .page-intro p { margin: 0; color: #667085; }
      .workflow-strip { display:grid; grid-template-columns:repeat(3, minmax(0, 1fr)); gap:10px; margin:0 0 20px; }
      .workflow-item { background:#fff; border:1px solid #dce3f0; border-radius:12px; padding:12px 14px; color:#53617a; font-size:.88rem; line-height:1.35; }
      .workflow-item strong { display:block; color:#172a5a; font-size:.92rem; margin-bottom:3px; }
      .workflow-item.active { border-color:#6e87df; background:#f3f6ff; box-shadow:0 3px 12px rgba(54,89,184,.09); }
      .workflow-number { color:#3659b8; font-weight:800; margin-right:6px; }
      .history-card { background: #ffffff; border: 1px solid #e7eaf2; border-radius: 16px; padding: 18px 20px; margin-bottom: 12px; }
      .history-card strong { color: #172a5a; }
      .history-meta { color: #667085; font-size: .88rem; margin-top: 4px; }
            .toolkit-heading { display:flex; align-items:center; gap:10px; margin: 26px 0 4px; color:#172a5a; font-size:1.35rem; font-weight:800; }
            .toolkit-heading::before { content:""; display:block; width:5px; height:25px; border-radius:4px; background:#2ba891; }
            .toolkit-subtitle { color:#3f4d66; font-size:.94rem; margin:0 0 14px; line-height:1.5; }
            .tool-index { display:grid; grid-template-columns:repeat(3, minmax(0, 1fr)); gap:10px; margin:0 0 18px; }
            .tool-index-card { background:#ffffff; border:1px solid #cbd5e5; border-left:4px solid #4667d9; border-radius:10px; padding:13px 14px; min-height:78px; box-sizing:border-box; box-shadow:0 3px 10px rgba(25,38,73,.045); }
            .tool-index-card:nth-child(3n+2) { border-left-color:#2ba891; }
            .tool-index-card:nth-child(3n) { border-left-color:#e7a329; }
            .tool-index-card strong { display:block; color:#172a5a; font-size:.92rem; margin-bottom:4px; }
            .tool-index-card span { display:block; color:#53617a; font-size:.8rem; line-height:1.35; }
            @media (max-width: 800px) { .tool-index { grid-template-columns:1fr; } }
            div[data-testid="stExpander"] { background:#ffffff !important; border:1px solid #cbd5e5 !important; border-radius:12px !important; margin:10px 0 !important; box-shadow:0 3px 12px rgba(25,38,73,.06); }
            div[data-testid="stExpander"] details summary { background:#f1f5ff !important; color:#172a5a !important; border-radius:11px !important; min-height:3rem; }
            div[data-testid="stExpander"] details summary svg { opacity:1 !important; color:#2449b8 !important; fill:#2449b8 !important; }
            div[data-testid="stExpander"] details summary:hover { background:#e5ecff !important; }
            div[data-testid="stExpander"] details summary p, div[data-testid="stExpander"] details summary span { color:#172a5a !important; -webkit-text-fill-color:#172a5a !important; font-weight:750 !important; }
            div[data-testid="stExpander"] details[open] summary { border-bottom:1px solid #d8e0ef; border-radius:11px 11px 0 0 !important; }
            div[data-testid="stExpander"] details > div { padding-top:12px; }
            .tool-status { display:inline-block; color:#176b5c; background:#e0f6f0; border:1px solid #a8dfd2; border-radius:999px; padding:3px 9px; font-size:.78rem; font-weight:750; }
            .tool-status.review { color:#8a5000; background:#fff1d6; border-color:#f2cf8e; }
            .visibility-note { background:#edf7f5; border:1px solid #b8e4da; border-radius:10px; color:#175d52; padding:10px 12px; margin:0 0 14px; font-size:.88rem; }
            @media (max-width: 800px) { .workflow-strip { grid-template-columns:1fr; } }
            /* Keep essential labels readable even when an external Streamlit theme overrides inherited text colour. */
            div[data-testid="stMetric"] { background:#ffffff !important; color:#172033 !important; }
            div[data-testid="stMetric"] [data-testid="stMetricLabel"],
            div[data-testid="stMetric"] [data-testid="stMetricValue"],
            div[data-testid="stMetric"] [data-testid="stMetricDelta"],
            div[data-testid="stMetric"] [data-testid="stMetricLabel"] *,
            div[data-testid="stMetric"] [data-testid="stMetricValue"] *,
            div[data-testid="stMetric"] [data-testid="stMetricDelta"] * { color:#172033 !important; -webkit-text-fill-color:#172033 !important; opacity:1 !important; text-shadow:none !important; }
            [data-testid="stTabs"] [role="tab"],
            [data-testid="stTabs"] button[data-baseweb="tab"],
            [data-baseweb="tab-list"] [data-baseweb="tab"] { background:#ffffff !important; color:#172033 !important; -webkit-text-fill-color:#172033 !important; opacity:1 !important; border:1px solid #d5ddeb !important; border-bottom:3px solid #d5ddeb !important; border-radius:9px 9px 0 0 !important; }
            [data-testid="stTabs"] [role="tab"] *,
            [data-testid="stTabs"] button[data-baseweb="tab"] *,
            [data-baseweb="tab-list"] [data-baseweb="tab"] * { color:#172033 !important; -webkit-text-fill-color:#172033 !important; opacity:1 !important; }
            [data-testid="stTabs"] [role="tab"][aria-selected="true"],
            [data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"],
            [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] { background:#eef2ff !important; border-color:#4667d9 !important; border-bottom-color:#4667d9 !important; }
            [data-testid="stTabs"] [role="tab"][aria-selected="true"] *,
            [data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] *,
            [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] * { color:#193f9d !important; -webkit-text-fill-color:#193f9d !important; }
            /* Product-level visual system */
            .stApp { background:linear-gradient(180deg, #f8faff 0%, #f5f7fc 55%, #f8fafc 100%) !important; }
            .block-container { max-width:1240px !important; padding-top:2.35rem !important; }
            [data-testid="stSidebar"] { background:linear-gradient(180deg, #0d1833 0%, #142852 100%) !important; border-right:1px solid rgba(255,255,255,.08); }
            .brand-mark { display:flex; gap:11px; align-items:center; padding:7px 2px 20px; }
            .brand-icon { width:36px; height:36px; border-radius:11px; display:grid; place-items:center; background:linear-gradient(135deg,#63d5c2,#5274e8); color:#07152f !important; font-weight:900; font-size:1.1rem; box-shadow:0 8px 18px rgba(70,103,217,.25); }
            .brand-name { color:#ffffff; font-size:1.08rem; font-weight:800; letter-spacing:-.02em; line-height:1.1; }
            .brand-subtitle { color:#aebddd !important; font-size:.72rem; margin-top:3px; }
            [data-testid="stSidebar"] [data-testid="stRadio"] { padding:8px 0; }
            [data-testid="stSidebar"] .stRadio label { min-height:40px; display:flex; align-items:center; transition:background .18s ease, transform .18s ease; }
            [data-testid="stSidebar"] .stRadio label:hover { background:rgba(100,130,220,.24); transform:translateX(2px); }
            h1, h2, h3 { color:#14213d; }
            .page-intro, .analysis-hero { position:relative; overflow:hidden; background:linear-gradient(115deg,#172a5a 0%,#284b9e 62%,#3d72c7 100%); border:0; box-shadow:0 18px 42px rgba(23,42,90,.18); padding:30px 34px; }
            .page-intro::after, .analysis-hero::after { content:""; position:absolute; width:240px; height:240px; border-radius:50%; background:rgba(120,235,216,.15); right:-68px; top:-110px; }
            .page-intro h1, .analysis-hero h1 { position:relative; z-index:1; color:#fff; font-size:2.2rem; }
            .page-intro p, .analysis-hero p { position:relative; z-index:1; color:#dbe6ff; font-size:1rem; max-width:720px; }
            .page-kicker { position:relative; z-index:1; margin:0 0 8px; color:#85e6d5; font-size:.74rem; font-weight:800; letter-spacing:.12em; text-transform:uppercase; }
            .workflow-strip { gap:12px; margin:22px 0 24px; }
            .workflow-item { border:1px solid #e0e7f3; border-radius:14px; padding:14px 16px; box-shadow:0 5px 16px rgba(19,40,86,.04); }
            .workflow-item.active { border-color:#93a9ec; background:linear-gradient(135deg,#f5f7ff,#eef5ff); box-shadow:0 8px 20px rgba(54,89,184,.10); }
            [data-testid="stFileUploader"], [data-testid="stTextArea"] { border-radius:16px; }
            [data-testid="stFileUploader"] { padding:16px; border:1.5px dashed #7d96df; box-shadow:0 4px 15px rgba(24,49,111,.04); }
            [data-testid="stFileUploader"] section { padding:12px 6px; }
            [data-testid="stTextArea"] textarea { min-height:164px; box-shadow:0 3px 10px rgba(19,40,86,.025); }
            .stButton > button, [data-testid="stDownloadButton"] button { transition:transform .16s ease, box-shadow .16s ease; }
            .stButton > button[kind="primary"] { border:0 !important; background:linear-gradient(135deg,#2449b8,#3e6ed4) !important; box-shadow:0 9px 18px rgba(36,73,184,.22); color:#fff !important; }
            .stButton > button[kind="primary"] * { color:#fff !important; }
            .stButton > button:hover, [data-testid="stDownloadButton"] button:hover { transform:translateY(-1px); box-shadow:0 8px 17px rgba(27,54,122,.16); }
            .stButton > button:focus-visible, [data-testid="stDownloadButton"] button:focus-visible, [role="tab"]:focus-visible { outline:3px solid rgba(70,103,217,.38) !important; outline-offset:2px; }
            .report-header { display:flex; align-items:center; justify-content:space-between; gap:16px; margin:30px 0 4px; }
            .report-header h2 { margin:0; font-size:1.55rem; }
            .report-badge { display:inline-flex; align-items:center; gap:7px; border:1px solid #bce4da; background:#eaf8f4; color:#176b5c; border-radius:999px; padding:6px 10px; font-size:.78rem; font-weight:750; white-space:nowrap; }
            div[data-testid="stMetric"] { border:1px solid #e0e7f3 !important; box-shadow:0 7px 20px rgba(19,40,86,.055); }
            .score-orb { box-shadow:0 10px 25px rgba(20,47,103,.07); }
            .history-card { border:1px solid #e0e7f3; box-shadow:0 5px 15px rgba(19,40,86,.045); transition:transform .16s ease, box-shadow .16s ease; }
            .history-card:hover { transform:translateY(-2px); box-shadow:0 11px 24px rgba(19,40,86,.09); }
            .feature-card, .insight-card { border-color:#e0e7f3; box-shadow:0 7px 18px rgba(19,40,86,.04); }
            .feature-card { transition:transform .16s ease, box-shadow .16s ease; }
            .feature-card:hover { transform:translateY(-3px); box-shadow:0 13px 26px rgba(19,40,86,.09); }
            .section-lead { margin:18px 0 18px; padding:18px 20px; border:1px solid #dbe4f3; border-radius:15px; background:linear-gradient(110deg,#f8faff,#f1f6ff); }
            .section-eyebrow { color:#4064c5; font-size:.72rem; letter-spacing:.11em; font-weight:800; text-transform:uppercase; margin-bottom:5px; }
            .section-lead h3 { color:#172a5a; margin:0 0 5px; font-size:1.15rem; }
            .section-lead p { color:#53617a; margin:0; font-size:.92rem; line-height:1.5; }
            .alignment-panel { background:#fff; border:1px solid #e0e7f3; border-radius:15px; padding:17px 18px; min-height:145px; box-shadow:0 5px 14px rgba(19,40,86,.035); }
            .alignment-panel.covered { border-top:4px solid #2ba891; }
            .alignment-panel.gaps { border-top:4px solid #e7a329; }
            .alignment-panel h4 { color:#172a5a; margin:0 0 5px; font-size:1rem; }
            .alignment-panel p { color:#667085; margin:0 0 10px; font-size:.84rem; }
            .signal-card { background:#fff; border:1px solid #e0e7f3; border-radius:15px; padding:17px; box-shadow:0 5px 14px rgba(19,40,86,.035); }
            .strategy-card { border:1px solid #c9d7f7; background:linear-gradient(135deg,#eef4ff,#f7f9ff); border-radius:15px; padding:20px; }
            .strategy-card p { color:#2b3c60; line-height:1.62; margin:0; }
            .danger-zone { background:#fff7f7; border:1px solid #f1c4c9; border-left:4px solid #d95d6a; border-radius:13px; padding:14px 16px; }
            .danger-zone strong { color:#922f3d; display:block; margin-bottom:3px; }
            .danger-zone span { color:#7e5960; font-size:.86rem; line-height:1.4; }
            .st-key-clear_history_button button { background:#c73645 !important; border:1px solid #a92332 !important; color:#ffffff !important; box-shadow:0 7px 14px rgba(174,35,50,.18) !important; }
            .st-key-clear_history_button button * { color:#ffffff !important; -webkit-text-fill-color:#ffffff !important; }
            .st-key-clear_history_button button:hover:not(:disabled) { background:#a92332 !important; box-shadow:0 9px 18px rgba(174,35,50,.26) !important; }
            .st-key-clear_history_button button:disabled { background:#e8b8be !important; border-color:#e0a2aa !important; color:#7b3d45 !important; box-shadow:none !important; opacity:1 !important; }
            .st-key-clear_history_button button:disabled * { color:#7b3d45 !important; -webkit-text-fill-color:#7b3d45 !important; }
            /* Report navigation: quiet underline tabs, not button-like boxes. */
            div[data-testid="stTabs"] { margin-top:34px !important; }
            div[data-testid="stTabs"] [data-baseweb="tab-list"] { gap:24px !important; padding:0 3px !important; border-bottom:1px solid #dce3ef !important; }
            div[data-testid="stTabs"] [role="tab"],
            div[data-testid="stTabs"] button[data-baseweb="tab"],
            div[data-testid="stTabs"] [data-baseweb="tab-list"] [data-baseweb="tab"] { height:44px !important; padding:0 0 10px !important; background:transparent !important; border:0 !important; border-bottom:2px solid transparent !important; border-radius:0 !important; color:#65718a !important; font-size:.91rem !important; font-weight:650 !important; box-shadow:none !important; }
            div[data-testid="stTabs"] [role="tab"] *,
            div[data-testid="stTabs"] button[data-baseweb="tab"] *,
            div[data-testid="stTabs"] [data-baseweb="tab-list"] [data-baseweb="tab"] * { color:#65718a !important; -webkit-text-fill-color:#65718a !important; }
            div[data-testid="stTabs"] [role="tab"][aria-selected="true"],
            div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"],
            div[data-testid="stTabs"] [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] { background:transparent !important; border:0 !important; border-bottom:2px solid #3659b8 !important; }
            div[data-testid="stTabs"] [role="tab"][aria-selected="true"] *,
            div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] *,
            div[data-testid="stTabs"] [data-baseweb="tab-list"] [data-baseweb="tab"][aria-selected="true"] * { color:#193f9d !important; -webkit-text-fill-color:#193f9d !important; font-weight:750 !important; }
            @media (max-width: 720px) { .page-intro, .analysis-hero { padding:24px 21px; } .page-intro h1, .analysis-hero h1 { font-size:1.75rem; } .report-header { align-items:flex-start; flex-direction:column; } }
    </style>
    """, unsafe_allow_html=True)


inject_styles()


def load_history():
    try:
        data = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def save_history(entry):
    history = load_history()
    history.append(entry)
    HISTORY_PATH.write_text(json.dumps(history[-30:], indent=2), encoding="utf-8")


@st.cache_data(show_spinner=False)
def extract_text(file_bytes, filename):
    extension = filename.rsplit(".", 1)[-1].lower()
    buffer = io.BytesIO(file_bytes)
    if extension == "pdf":
        return "\n".join(page.extract_text() or "" for page in PyPDF2.PdfReader(buffer).pages)
    if extension == "docx":
        return "\n".join(paragraph.text for paragraph in Document(buffer).paragraphs)
    return ""


def render_resume_preview(file_bytes, filename, resume_text):
    """Keep the original live document preview, presented as a collapsible review step."""
    extension = filename.rsplit(".", 1)[-1].lower()
    with st.expander("Live resume preview", expanded=False):
        if extension == "pdf":
            encoded = base64.b64encode(file_bytes).decode("utf-8")
            st.markdown(f"<iframe title='Resume preview' src='data:application/pdf;base64,{encoded}' width='100%' height='720' style='border:1px solid #e7eaf2;border-radius:12px;'></iframe>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='insight-card'><h3>DOCX text preview</h3><p style='white-space:pre-wrap;max-height:650px;overflow:auto'>"
                        + html.escape(resume_text[:12000]) + "</p></div>", unsafe_allow_html=True)


def jd_terms(job_description):
    words = re.findall(r"\b[a-z][a-z+#.]{2,}\b", job_description.lower())
    common = [word for word, _ in Counter(words).most_common(20) if word not in STOP_WORDS]
    known = [skill for skill in SKILLS if re.search(rf"(?<!\w){re.escape(skill)}(?!\w)", job_description.lower())]
    return list(dict.fromkeys(known + common))[:15]


def term_variants(term):
    """Return close, role-relevant wording variants without treating unrelated words as matches."""
    normalized = term.lower().strip()
    variants = {normalized}
    role_aliases = {
        "tester": {"test", "testing", "qa", "quality assurance", "quality analyst", "test engineer", "software tester"},
        "testing": {"test", "tester", "qa", "quality assurance", "quality analyst", "test engineer", "software tester"},
        "qa": {"quality assurance", "quality analyst", "tester", "testing", "test engineer"},
        "quality assurance": {"qa", "quality analyst", "tester", "testing", "test engineer"},
    }
    variants.update(role_aliases.get(normalized, set()))
    if normalized.endswith("er") and len(normalized) > 4:
        root = normalized[:-2]
        variants.update({root, root + "ing"})
    elif normalized.endswith("ing") and len(normalized) > 5:
        root = normalized[:-3]
        variants.update({root, root + "er"})
    return variants


def has_term(text, term):
    lowered = text.lower()
    return any(bool(re.search(rf"(?<!\w){re.escape(variant)}(?!\w)", lowered)) for variant in term_variants(term))


def analyze(text, job_description=""):
    lowered = text.lower()
    contact = {
        "Email": bool(re.search(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b", text)),
        "Phone": bool(re.search(r"(?:\+?\d[\d .()-]{7,}\d)", text)),
        "LinkedIn": "linkedin.com/in" in lowered,
        "Portfolio": "github.com" in lowered or "portfolio" in lowered,
    }
    headings = {
        "Experience": ["experience", "employment", "work history", "internship"],
        "Education": ["education", "bachelor", "master", "university", "college"],
        "Skills": ["skills", "technical skills", "competencies"],
        "Projects": ["projects", "project experience", "selected projects"],
    }
    sections = {name: 100 if any(term in lowered for term in terms) else 25 for name, terms in headings.items()}
    found_skills = [skill for skill in SKILLS if has_term(lowered, skill)]
    sections["Skills"] = min(100, max(sections["Skills"], round(len(found_skills) / 7 * 100)))
    action_count = sum(len(re.findall(rf"(?<!\w){verb}(?:ed|ing|s)?(?!\w)", lowered)) for verb in ACTION_VERBS)
    metrics_count = len(re.findall(r"(?:\$|₹|€|£)?\s*\d+(?:\.\d+)?\s*(?:%|x|k\b|m\b|million|thousand|users|customers|hours|days)", lowered))
    readability = max(0, min(100, round(flesch_reading_ease(text))))
    structure = round(sum(sections.values()) / len(sections))
    impact = min(100, 45 + action_count * 7 + metrics_count * 8)
    contact_score = round(sum(contact.values()) / len(contact) * 100)
    core_score = round(structure * .48 + impact * .32 + contact_score * .20)
    terms = jd_terms(job_description) if job_description.strip() else []
    matched_terms = [term for term in terms if has_term(lowered, term)]
    missing_terms = [term for term in terms if term not in matched_terms]
    match_score = round(len(matched_terms) / len(terms) * 100) if terms else 0
    score = min(100, round(core_score * .72 + match_score * .28) if terms else core_score)
    suggestions = []
    if not contact["Email"] or not contact["Phone"]:
        suggestions.append(("Contact details", "Add a professional email address and phone number in the header."))
    if not contact["LinkedIn"]:
        suggestions.append(("Professional presence", "Add a LinkedIn URL so recruiters can verify your professional profile."))
    if metrics_count < 3:
        suggestions.append(("Show impact", "Add 2–3 measurable outcomes, such as time saved, revenue influenced, scale, or percentage change."))
    if action_count < 5:
        suggestions.append(("Strengthen bullets", "Lead achievement bullets with specific action verbs: built, led, improved, delivered, or optimized."))
    if missing_terms:
        suggestions.append(("Tailor to this role", f"Where truthful, reflect these priority terms in your summary and experience: {', '.join(missing_terms[:5])}."))
    if sections["Projects"] < 100:
        suggestions.append(("Evidence of work", "Add a Projects section with scope, tools used, and one measurable result."))
    exp_years = sum(int(year) for year in re.findall(r"(\d+)\s+years?", lowered))
    projects_count = len(re.findall(r"\bprojects?\b", lowered))
    return {"score": score, "core_score": core_score, "match_score": match_score, "sections": sections, "contact": contact, "skills": found_skills, "action_count": action_count, "metrics_count": metrics_count, "readability": readability, "terms": terms, "matched_terms": matched_terms, "missing_terms": missing_terms, "suggestions": suggestions[:5], "exp_years": exp_years, "projects_count": projects_count}


def resume_bullets(text):
    """Extract likely achievement lines for rewrite prompts without claiming they are facts."""
    candidates = []
    for line in text.splitlines():
        clean = re.sub(r"^[\s•*\-–—]+", "", line).strip()
        if len(clean) >= 35 and (clean.startswith(tuple(verb.title() for verb in ACTION_VERBS)) or re.match(r"^(built|led|created|managed|developed|implemented|improved|analyzed)\b", clean, re.I)):
            candidates.append(clean)
    if not candidates:
        candidates = [item.strip() for item in re.split(r"(?<=[.!?])\s+", text) if len(item.strip()) >= 45]
    return candidates[:4]


def gap_coach(text, result):
    source_lines = resume_bullets(text)
    rewrites = []
    for index, term in enumerate(result["missing_terms"][:5]):
        current = source_lines[index % len(source_lines)] if source_lines else "Choose a relevant project or experience bullet."
        verb = ["Delivered", "Improved", "Built", "Partnered", "Optimized"][index % 5]
        rewrite = f"{verb} [project or process] using {term.title()} to [specific outcome], improving [metric] by [amount]."
        rewrites.append((term, current, rewrite))
    return rewrites


def ats_scan(text, job_description, result):
    words = re.findall(r"\b\w+\b", text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    keyword_hits = sum(len(re.findall(rf"(?<!\w){re.escape(term)}(?!\w)", text.lower())) for term in result["terms"])
    density = round(keyword_hits / max(len(words), 1) * 100, 2)
    issues = []
    if len(words) < 180:
        issues.append(("Content length", "Your extracted resume is short. Add concise evidence for experience, projects, and outcomes."))
    if not result["contact"]["Email"] or not result["contact"]["Phone"]:
        issues.append(("Header details", "Add a readable email address and phone number near the top of the resume."))
    missing_sections = [name for name, value in result["sections"].items() if value < 100]
    if missing_sections:
        issues.append(("Section labels", "Use clear ATS-friendly headings: " + ", ".join(missing_sections) + "."))
    if sum(1 for line in lines if len(line) > 180) > 3:
        issues.append(("Dense text", "Split long paragraphs into short achievement bullets so both parsers and recruiters can scan them."))
    if job_description and density > 4:
        issues.append(("Keyword balance", "Role-term density is high. Keep only terms supported by a concrete accomplishment to avoid keyword stuffing."))
    if not issues:
        issues.append(("Parser check", "No obvious text-extraction risks were detected. Verify the final PDF still has selectable text before applying."))
    order = [name for name in ("Experience", "Education", "Skills", "Projects") if result["sections"][name] >= 100]
    return {"issues": issues, "density": density, "keyword_hits": keyword_hits, "word_count": len(words), "section_order": order}


def interview_questions(result):
    questions = ["Walk me through the achievement on your resume that best demonstrates impact for this role."]
    for term in result["matched_terms"][:3]:
        questions.append(f"Tell me about a time you used {term.title()}. What was the problem, what did you do, and what changed?")
    for term in result["missing_terms"][:2]:
        questions.append(f"This role emphasizes {term.title()}. Which transferable experience would you use to show you can succeed with it?")
    questions.append("Which result from your resume would you quantify differently after reviewing this job description?")
    return questions[:6]


def tailored_summary(result):
    strengths = ", ".join(skill.title() for skill in result["skills"][:4]) or "relevant technical and business skills"
    target = ", ".join(term.title() for term in result["matched_terms"][:3]) or "the target role's core requirements"
    gap = result["missing_terms"][0].title() if result["missing_terms"] else None
    summary = f"Results-focused professional with experience in {strengths}, bringing evidence of delivery through projects and measurable outcomes."
    summary += f" Tailored for roles that value {target}, with a clear focus on translating relevant experience into business impact."
    if gap:
        summary += f" Where accurate, add one concrete example that demonstrates {gap}."
    return summary


def skill_roadmap(job_description, result):
    requested_skills = [skill for skill in SKILLS if has_term(job_description, skill)]
    missing_skills = [skill for skill in requested_skills if skill not in result["skills"]]
    priorities = []
    for index, skill in enumerate(missing_skills[:5]):
        priority = "High priority" if index < 2 else "Build next"
        evidence = f"Complete a small {skill.title()} project and add the tool, scope, and result to your resume."
        priorities.append((skill, priority, evidence))
    if not priorities and requested_skills:
        priorities.append(("Evidence depth", "Strengthen", "Your requested technical skills are covered. Add a project or quantified outcome for the two most relevant ones."))
    if not priorities and job_description.strip():
        testing_role = any(has_term(job_description, term) for term in ("tester", "testing", "qa", "quality assurance"))
        if testing_role:
            priorities.append(("QA and testing evidence", "Strengthen", "Add one concise bullet showing your test approach, such as test cases, regression coverage, defect reporting, automation, or a quality outcome."))
            priorities.append(("Role-ready proof", "Build next", "Prepare a small testing portfolio example: the scenario, checks performed, tools used, defect found or prevented, and the measurable result."))
        else:
            focus_terms = jd_terms(job_description)[:3]
            focus = ", ".join(term.title() for term in focus_terms) or "the target role"
            priorities.append(("Role evidence", "Strengthen", f"The target role emphasizes {focus}. Add a truthful experience or project bullet that demonstrates the most relevant responsibility and its outcome."))
    return priorities


def evidence_trace(text, job_description, result):
    """Return explainable source lines for the signals used by the local scorer."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    traces = []

    def add_trace(signal, status, matches, explanation):
        if matches:
            traces.append((signal, status, matches[:3], explanation))
        else:
            traces.append((signal, status, [], explanation))

    contact_matches = [line for line in lines if re.search(r"@|linkedin\.com/in|github\.com|(?:\+?\d[\d .()\-]{7,}\d)", line, re.I)]
    add_trace("Contact details", "Supported" if all(result["contact"].values()) else "Needs review", contact_matches, "Email, phone, LinkedIn, and portfolio signals contribute to contact completeness.")

    section_matches = [line for line in lines if any(term in line.lower() for terms in ("experience", "education", "skills", "projects") for term in terms)]
    add_trace("Resume sections", "Supported" if all(value == 100 for value in result["sections"].values()) else "Needs review", section_matches, "Recognized section headings contribute to structural completeness.")

    action_matches = [line for line in lines if any(re.search(rf"(?<!\w){re.escape(verb)}(?:ed|ing|s)?(?!\w)", line.lower()) for verb in ACTION_VERBS)]
    add_trace("Action-oriented language", "Supported" if result["action_count"] else "Needs review", action_matches, "Action verbs contribute to the impact signal.")

    metric_matches = [line for line in lines if re.search(r"(?:\$|₹|€|£)?\s*\d+(?:\.\d+)?\s*(?:%|x|k\b|m\b|million|thousand|users|customers|hours|days)", line.lower())]
    add_trace("Measurable outcomes", "Supported" if result["metrics_count"] else "Needs review", metric_matches, "Numbers, percentages, scale, and time measures contribute to evidence of impact.")

    role_matches = [line for line in lines if any(has_term(line, term) for term in result["matched_terms"])]
    add_trace("Role alignment", "Supported" if role_matches else ("Unavailable" if not job_description.strip() else "Needs review"), role_matches, "Terms shared with the target job description contribute to role alignment.")
    return traces


def claim_review(text):
    """Flag claims that deserve human verification; this is not a truth detector."""
    patterns = [
        (r"\b(certified|certification|licensed|licence|license)\b", "Certification or licence claim"),
        (r"\b(expert|mastery|proficient|specialist)\b", "Strength-of-skill claim"),
        (r"\b\d+\+?\s+years?\b", "Experience-duration claim"),
        (r"\b(?:increased|reduced|improved|saved|generated|grew)\b[^\n.]{0,100}\b\d+(?:\.\d+)?\s*%", "Quantified achievement"),
    ]
    findings = []
    for line in [item.strip() for item in text.splitlines() if item.strip()]:
        for pattern, label in patterns:
            if re.search(pattern, line, re.I):
                findings.append((label, line))
                break
    return findings[:12]


def simulated_score(result, improvements):
    """Estimate the score impact without mutating the original analysis result."""
    sections = dict(result["sections"])
    contact = dict(result["contact"])
    action_count = result["action_count"]
    metrics_count = result["metrics_count"]
    matched_terms = len(result["matched_terms"])
    if improvements.get("sections"):
        sections.update({name: 100 for name, value in sections.items() if value < 100})
    if improvements.get("contact"):
        contact.update({name: True for name, value in contact.items() if not value})
    if improvements.get("impact"):
        action_count = max(action_count, 5)
        metrics_count = max(metrics_count, 3)
    if improvements.get("role"):
        matched_terms = len(result["terms"])
    structure = round(sum(sections.values()) / len(sections))
    impact = min(100, 45 + action_count * 7 + metrics_count * 8)
    contact_score = round(sum(contact.values()) / len(contact) * 100)
    core_score = round(structure * .48 + impact * .32 + contact_score * .20)
    match_score = round(matched_terms / len(result["terms"]) * 100) if result["terms"] else 0
    score = min(100, round(core_score * .72 + match_score * .28) if result["terms"] else core_score)
    return score, core_score, match_score


def version_comparison(original_text, revised_text, job_description):
    original = analyze(original_text, job_description)
    revised = analyze(revised_text, job_description)
    original_words = set(re.findall(r"\b[a-z][a-z+#.]{2,}\b", original_text.lower()))
    revised_words = set(re.findall(r"\b[a-z][a-z+#.]{2,}\b", revised_text.lower()))
    original_sections = {name for name, value in original["sections"].items() if value == 100}
    revised_sections = {name for name, value in revised["sections"].items() if value == 100}
    added_terms = sorted(revised["matched_terms"] if revised["terms"] else revised_words - original_words)
    removed_terms = sorted(original["matched_terms"] if original["terms"] else original_words - revised_words)
    changed_lines = list(difflib.unified_diff(original_text.splitlines(), revised_text.splitlines(), lineterm=""))
    return {
        "original": original,
        "revised": revised,
        "added_sections": sorted(revised_sections - original_sections),
        "removed_sections": sorted(original_sections - revised_sections),
        "added_terms": added_terms[:12],
        "removed_terms": removed_terms[:12],
        "changed_lines": changed_lines[:24],
    }


def score_color(score):
    return "#2ba891" if score >= 75 else "#e7a329" if score >= 55 else "#d95d6a"


def score_orb(score, label):
    st.markdown(f"<div class='score-orb' style='--score:{score};--score-color:{score_color(score)}'><strong>{score}</strong><span>{label}</span></div>", unsafe_allow_html=True)


def progress(label, value):
    st.markdown(f"<div class='bar-label'><span>{html.escape(label)}</span><span>{value}%</span></div><div class='bar-track'><div class='bar-fill' style='width:{value}%'></div></div>", unsafe_allow_html=True)


def tag_markup(items, missing=False):
    if not items:
        return "<span class='small-muted'>Nothing to show yet.</span>"
    kind = " missing" if missing else ""
    return "".join(f"<span class='tag{kind}'>{html.escape(item.title())}</span>" for item in items)


def tags(items, missing=False):
    st.markdown(tag_markup(items, missing), unsafe_allow_html=True)


def report_markdown(result, file_name, resume_text="", job_description=""):
    lines = [f"# Resume Insight Report — {file_name}", f"Generated: {dt.datetime.now():%d %b %Y, %H:%M}", "", f"## Overall score: {result['score']}/100", f"Core resume quality: {result['core_score']}/100", f"Role match: {result['match_score']}/100", "", "## Recommended next moves"]
    lines.extend(f"- **{title}:** {detail}" for title, detail in result["suggestions"])
    if result["missing_terms"]:
        lines += ["", "## Missing role terms", ", ".join(result["missing_terms"])]
    if result["terms"]:
        ats = ats_scan(resume_text, job_description, result)
        lines += ["", "## Tailored professional summary", tailored_summary(result), "", "## ATS scan snapshot", f"- Extracted words: {ats['word_count']}", f"- Role-term mentions: {ats['keyword_hits']}", f"- Keyword density: {ats['density']}%"]
        lines.extend(f"- **{title}:** {detail}" for title, detail in ats["issues"])
        lines += ["", "## Interview practice"]
        lines.extend(f"{index}. {question}" for index, question in enumerate(interview_questions(result), 1))
        roadmap = skill_roadmap(job_description, result)
        if roadmap:
            lines += ["", "## Skill roadmap"]
            lines.extend(f"- **{priority} - {skill.title()}:** {action}" for skill, priority, action in roadmap)
    return "\n".join(lines)


def mock_ai_feedback(result):
    """Retain the original AI strategy layer, with concrete language from the newer scoring model."""
    if not result["terms"]:
        return "Add a target job description to unlock a role-specific strategy. Your current priorities are to strengthen evidence of impact and make the most relevant achievements easier to find."
    match = result["match_score"]
    confidence = "High" if match >= 80 else "Medium" if match >= 60 else "Developing"
    missing = ", ".join(result["missing_terms"][:4])
    return f"**Match confidence: {confidence}.** Your role alignment is **{match}%**. " + (
        "Your resume already reflects much of the language recruiters are likely to scan for. Move your strongest matched skills and outcomes nearer the top."
        if match >= 70 else
        f"Focus your next edit on translating real experience into the role's vocabulary — especially: **{missing or 'the missing priority terms'}**. Add those terms only beside evidence that proves them."
    )


def build_audio_summary(result, rate=165, volume=1.0, voice_id=None):
    """Restore configurable voice playback with pyttsx3 and gTTS fallback."""
    message = (
        f"Your overall resume readiness score is {result['score']} out of 100. "
        f"Your core resume quality is {result['core_score']}. "
        + (f"Your match to the target role is {result['match_score']} percent. " if result["terms"] else "")
        + "Your next priority is: "
        + (result["suggestions"][0][1] if result["suggestions"] else "tailoring your strongest achievements for the role.")
    )
    if TTS_AVAILABLE:
        descriptor, path = tempfile.mkstemp(suffix=".wav", prefix="resume_insight_")
        os.close(descriptor)
        try:
            engine = pyttsx3.init()
            engine.setProperty("rate", rate)
            engine.setProperty("volume", volume)
            if voice_id:
                engine.setProperty("voice", voice_id)
            engine.save_to_file(message, path)
            engine.runAndWait()
            with open(path, "rb") as audio_file:
                return audio_file.read(), "audio/wav"
        except Exception:
            pass
        finally:
            if os.path.exists(path):
                os.remove(path)
    if GTTS_AVAILABLE:
        try:
            buffer = io.BytesIO()
            gTTS(message).write_to_fp(buffer)
            return buffer.getvalue(), "audio/mp3"
        except Exception:
            pass
    return None, None


with st.sidebar:
    st.markdown("<div class='brand-mark'><div class='brand-icon'>R</div><div><div class='brand-name'>Resume Insight</div><div class='brand-subtitle'>Application intelligence</div></div></div>", unsafe_allow_html=True)
    page = st.radio("Navigation", ["Home", "Analyze", "History"], label_visibility="collapsed")
    st.divider()
    st.markdown("**Your workflow**")
    st.caption("01  Upload a resume\n\n02  Add a target role\n\n03  Make evidence-led edits")


if page == "Home":
    st.markdown("""<section class='hero'><div class='eyebrow'>Resume intelligence, made practical</div><h1>Build a resume recruiters can scan with confidence.</h1><p>Analyze structure, evidence of impact, and alignment to a specific role — then leave with a focused edit plan.</p></section>""", unsafe_allow_html=True)
    st.write("")
    left, right = st.columns([1.25, 1])
    with left:
        st.subheader("A sharper application workflow")
        for number, title, copy in [("1", "Upload once", "Bring in a text-based PDF or DOCX and check the extracted preview."), ("2", "Tailor intentionally", "Paste the role description to surface the language that matters for this application."), ("3", "Edit with purpose", "Use a prioritized checklist, skill coverage, and downloadable report to make the next revision count.")]:
            st.markdown(f"<div class='step'><div class='step-num'>{number}</div><div><b>{title}</b><br><span class='small-muted'>{copy}</span></div></div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='insight-card'><h3>What’s new in the experience</h3><p>Role-keyword coverage, a concise action plan, an ATS-style readiness checklist, comparison-friendly history, and a portable report for every analysis.</p></div>", unsafe_allow_html=True)
        st.write("")
        st.info("Tip: paste the exact job description. The match view is most useful when it reflects the role you plan to apply for.")
    st.write("")
    st.subheader("Designed for the decisions that matter")
    cards = st.columns(3)
    content = [("Role alignment", "See which relevant terms already appear in your resume and which deserve a truthful, evidence-backed mention."), ("Impact signal", "Track action verbs, quantified achievements, structure, and contact completeness in one glance."), ("Progress over time", "Keep a lightweight private record of analyses so revisions become visible, not vague.")]
    for col, (title, copy) in zip(cards, content):
        with col:
            st.markdown(f"<div class='feature-card'><h3>{title}</h3><p>{copy}</p></div>", unsafe_allow_html=True)

elif page == "Analyze":
    st.markdown("<section class='analysis-hero'><div class='page-kicker'>Resume review workspace</div><h1>Make every application more convincing.</h1><p>Upload your resume and a target role to surface the strongest edits, evidence gaps, and ATS signals before you apply.</p></section>", unsafe_allow_html=True)
    st.markdown("""
        <div class='workflow-strip'>
          <div class='workflow-item active'><strong><span class='workflow-number'>1</span>Add your resume</strong>Upload a text-based PDF or DOCX.</div>
          <div class='workflow-item active'><strong><span class='workflow-number'>2</span>Add the role</strong>Paste the job post for a tailored match score.</div>
          <div class='workflow-item'><strong><span class='workflow-number'>3</span>Review your plan</strong>Create a report, then use the clearly labelled tools below.</div>
        </div>
    """, unsafe_allow_html=True)
    upload_col, role_col = st.columns([1, 1.15], gap="large")
    with upload_col:
        st.markdown("**Required: your resume**")
        uploaded = st.file_uploader("Resume file", type=["pdf", "docx"], help="Use a text-based PDF or DOCX. Scanned PDFs may not contain extractable text.")
        st.caption("PDF or DOCX · text-based files work best")
    with role_col:
        st.markdown("**Recommended: target job description**")
        job_description = st.text_area("Target job description (recommended)", height=166, placeholder="Paste responsibilities, required skills, and qualifications here…", help="This unlocks role match and keyword coverage.")

    st.caption("Tip: paste the exact job post to make the match score more useful.")

    resume_text, file_name = "", "Resume"
    if uploaded:
        file_name, file_bytes = uploaded.name, uploaded.getvalue()
        try:
            with st.spinner("Reading your resume…"):
                resume_text = extract_text(file_bytes, uploaded.name)
        except Exception:
            st.error("I couldn't read this file. Please try a different PDF/DOCX or export it again.")
        if resume_text.strip():
            meta1, meta2, meta3 = st.columns(3)
            meta1.metric("File", uploaded.name)
            meta2.metric("Size", f"{len(file_bytes) / 1024:.0f} KB")
            meta3.metric("Extracted words", len(re.findall(r"\b\w+\b", resume_text)))
            render_resume_preview(file_bytes, uploaded.name, resume_text)
            with st.expander("Review extracted resume text", expanded=False):
                st.text(resume_text[:8000])
        else:
            st.warning("No readable text was found. If this is a scan, use a text-based export before analyzing.")

    st.divider()
    action_col, note_col = st.columns([1, 1.8], vertical_alignment="center")
    with action_col:
        analyze_clicked = st.button("Create my insight report", type="primary", use_container_width=True, disabled=not bool(resume_text.strip()))
    with note_col:
        st.markdown("<div class='visibility-note'><b>Private by design.</b> Your document is processed in this local app. Add only keywords you can support with real experience.</div>", unsafe_allow_html=True)
    if analyze_clicked:
        result = analyze(resume_text, job_description)
        st.session_state["latest_analysis"] = {
            "result": result,
            "file_name": file_name,
            "has_jd": bool(job_description.strip()),
            "resume_text": resume_text,
            "job_description": job_description,
        }
        save_history({"timestamp": dt.datetime.now().isoformat(timespec="seconds"), "file_name": file_name, "overall_score": result["score"], "match_score": result["match_score"], "readability": result["readability"], "top_skills": result["skills"][:6], "feedback_summary": result["suggestions"][0][1] if result["suggestions"] else "Strong baseline — keep tailoring for each role."})
        st.toast("Analysis saved to History", icon="✅")

    latest = st.session_state.get("latest_analysis")
    if latest:
        result = latest["result"]
        st.divider()
        st.markdown("<div class='report-header'><h2>Your insight report</h2><span class='report-badge'>● Analysis complete</span></div>", unsafe_allow_html=True)
        st.caption(f"Based on {latest['file_name']}" + (" and the target job description." if latest["has_jd"] else ". Add a job description next time for a role-match score."))
        score_cols = st.columns(3)
        with score_cols[0]: score_orb(result["score"], "overall readiness")
        with score_cols[1]: score_orb(result["core_score"], "resume quality")
        with score_cols[2]: score_orb(result["match_score"], "role match" if latest["has_jd"] else "add a target role")

        overview, alignment, checklist, strategy, toolkit = st.tabs(["Action plan", "Role alignment", "ATS review", "Strategy", "Career toolkit"])
        with overview:
            st.markdown("<div class='section-lead'><div class='section-eyebrow'>Your edit sequence</div><h3>Start with the changes that move the score.</h3><p>Work through these actions from top to bottom. Each recommendation comes from a visible signal in your uploaded resume.</p></div>", unsafe_allow_html=True)
            if result["suggestions"]:
                for index, (title, detail) in enumerate(result["suggestions"], 1):
                    st.markdown(f"<div class='step'><div class='step-num'>{index}</div><div><b>{html.escape(title)}</b><br><span class='small-muted'>{html.escape(detail)}</span></div></div>", unsafe_allow_html=True)
            else:
                st.success("Strong foundation. Tailor your wording and order of achievements for each role.")
            st.download_button("Download full report (.md)", report_markdown(result, latest["file_name"], latest.get("resume_text", ""), latest.get("job_description", "")), file_name="resume-insight-report.md", mime="text/markdown")
            with st.expander("Listen to your insight summary"):
                if TTS_AVAILABLE or GTTS_AVAILABLE:
                    voice_id = None
                    control_left, control_right, control_voice = st.columns(3)
                    with control_left:
                        rate = st.slider("Speech rate", 80, 250, 165, key="summary_rate")
                    with control_right:
                        volume = st.slider("Volume", 0.1, 1.0, 1.0, step=0.1, key="summary_volume")
                    if TTS_AVAILABLE:
                        try:
                            voice_engine = pyttsx3.init()
                            voices = voice_engine.getProperty("voices")
                            voice_labels = {voice.id: getattr(voice, "name", voice.id) for voice in voices}
                            with control_voice:
                                voice_id = st.selectbox("Voice", list(voice_labels), format_func=lambda item: voice_labels[item], key="summary_voice")
                            voice_engine.stop()
                        except Exception:
                            pass
                    if st.button("Generate voice summary", key="voice_summary"):
                        with st.spinner("Generating audio summary…"):
                            audio, audio_format = build_audio_summary(result, rate, volume, voice_id)
                            st.session_state["insight_audio"] = audio
                            st.session_state["insight_audio_format"] = audio_format
                        if not st.session_state["insight_audio"]:
                            st.warning("Audio could not be generated by the available voice engine.")
                    if st.session_state.get("insight_audio"):
                        st.audio(st.session_state["insight_audio"], format=st.session_state.get("insight_audio_format", "audio/wav"))
                else:
                    st.caption("Install pyttsx3 for local speech or gTTS for the online fallback.")
        with alignment:
            st.markdown("<div class='section-lead'><div class='section-eyebrow'>Target-role fit</div><h3>Compare your evidence with the job language.</h3><p>Use missing terms only when a genuine achievement, project, or skill can support them.</p></div>", unsafe_allow_html=True)
            if not latest["has_jd"]:
                st.info("Paste a target job description above, then create a new report to see matching and missing role terms.")
            else:
                left, right = st.columns(2)
                with left:
                    st.markdown("<div class='alignment-panel covered'><h4>Covered in your resume</h4><p>Terms already supported by your content.</p>" + tag_markup(result["matched_terms"]) + "</div>", unsafe_allow_html=True)
                with right:
                    st.markdown("<div class='alignment-panel gaps'><h4>Potentially missing terms</h4><p>Opportunities to add evidence, not keyword stuffing.</p>" + tag_markup(result["missing_terms"], missing=True) + "</div>", unsafe_allow_html=True)
                st.caption("Use only terms you can support with real experience. The goal is clarity, not keyword stuffing.")
        with checklist:
            st.markdown("<div class='section-lead'><div class='section-eyebrow'>Parser and recruiter check</div><h3>Make the important information easy to find.</h3><p>These signals reflect resume structure, contact completeness, evidence of impact, and scan-friendly language.</p></div>", unsafe_allow_html=True)
            left, right = st.columns([1.15, 1])
            with left:
                st.markdown("**Section completeness**")
                for label, value in result["sections"].items(): progress(label, value)
            with right:
                st.markdown("**Resume signals**")
                for label, passed in result["contact"].items():
                    st.write(("✅" if passed else "○") + f" {label}")
                st.write("")
                st.metric("Action verbs", result["action_count"], help="Strong verbs make achievement bullets easier to scan.")
                st.metric("Quantified outcomes", result["metrics_count"], help="A percentage, scale, money value, or concrete volume counts as evidence.")
                st.metric("Readability", f"{result['readability']}/100", help="Aim for concise, skimmable language rather than a perfect score.")
                st.metric("Experience years", result["exp_years"])
                st.metric("Project mentions", result["projects_count"])
        with strategy:
            st.markdown("<div class='section-lead'><div class='section-eyebrow'>Application strategy</div><h3>Turn the analysis into a focused narrative.</h3><p>Use this as a decision guide while revising your summary, skills ordering, and achievement bullets.</p></div>", unsafe_allow_html=True)
            strategy_text = html.escape(mock_ai_feedback(result)).replace("**", "")
            st.markdown("<div class='strategy-card'><p>" + strategy_text + "</p></div>", unsafe_allow_html=True)
            st.caption("Guidance is based on structural signals and the supplied job description. Validate every recommendation against your actual experience.")
        with toolkit:
            resume_source = latest.get("resume_text", "")
            job_source = latest.get("job_description", "")
            st.markdown("<div class='toolkit-heading'>Evidence and revision lab</div>", unsafe_allow_html=True)
            st.markdown("<p class='toolkit-subtitle'>Explore the evidence behind your score, test possible improvements, and compare a revised resume. These tools are read-only until you choose to update your document.</p>", unsafe_allow_html=True)
            st.markdown("""
                        <div class='tool-index'>
                            <div class='tool-index-card'><strong>1. Evidence traceability</strong><span>See which resume lines support each score.</span></div>
                            <div class='tool-index-card'><strong>2. Claim review</strong><span>Review certifications, expertise, and outcome claims.</span></div>
                            <div class='tool-index-card'><strong>3. Score simulator</strong><span>Preview the impact of possible improvements.</span></div>
                            <div class='tool-index-card'><strong>4. Compare versions</strong><span>Measure changes between two resume drafts.</span></div>
                            <div class='tool-index-card'><strong>5. Gap coach</strong><span>Turn missing role terms into evidence-first templates.</span></div>
                            <div class='tool-index-card'><strong>6. ATS scan</strong><span>Check parser-friendly structure and keyword balance.</span></div>
                            <div class='tool-index-card'><strong>7. Interview practice</strong><span>Generate role-focused questions from your analysis.</span></div>
                            <div class='tool-index-card'><strong>8. Professional summary</strong><span>Draft a truthful, role-aligned summary.</span></div>
                            <div class='tool-index-card'><strong>9. Skill roadmap</strong><span>Prioritize skills explicitly requested by the role.</span></div>
                        </div>
            """, unsafe_allow_html=True)
            st.markdown("<div class='visibility-note'><b>All nine tools are available below.</b> Their titles and purpose are listed above; open only the tool you need. Evidence traceability and gap coaching are expanded first because they are the most useful places to start.</div>", unsafe_allow_html=True)
            if not latest["has_jd"]:
                st.info("Add a target job description and create a new report to unlock role-specific gap coaching, interview questions, and a skills roadmap.")

            with st.expander("1. Evidence traceability", expanded=True):
                st.caption("See the resume lines behind each local score signal. This view explains the analysis without changing your document.")
                for signal, status, matches, explanation in evidence_trace(resume_source, job_source, result):
                    st.markdown(f"**{signal}** · {status}")
                    st.caption(explanation)
                    if matches:
                        st.code("\n".join(matches), language=None)
                    else:
                        st.caption("No supporting line was detected.")

            with st.expander("2. Claim review", expanded=False):
                st.caption("These are review prompts, not accusations. Confirm that strong claims are supported by real evidence.")
                findings = claim_review(resume_source)
                if not findings:
                    st.success("No high-attention claims were detected by the local review rules.")
                else:
                    for label, line in findings:
                        st.markdown(f"**{label}**")
                        st.write(line)

            with st.expander("3. Score simulator", expanded=False):
                st.caption("Preview the possible effect of improvements. Nothing here changes the original resume or saved history.")
                simulate_left, simulate_right = st.columns(2)
                with simulate_left:
                    simulate_sections = st.checkbox("Complete missing sections", key="simulate_sections")
                    simulate_contact = st.checkbox("Complete missing contact details", key="simulate_contact")
                with simulate_right:
                    simulate_impact = st.checkbox("Add stronger action verbs and 3 measurable outcomes", key="simulate_impact")
                    simulate_role = st.checkbox("Address all missing role terms truthfully", key="simulate_role", disabled=not latest["has_jd"])
                projected = simulated_score(result, {"sections": simulate_sections, "contact": simulate_contact, "impact": simulate_impact, "role": simulate_role})
                sim_a, sim_b, sim_c = st.columns(3)
                sim_a.metric("Projected overall", f"{projected[0]}/100", delta=projected[0] - result["score"])
                sim_b.metric("Projected quality", f"{projected[1]}/100", delta=projected[1] - result["core_score"])
                sim_c.metric("Projected role match", f"{projected[2]}/100", delta=projected[2] - result["match_score"] if latest["has_jd"] else None)

            with st.expander("4. Compare resume versions", expanded=False):
                st.caption("Upload a revised PDF/DOCX or paste a revision to see score movement and content changes side by side.")
                comparison_file = st.file_uploader("Revised resume", type=["pdf", "docx"], key="comparison_resume")
                revised_text = ""
                if comparison_file:
                    try:
                        revised_text = extract_text(comparison_file.getvalue(), comparison_file.name)
                    except Exception:
                        st.error("The revised file could not be read. Try another PDF or DOCX.")
                if not revised_text:
                    revised_text = st.text_area("Or paste revised resume text", height=150, key="comparison_text")
                if revised_text.strip():
                    comparison = version_comparison(resume_source, revised_text, job_source)
                    original_version, revised_version = comparison["original"], comparison["revised"]
                    version_a, version_b, version_c = st.columns(3)
                    version_a.metric("Overall score", f"{revised_version['score']}/100", delta=revised_version["score"] - original_version["score"])
                    version_b.metric("Role match", f"{revised_version['match_score']}/100", delta=revised_version["match_score"] - original_version["match_score"] if latest["has_jd"] else None)
                    version_c.metric("Readability", f"{revised_version['readability']}/100", delta=revised_version["readability"] - original_version["readability"])
                    if comparison["added_sections"]:
                        st.markdown("**Sections added:** " + ", ".join(comparison["added_sections"]))
                    if comparison["removed_sections"]:
                        st.markdown("**Sections removed:** " + ", ".join(comparison["removed_sections"]))
                    if latest["has_jd"]:
                        st.markdown("**Role terms added:**")
                        tags(comparison["added_terms"])
                        st.markdown("**Role terms removed:**")
                        tags(comparison["removed_terms"], missing=True)
                    if comparison["changed_lines"]:
                        with st.expander("View text changes"):
                            st.code("\n".join(comparison["changed_lines"]), language="diff")
                else:
                    st.info("Add a revised version to activate comparison.")

            with st.expander("5. Resume-to-job gap coach", expanded=True):
                rewrites = gap_coach(resume_source, result) if latest["has_jd"] else []
                if not rewrites:
                    st.success("No priority role terms are missing from this analysis. Focus on making the evidence behind your strongest matches more specific.")
                else:
                    st.caption("Use these as evidence-first templates. Replace every bracketed prompt with a fact from your own experience.")
                    for term, current, rewrite in rewrites:
                        st.markdown(f"**Opportunity: {term.title()}**")
                        st.caption("Starting point from your resume")
                        st.write(current[:420])
                        st.caption("Truthful rewrite template")
                        st.code(rewrite, language=None)

            with st.expander("6. ATS scan simulator"):
                ats = ats_scan(resume_source, job_source, result)
                ats_a, ats_b, ats_c = st.columns(3)
                ats_a.metric("Extracted words", ats["word_count"])
                ats_b.metric("Role-term mentions", ats["keyword_hits"])
                ats_c.metric("Keyword density", f"{ats['density']}%")
                st.caption("Healthy resumes use relevant terms naturally in the summary, skills, and evidence bullets — not in a keyword list.")
                st.markdown("**Likely parser and scan checks**")
                for title, detail in ats["issues"]:
                    st.markdown(f"- **{title}:** {detail}")
                st.markdown("**Detected ATS-friendly section order**")
                st.write(" → ".join(ats["section_order"]) if ats["section_order"] else "Use clearly labeled Experience, Education, Skills, and Projects sections.")

            with st.expander("7. Interview question generator"):
                questions = interview_questions(result) if latest["has_jd"] else []
                if not questions:
                    st.info("Add a job description to generate questions tailored to the role.")
                else:
                    st.caption("Practice with STAR: Situation, Task, Action, Result. Prepare a real example before answering each question.")
                    for number, question in enumerate(questions, 1):
                        st.markdown(f"**{number}.** {question}")

            with st.expander("8. Tailored professional summary"):
                summary = tailored_summary(result)
                st.caption("A starting draft — edit it until every statement is accurate, specific, and written in your own voice.")
                st.text_area("Suggested summary", value=summary, height=130, key="tailored_summary")
                st.download_button("Download summary (.txt)", summary, file_name="tailored-professional-summary.txt", mime="text/plain", key="download_summary")

            with st.expander("9. Skill roadmap"):
                roadmap = skill_roadmap(job_source, result) if latest["has_jd"] else []
                if not roadmap:
                    st.info("Add a job description to identify the role's requested technical skills and prioritize a roadmap.")
                else:
                    st.caption("Priorities are based on skills explicitly mentioned in this job description, not generic market trends.")
                    for skill, priority, action in roadmap:
                        st.markdown(f"**{priority} · {skill.title()}**")
                        st.write(action)

        bottom_left, bottom_right = st.columns([1, 1.3])
        with bottom_left:
            st.markdown("#### Detected skills")
            tags(result["skills"])
        with bottom_right:
            st.markdown("#### Skill signal")
            if result["skills"]:
                cloud = WordCloud(width=700, height=230, background_color="white", colormap="viridis").generate_from_frequencies({skill: 1 for skill in result["skills"]})
                figure, axis = plt.subplots(figsize=(8, 2.6))
                axis.imshow(cloud, interpolation="bilinear")
                axis.axis("off")
                st.pyplot(figure, use_container_width=True)
                plt.close(figure)
            else:
                st.info("Add a dedicated Skills section to make technical strengths easier to spot.")

else:
    st.markdown("<section class='page-intro'><h1>Your progress</h1><p>A private timeline of saved analyses. Use it to see whether your revisions are improving resume quality and role alignment.</p></section>", unsafe_allow_html=True)
    history = [entry for entry in load_history() if all(key in entry for key in ("timestamp", "file_name", "overall_score", "match_score", "readability"))]
    if not history:
        st.info("No saved analyses yet. Your first report will appear here after you analyze a resume.")
    else:
        history = sorted(history, key=lambda item: item["timestamp"])
        latest_entry, previous = history[-1], history[-2] if len(history) > 1 else None
        st.markdown("#### Latest snapshot")
        c1, c2, c3 = st.columns(3)
        c1.metric("Latest overall score", f"{latest_entry['overall_score']}/100", delta=(latest_entry["overall_score"] - previous["overall_score"] if previous else None))
        c2.metric("Latest role match", f"{latest_entry['match_score']}/100", delta=(latest_entry["match_score"] - previous["match_score"] if previous else None))
        c3.metric("Analyses tracked", len(history))
        try:
            import pandas as pd
            chart_data = pd.DataFrame(history)
            chart_data["date"] = pd.to_datetime(chart_data["timestamp"])
            st.markdown("#### Score trend")
            st.line_chart(chart_data, x="date", y=["overall_score", "match_score"], use_container_width=True)
        except Exception:
            st.caption("Trend chart is unavailable, but your saved entries remain below.")
        st.markdown("#### All saved reports")
        st.caption("Every saved analysis is shown here. Open a report below only when you want its full detail.")
        for entry in reversed(history):
            timestamp = entry["timestamp"].replace("T", " ")
            st.markdown(
                "<div class='history-card'><strong>"
                + html.escape(entry["file_name"])
                + "</strong><div class='history-meta'>"
                + html.escape(timestamp)
                + " &nbsp;•&nbsp; Overall <b>"
                + str(entry["overall_score"])
                + "/100</b> &nbsp;•&nbsp; Role match <b>"
                + str(entry["match_score"])
                + "/100</b> &nbsp;•&nbsp; Readability <b>"
                + str(entry["readability"])
                + "/100</b></div></div>",
                unsafe_allow_html=True,
            )
        st.markdown("#### Analysis timeline")
        for index, entry in enumerate(reversed(history)):
            timestamp = entry["timestamp"].replace("T", " ")
            with st.expander(f"{entry['file_name']}  ·  {entry['overall_score']}/100  ·  {timestamp}", expanded=index == 0):
                a, b, c = st.columns(3)
                a.metric("Overall", f"{entry['overall_score']}/100")
                b.metric("Role match", f"{entry['match_score']}/100")
                c.metric("Readability", f"{entry['readability']}/100")
                st.caption(entry.get("feedback_summary", "No summary saved."))
                tags(entry.get("top_skills", []))
        st.divider()
        export, clear = st.columns(2)
        with export:
            st.markdown("**Keep a copy of your progress**")
            st.caption("Download your complete private history before deleting anything.")
            st.download_button("Export history (.json)", json.dumps(history, indent=2), file_name="resume-insight-history.json", mime="application/json")
        with clear:
            st.markdown("<div class='danger-zone'><strong>Clear saved history</strong><span>This permanently removes all saved reports and the current analysis from this local app.</span></div>", unsafe_allow_html=True)
            confirm_clear = st.checkbox("I understand this cannot be undone", key="confirm_clear_history")
            if st.button("Permanently clear history", type="secondary", disabled=not confirm_clear, use_container_width=True, key="clear_history_button"):
                HISTORY_PATH.write_text("[]", encoding="utf-8")
                st.session_state.pop("latest_analysis", None)
                st.toast("Saved history cleared", icon="✅")
                st.rerun()
