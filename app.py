"""
HireAi — AI-Powered Resume Intelligence Platform
Production-ready Streamlit SaaS app for Streamlit Cloud deployment.
Supports single resume analysis AND batch ZIP processing with comparison dashboard.
"""

import streamlit as st
import json
import re
import io
import os
import csv
import hashlib
import zipfile
from datetime import datetime

import pdfplumber
import docx
import pandas as pd
from groq import Groq
from fpdf import FPDF

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="HireAi",
    page_icon="🔮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# GLOBAL CSS — HireAi Purple Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Space Grotesk', sans-serif;
    background: #0C0516;
    color: #E8EAF0;
}

::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: #0C0516; }
::-webkit-scrollbar-thumb { background: #5416B5; border-radius: 3px; }

.main { background: #0C0516; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0F083B 0%, #0C0516 100%);
    border-right: 1px solid #2A1A4A;
}

h1, h2, h3 { color: #7F3AA1; font-weight: 700; }

.hire-card {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 16px;
    padding: 24px;
    margin: 10px 0;
}

.score-card {
    background: linear-gradient(135deg, #0F083B, #5416B5 200%);
    border: 1px solid #5416B5;
    border-radius: 16px;
    padding: 28px 16px;
    text-align: center;
    margin: 8px 0;
    transition: transform 0.2s;
}
.score-card:hover { transform: translateY(-3px); }
.score-num {
    font-size: 3rem;
    font-weight: 700;
    line-height: 1;
}
.score-lbl {
    color: #A090C0;
    font-size: .75rem;
    margin-top: 8px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    font-family: 'DM Mono', monospace;
}

.tag {
    display: inline-block;
    background: #1A0F4A;
    color: #C89AF0;
    border: 1px solid #5416B5;
    padding: 4px 14px;
    border-radius: 20px;
    font-size: .78rem;
    margin: 3px;
    font-family: 'DM Mono', monospace;
}
.tag-miss {
    background: #2A0A1A;
    border-color: #8A2060;
    color: #F090A0;
}

.imp-item {
    background: #0F083B;
    border-left: 3px solid #7F3AA1;
    padding: 12px 16px;
    margin: 8px 0;
    border-radius: 0 10px 10px 0;
    font-size: .9rem;
    line-height: 1.5;
}

.sec-header {
    font-size: 1.2rem;
    font-weight: 700;
    color: #C89AF0;
    margin: 28px 0 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid #2A1A4A;
    text-transform: uppercase;
    letter-spacing: 1px;
    font-family: 'DM Mono', monospace;
}

.brand-header {
    font-size: 2.4rem;
    font-weight: 700;
    background: linear-gradient(90deg, #7F3AA1, #C89AF0);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    letter-spacing: -1px;
}
.brand-sub {
    color: #7060A0;
    font-size: .95rem;
    margin-top: 4px;
}

.stButton > button {
    background: linear-gradient(135deg, #7F3AA1, #5416B5);
    color: #E8EAF0;
    font-weight: 700;
    border: none;
    border-radius: 10px;
    padding: 10px 28px;
    font-size: 1rem;
    transition: opacity 0.2s, transform 0.2s;
    font-family: 'Space Grotesk', sans-serif;
}
.stButton > button:hover { opacity: 0.85; transform: translateY(-1px); }

.stTextArea textarea {
    background: #0F083B;
    color: #E8EAF0;
    border: 1px solid #3A1A6A;
    border-radius: 10px;
    font-family: 'Space Grotesk', sans-serif;
}

[data-testid="stFileUploader"] {
    background: #0F083B;
    border: 1px dashed #5416B5;
    border-radius: 12px;
    padding: 8px;
}

[data-testid="stMetric"] {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 12px;
    padding: 16px;
}

[data-testid="stTab"] { color: #A090C0; }

.stAlert { border-radius: 10px; }

.user-badge {
    background: linear-gradient(135deg, #1A0F4A, #2A1060);
    border: 1px solid #5416B5;
    border-radius: 12px;
    padding: 14px 16px;
    text-align: center;
    margin-bottom: 16px;
}
.user-badge-name { font-weight: 700; color: #C89AF0; font-size: 1rem; }
.user-badge-email { color: #7060A0; font-size: .8rem; margin-top: 2px; }

.stat-pill {
    background: linear-gradient(135deg, #0F083B, #5416B5 250%);
    border: 1px solid #5416B5;
    border-radius: 12px;
    padding: 16px;
    text-align: center;
}
.stat-pill-num { font-size: 2rem; font-weight: 700; color: #C89AF0; }
.stat-pill-lbl { font-size: .75rem; color: #7060A0; text-transform: uppercase; letter-spacing: 1px; font-family: 'DM Mono', monospace; }

.hist-row {
    background: #0F083B;
    border: 1px solid #2A1A4A;
    border-radius: 10px;
    padding: 14px 18px;
    margin: 6px 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.job-card {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 14px;
    padding: 20px;
    margin: 10px 0;
}
.job-card-title { font-weight: 700; color: #C89AF0; font-size: 1rem; }
.job-card-meta { color: #7060A0; font-size: .8rem; margin-top: 4px; font-family: 'DM Mono', monospace; }

.purple-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #5416B5, transparent);
    margin: 24px 0;
}

/* Batch comparison table */
[data-testid="stDataFrame"] {
    border: 1px solid #3A1A6A;
    border-radius: 12px;
    overflow: hidden;
}

/* Top pick badge */
.top-pick-badge {
    display: inline-block;
    background: linear-gradient(135deg, #7F3AA1, #C89AF0);
    color: #0C0516;
    font-weight: 700;
    font-size: .72rem;
    padding: 3px 10px;
    border-radius: 20px;
    margin-left: 8px;
    font-family: 'DM Mono', monospace;
    letter-spacing: 0.5px;
}

/* Progress bar color override */
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, #7F3AA1, #C89AF0);
}

/* Batch result card */
.batch-card {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 14px;
    padding: 18px 22px;
    margin: 8px 0;
}
.batch-rank {
    font-size: 1.4rem;
    font-weight: 700;
    color: #5416B5;
    font-family: 'DM Mono', monospace;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────
def init_session():
    defaults = {
        "logged_in": False,
        "user_name": "",
        "user_email": "",
        "analysis_history": [],
        "saved_jobs": [],
        "current_tab": "Analyzer",
        "users_db": {},
        "batch_results": [],       # list of {filename, analysis} dicts from batch run
        "last_single_analysis": None,
        "last_single_resume_text": "",
        "last_single_opt_cv": "",
        "last_single_jd_title": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session()

# ─────────────────────────────────────────────
# GROQ CLIENT (cached)
# ─────────────────────────────────────────────
@st.cache_resource
def get_groq_client():
    api_key = st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", ""))
    if not api_key:
        st.error("⚠️ GROQ_API_KEY not found. Add it to Streamlit Cloud Secrets.")
        st.stop()
    return Groq(api_key=api_key)

MODEL = "llama-3.3-70b-versatile"

# ─────────────────────────────────────────────
# AUTH HELPERS
# ─────────────────────────────────────────────
def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def register_user(name: str, email: str, pw: str) -> bool:
    if email in st.session_state.users_db:
        return False
    st.session_state.users_db[email] = {"pw_hash": hash_pw(pw), "name": name}
    return True

def login_user(email: str, pw: str) -> bool:
    user = st.session_state.users_db.get(email)
    if user and user["pw_hash"] == hash_pw(pw):
        st.session_state.logged_in = True
        st.session_state.user_name = user["name"]
        st.session_state.user_email = email
        return True
    return False

# ─────────────────────────────────────────────
# CORE HELPERS
# ─────────────────────────────────────────────
def clean_text(text: str) -> str:
    """Sanitize text for FPDF (latin-1 safe, no UnicodeEncodeError)."""
    if not isinstance(text, str):
        text = str(text)
    return text.encode("latin-1", "replace").decode("latin-1")


@st.cache_data(show_spinner=False)
def extract_text(file_bytes: bytes, file_name: str) -> str:
    """Extract plain text from PDF, DOCX, or TXT bytes."""
    buf = io.BytesIO(file_bytes)
    lower = file_name.lower()
    try:
        if lower.endswith(".pdf"):
            with pdfplumber.open(buf) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages)
        if lower.endswith(".docx"):
            d = docx.Document(buf)
            return "\n".join(p.text for p in d.paragraphs)
    except Exception as e:
        return f"[extraction error: {e}]"
    return file_bytes.decode("utf-8", errors="ignore")


def extract_from_zip(zip_bytes: bytes) -> list[dict]:
    """
    Open a ZIP archive and extract text from every PDF/DOCX/TXT inside.
    Returns a list of {filename, text} dicts.
    """
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for name in z.namelist():
                lower = name.lower()
                # skip macOS metadata and directories
                if name.startswith("__MACOSX") or name.endswith("/"):
                    continue
                if not any(lower.endswith(ext) for ext in (".pdf", ".docx", ".txt")):
                    continue
                try:
                    file_bytes = z.read(name)
                    basename = os.path.basename(name)
                    text = extract_text(file_bytes, basename)
                    if len(text.strip()) >= 50:
                        results.append({"filename": basename, "text": text})
                except Exception:
                    pass
    except Exception as e:
        st.error(f"Could not open ZIP archive: {e}")
    return results


def ask_groq(sys_p: str, usr_p: str, temp: float = 0.3) -> str:
    """Call Groq API with full error handling — never crashes the app."""
    try:
        client = get_groq_client()
        r = client.chat.completions.create(
            model=MODEL,
            temperature=temp,
            max_tokens=4096,
            messages=[
                {"role": "system", "content": sys_p},
                {"role": "user", "content": usr_p},
            ],
        )
        return r.choices[0].message.content.strip()
    except Exception as e:
        return f"[API_ERROR: {e}]"


def parse_json(text: str):
    """Extract and parse JSON from LLM output, with multiple fallback strategies."""
    if text.startswith("[API_ERROR"):
        return None
    # Try fenced code block first
    m = re.search(r"```json\s*([\s\S]+?)```", text)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    # Try raw JSON object / array
    s = text.find("{") if "{" in text else text.find("[")
    if s != -1:
        try:
            return json.loads(text[s:])
        except Exception:
            pass
    # Last resort: parse as-is
    try:
        return json.loads(text)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def analyze_resume(resume: str, jd: str) -> dict:
    """Run ATS analysis via LLaMA 3.3-70B and return structured dict."""
    schema = (
        '{"ats_score":<int>,"keyword_match_score":<int>,"format_score":<int>,'
        '"readability_score":<int>,"candidate_name":"","email":"","phone":"",'
        '"location":"","linkedin":"","years_experience":"","current_title":"",'
        '"education":"","skills":[],"matched_keywords":[],"missing_keywords":[],'
        '"strengths":[],"improvements":[{"priority":"HIGH","issue":"","suggestion":""}],'
        '"overall_recommendation":"","summary":""}'
    )
    sys_p = (
        "You are an expert ATS resume analyst and career coach. "
        "Reply ONLY with a valid JSON object matching this schema exactly: " + schema
    )
    usr_p = "RESUME:\n" + resume + "\n\nJOB DESCRIPTION:\n" + jd
    raw = ask_groq(sys_p, usr_p)
    data = parse_json(raw)
    return data if isinstance(data, dict) else {}


@st.cache_data(show_spinner=False)
def gen_cv(resume: str, jd: str, analysis: dict) -> str:
    """Generate ATS-optimized CV text."""
    missing = ", ".join(analysis.get("missing_keywords", []))
    sys_p = (
        "You are a professional resume writer specializing in ATS-optimized CVs. "
        "Generate a complete, polished resume in plain text using this structure: "
        "NAME | Email | Phone | LinkedIn | Location\n"
        "PROFESSIONAL SUMMARY\nCORE COMPETENCIES\n"
        "PROFESSIONAL EXPERIENCE (bullet achievements with metrics)\n"
        "EDUCATION\nCERTIFICATIONS & SKILLS\n"
        "Naturally weave in these missing keywords: " + missing + ". "
        "Use strong action verbs. Output ONLY the resume text."
    )
    usr_p = "Original resume:\n" + resume + "\n\nJob description:\n" + jd
    result = ask_groq(sys_p, usr_p, temp=0.5)
    return result if not result.startswith("[API_ERROR") else "CV generation failed. Please retry."


def score_color(s: int) -> str:
    if s >= 75:
        return "#A0F080"
    if s >= 50:
        return "#F0D070"
    return "#F08090"


# ─────────────────────────────────────────────
# PDF EXPORT — SINGLE CANDIDATE
# ─────────────────────────────────────────────
def build_analysis_pdf(analysis: dict, opt_cv: str) -> bytes:
    """Generate a polished PDF report. All strings pass through clean_text()."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # ── Page 1: Analysis Report ──
    pdf.add_page()

    # Header bar
    pdf.set_fill_color(15, 8, 59)
    pdf.rect(0, 0, 210, 30, "F")
    pdf.set_xy(10, 8)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 14, clean_text("HireAi -- Resume Analysis Report"), ln=True)

    pdf.set_xy(10, 22)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(112, 96, 160)
    pdf.cell(0, 6, clean_text(f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}"), ln=True)

    pdf.ln(10)

    # Candidate Info
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, clean_text("Candidate Profile"), ln=True)
    pdf.set_draw_color(84, 22, 181)
    pdf.set_line_width(0.4)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    fields = [
        ("Name", "candidate_name"), ("Email", "email"), ("Phone", "phone"),
        ("Location", "location"), ("LinkedIn", "linkedin"),
        ("Years of Experience", "years_experience"),
        ("Current Title", "current_title"), ("Education", "education"),
    ]
    pdf.set_font("Helvetica", "", 10)
    for label, key in fields:
        val = analysis.get(key, "-") or "-"
        pdf.set_text_color(112, 96, 160)
        pdf.cell(55, 7, clean_text(f"  {label}:"), ln=False)
        pdf.set_text_color(232, 234, 240)
        pdf.cell(0, 7, clean_text(str(val)), ln=True)

    pdf.ln(6)

    # Scores
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, clean_text("ATS Score Summary"), ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)

    score_items = [
        ("Overall ATS Score", "ats_score"),
        ("Keyword Match", "keyword_match_score"),
        ("Format Score", "format_score"),
        ("Readability Score", "readability_score"),
    ]
    pdf.set_font("Helvetica", "", 11)
    for label, key in score_items:
        val = analysis.get(key, 0)
        pdf.set_text_color(112, 96, 160)
        pdf.cell(90, 8, clean_text(f"  {label}"), ln=False)
        if val >= 75:
            pdf.set_text_color(100, 220, 80)
        elif val >= 50:
            pdf.set_text_color(220, 200, 70)
        else:
            pdf.set_text_color(220, 80, 100)
        pdf.cell(0, 8, clean_text(f"{val} / 100"), ln=True)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(160, 200, 100)
    rec = analysis.get("overall_recommendation", "")
    pdf.cell(0, 8, clean_text(f"  Recommendation: {rec}"), ln=True)

    pdf.ln(3)
    pdf.set_font("Helvetica", "I", 10)
    pdf.set_text_color(180, 170, 210)
    summary = analysis.get("summary", "")
    pdf.multi_cell(0, 6, clean_text(f"  {summary}"))

    pdf.ln(6)

    # Keywords
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, clean_text("Keyword Analysis"), ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(100, 220, 80)
    pdf.cell(0, 7, clean_text("  Matched Keywords:"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    matched = ", ".join(analysis.get("matched_keywords", []))
    pdf.multi_cell(0, 6, clean_text(f"  {matched or 'None identified'}"))

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(220, 80, 100)
    pdf.cell(0, 7, clean_text("  Missing Keywords:"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    missing = ", ".join(analysis.get("missing_keywords", []))
    pdf.multi_cell(0, 6, clean_text(f"  {missing or 'None identified'}"))

    pdf.ln(4)

    # Strengths
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, clean_text("Strengths"), ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    for s in analysis.get("strengths", []):
        pdf.cell(0, 7, clean_text(f"  [+]  {s}"), ln=True)

    pdf.ln(4)

    # Improvements
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, clean_text("Improvement Recommendations"), ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)
    icons = {"HIGH": "[HIGH]", "MEDIUM": "[MED]", "LOW": "[LOW]"}
    pdf.set_font("Helvetica", "", 10)
    for imp in analysis.get("improvements", []):
        p = imp.get("priority", "MEDIUM")
        if p == "HIGH":
            pdf.set_text_color(220, 80, 100)
        elif p == "MEDIUM":
            pdf.set_text_color(220, 200, 70)
        else:
            pdf.set_text_color(100, 220, 80)
        pdf.cell(20, 7, clean_text(f"  {icons.get(p, '')}"), ln=False)
        pdf.set_text_color(232, 234, 240)
        issue = clean_text(imp.get("issue", ""))
        suggestion = clean_text(imp.get("suggestion", ""))
        pdf.multi_cell(0, 7, f"{issue}  ->  {suggestion}")
        pdf.ln(1)

    # ── Page 2: Optimized CV ──
    if opt_cv and not opt_cv.startswith("CV generation failed"):
        pdf.add_page()
        pdf.set_fill_color(15, 8, 59)
        pdf.rect(0, 0, 210, 30, "F")
        pdf.set_xy(10, 8)
        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 14, clean_text("HireAi -- ATS-Optimized CV"), ln=True)

        pdf.set_xy(10, 22)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(112, 96, 160)
        pdf.cell(0, 6, clean_text("AI-generated, keyword-optimized resume template"), ln=True)

        pdf.ln(8)
        pdf.set_font("Courier", "", 9)
        pdf.set_text_color(220, 214, 240)
        for line in opt_cv.split("\n"):
            pdf.multi_cell(0, 5, clean_text(line)[:200])  # truncate all lines to 200 chars

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)


# ─────────────────────────────────────────────
# PDF EXPORT — BATCH TOP-3 MERGED REPORT
# ─────────────────────────────────────────────
def build_batch_pdf(batch_results: list, jd_title: str) -> bytes:
    """Generate a merged PDF report for the top 3 candidates from a batch."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    top3 = sorted(batch_results, key=lambda x: x["analysis"].get("ats_score", 0), reverse=True)[:3]

    for rank, item in enumerate(top3, 1):
        analysis = item["analysis"]
        filename = item["filename"]

        pdf.add_page()

        # Header
        pdf.set_fill_color(15, 8, 59)
        pdf.rect(0, 0, 210, 30, "F")
        pdf.set_xy(10, 8)
        pdf.set_font("Helvetica", "B", 18)
        pdf.set_text_color(200, 154, 240)
        badge = " [TOP PICK]" if rank == 1 else f" [#{rank}]"
        pdf.cell(0, 14, clean_text(f"HireAi Batch Report{badge}"), ln=True)

        pdf.set_xy(10, 22)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(112, 96, 160)
        pdf.cell(0, 6, clean_text(
            f"Role: {jd_title or 'N/A'}  |  File: {filename}  |  Rank: #{rank} of {len(batch_results)}"
        ), ln=True)

        pdf.ln(8)

        # Candidate info
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 8, clean_text("Candidate Profile"), ln=True)
        pdf.set_draw_color(84, 22, 181)
        pdf.set_line_width(0.4)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(3)

        fields = [
            ("Name", "candidate_name"), ("Email", "email"), ("Phone", "phone"),
            ("Location", "location"), ("Experience", "years_experience"),
            ("Title", "current_title"), ("Education", "education"),
        ]
        pdf.set_font("Helvetica", "", 10)
        for label, key in fields:
            val = analysis.get(key, "-") or "-"
            pdf.set_text_color(112, 96, 160)
            pdf.cell(45, 7, clean_text(f"  {label}:"), ln=False)
            pdf.set_text_color(232, 234, 240)
            pdf.cell(0, 7, clean_text(str(val)), ln=True)

        pdf.ln(4)

        # Scores
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 8, clean_text("ATS Scores"), ln=True)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(3)

        for label, key in [
            ("Overall ATS", "ats_score"), ("Keyword Match", "keyword_match_score"),
            ("Format", "format_score"), ("Readability", "readability_score"),
        ]:
            val = analysis.get(key, 0)
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(112, 96, 160)
            pdf.cell(70, 7, clean_text(f"  {label}"), ln=False)
            if val >= 75:
                pdf.set_text_color(100, 220, 80)
            elif val >= 50:
                pdf.set_text_color(220, 200, 70)
            else:
                pdf.set_text_color(220, 80, 100)
            pdf.cell(0, 7, clean_text(f"{val} / 100"), ln=True)

        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(160, 200, 100)
        pdf.cell(0, 7, clean_text(f"  Recommendation: {analysis.get('overall_recommendation', '')}"), ln=True)

        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(180, 170, 210)
        pdf.multi_cell(0, 5, clean_text(f"  {analysis.get('summary', '')}"))

        pdf.ln(3)

        # Skills
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 7, clean_text("Skills"), ln=True)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(232, 234, 240)
        skills_str = ", ".join(analysis.get("skills", []))
        pdf.multi_cell(0, 6, clean_text(f"  {skills_str or 'None extracted'}"))

        pdf.ln(3)

        # Top improvements
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 7, clean_text("Key Improvement Areas"), ln=True)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
        pdf.set_font("Helvetica", "", 9)
        for imp in analysis.get("improvements", [])[:3]:
            p = imp.get("priority", "")
            if p == "HIGH":
                pdf.set_text_color(220, 80, 100)
            elif p == "MEDIUM":
                pdf.set_text_color(220, 200, 70)
            else:
                pdf.set_text_color(100, 220, 80)
            issue = clean_text(imp.get("issue", ""))
            sug = clean_text(imp.get("suggestion", ""))
            pdf.multi_cell(0, 5, f"  [{p}] {issue} -> {sug}")
            pdf.ln(1)

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)


# ─────────────────────────────────────────────
# CSV EXPORT — BATCH SUMMARY
# ─────────────────────────────────────────────
def build_batch_csv(batch_results: list) -> bytes:
    """Build a summary CSV of all batch-analyzed candidates."""
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Rank", "Filename", "Candidate Name", "ATS Score",
        "Keyword Match", "Format Score", "Readability Score",
        "Overall Recommendation", "Years Experience",
        "Current Title", "Matched Keywords", "Missing Keywords",
        "Summary",
    ])
    sorted_results = sorted(
        batch_results,
        key=lambda x: x["analysis"].get("ats_score", 0),
        reverse=True
    )
    for rank, item in enumerate(sorted_results, 1):
        a = item["analysis"]
        writer.writerow([
            rank,
            item["filename"],
            a.get("candidate_name", ""),
            a.get("ats_score", 0),
            a.get("keyword_match_score", 0),
            a.get("format_score", 0),
            a.get("readability_score", 0),
            a.get("overall_recommendation", ""),
            a.get("years_experience", ""),
            a.get("current_title", ""),
            "; ".join(a.get("matched_keywords", [])),
            "; ".join(a.get("missing_keywords", [])),
            a.get("summary", ""),
        ])
    return output.getvalue().encode("utf-8")


# ─────────────────────────────────────────────
# RENDER: SINGLE ANALYSIS RESULTS
# ─────────────────────────────────────────────
def render_single_results(analysis: dict, resume_text: str, jd_text: str, jd_title: str):
    """Render the full results panel for a single resume analysis."""
    # ── Scores ──
    st.markdown('<div class="sec-header">🎯 ATS Scores</div>', unsafe_allow_html=True)
    s1, s2, s3, s4 = st.columns(4)
    for col, lbl, key in [
        (s1, "Overall ATS", "ats_score"),
        (s2, "Keyword Match", "keyword_match_score"),
        (s3, "Format", "format_score"),
        (s4, "Readability", "readability_score"),
    ]:
        v = analysis.get(key, 0)
        col.markdown(
            f'<div class="score-card">'
            f'<div class="score-num" style="color:{score_color(v)}">{v}</div>'
            f'<div class="score-lbl">{lbl}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    rec = analysis.get("overall_recommendation", "")
    rec_colors = {
        "Highly Recommended": "#A0F080",
        "Recommended": "#80C060",
        "Needs Improvement": "#F0D070",
        "Not Recommended": "#F08090",
    }
    rc = rec_colors.get(rec, "#C89AF0")
    st.markdown(
        f'**Recommendation:** <span style="color:{rc};font-weight:700">{rec}</span>',
        unsafe_allow_html=True,
    )
    st.markdown(f"> {analysis.get('summary', '')}")

    # ── Candidate Info ──
    st.markdown('<div class="sec-header">👤 Candidate Info</div>', unsafe_allow_html=True)
    ci1, ci2 = st.columns(2)
    info_fields = [
        ("Name", "candidate_name"), ("Email", "email"),
        ("Phone", "phone"), ("Location", "location"),
        ("LinkedIn", "linkedin"), ("Experience", "years_experience"),
        ("Current Title", "current_title"), ("Education", "education"),
    ]
    for i, (lbl, k) in enumerate(info_fields):
        v = analysis.get(k, "—") or "—"
        (ci1 if i % 2 == 0 else ci2).markdown(f"**{lbl}:** {v}")

    # ── Skills ──
    st.markdown('<div class="sec-header">🛠 Skills</div>', unsafe_allow_html=True)
    skills_html = " ".join(f'<span class="tag">{s}</span>' for s in analysis.get("skills", []))
    st.markdown(skills_html or "No skills extracted.", unsafe_allow_html=True)

    # ── Keywords ──
    st.markdown('<div class="sec-header">🔑 Keyword Analysis</div>', unsafe_allow_html=True)
    kw1, kw2 = st.columns(2)
    with kw1:
        st.markdown("**✅ Matched Keywords**")
        matched_html = " ".join(f'<span class="tag">{k}</span>' for k in analysis.get("matched_keywords", []))
        st.markdown(matched_html or "None identified.", unsafe_allow_html=True)
    with kw2:
        st.markdown("**❌ Missing Keywords**")
        miss_html = " ".join(f'<span class="tag tag-miss">{k}</span>' for k in analysis.get("missing_keywords", []))
        st.markdown(miss_html or "None identified.", unsafe_allow_html=True)

    # ── Strengths ──
    st.markdown('<div class="sec-header">💪 Strengths</div>', unsafe_allow_html=True)
    for s in analysis.get("strengths", []):
        st.markdown(f"✅ {s}")

    # ── Improvements ──
    st.markdown('<div class="sec-header">🔧 Improvements</div>', unsafe_allow_html=True)
    icons = {"HIGH": "🔴", "MEDIUM": "🟡", "LOW": "🟢"}
    for imp in analysis.get("improvements", []):
        p = imp.get("priority", "MEDIUM")
        st.markdown(
            f'<div class="imp-item">{icons.get(p, "⚪")} <strong>[{p}]</strong> '
            f'{imp.get("issue", "")} <br>'
            f'<span style="color:#C89AF0">→ {imp.get("suggestion", "")}</span></div>',
            unsafe_allow_html=True,
        )

    # ── Optimized CV ──
    st.markdown('<div class="sec-header">📝 ATS-Optimized CV</div>', unsafe_allow_html=True)
    with st.spinner("✍️ Generating optimized CV…"):
        opt_cv = gen_cv(resume_text, jd_text, analysis)

    st.text_area("ATS-Optimized CV (copy or download)", opt_cv, height=400)

    # ── PDF Export ──
def build_analysis_pdf(analysis, opt_cv):
    from fpdf import FPDF
    import io

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    # Ensure margins are defined so multi_cell has a boundary
    pdf.set_margins(15, 15, 15)
    
    # Helper to clean text for FPDF compatibility
    def format_line(text):
        # Encode to latin-1 to avoid non-standard character issues, 
        # replacing unknown chars with '?'
        return str(text).encode('latin-1', 'replace').decode('latin-1')

    # Convert the analysis dict to a readable string format
    content = f"Candidate: {analysis.get('candidate_name', 'N/A')}\n\n"
    content += f"Summary:\n{analysis.get('recommendation', 'N/A')}\n\n"
    content += f"Strengths:\n{analysis.get('strengths', 'N/A')}\n\n"
    content += f"Improvements:\n{analysis.get('improvements', 'N/A')}\n\n"
    content += "--- OPTIMIZED CV ---\n"
    content += opt_cv

    # Use multi_cell with an explicit width calculation to prevent overflow
    # 'w=0' means reach the right margin
    for line in content.split('\n'):
        clean_str = format_line(line)
        # Only process if line is not empty to avoid unnecessary breaks
        if clean_str.strip():
            # We use 5 as height. If this still fails, ensure font size 
            # is small enough for the margin space
            pdf.multi_cell(0, 5, clean_str)
        else:
            pdf.ln(5)

    return pdf.output(dest='S').encode('latin-1')

# ─────────────────────────────────────────────
# RENDER: BATCH COMPARISON DASHBOARD
# ─────────────────────────────────────────────
def render_batch_dashboard(batch_results: list, jd_title: str):
    """Render the sortable comparison table and batch export controls."""
    st.markdown('<div class="sec-header">👥 Candidate Comparison Dashboard</div>', unsafe_allow_html=True)

    if not batch_results:
        st.info("No batch results to display.")
        return

    sorted_results = sorted(
        batch_results,
        key=lambda x: x["analysis"].get("ats_score", 0),
        reverse=True
    )
    top_score = sorted_results[0]["analysis"].get("ats_score", 0) if sorted_results else 0

    # Build dataframe
    rows = []
    for rank, item in enumerate(sorted_results, 1):
        a = item["analysis"]
        ats = a.get("ats_score", 0)
        rows.append({
            "Rank": rank,
            "File": item["filename"],
            "Name": a.get("candidate_name", "Unknown") or "Unknown",
            "ATS Score": ats,
            "KW Match": a.get("keyword_match_score", 0),
            "Format": a.get("format_score", 0),
            "Readability": a.get("readability_score", 0),
            "Recommendation": a.get("overall_recommendation", ""),
            "Experience": a.get("years_experience", "") or "",
            "Top Pick": "⭐ YES" if ats == top_score else "",
        })

    df = pd.DataFrame(rows)

    # Display sortable dataframe
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ATS Score": st.column_config.ProgressColumn(
                "ATS Score", min_value=0, max_value=100, format="%d"
            ),
            "KW Match": st.column_config.ProgressColumn(
                "KW Match", min_value=0, max_value=100, format="%d"
            ),
            "Top Pick": st.column_config.TextColumn("Top Pick", width="small"),
        },
    )

    # Top pick highlight
    if sorted_results:
        top = sorted_results[0]
        ta = top["analysis"]
        st.markdown(
            f'<div class="hire-card" style="border-color:#7F3AA1;">'
            f'<span class="top-pick-badge">⭐ TOP PICK</span>'
            f'<h3 style="margin-top:10px">{ta.get("candidate_name", top["filename"])}</h3>'
            f'<p style="color:#A090C0">{ta.get("current_title", "")} &nbsp;|&nbsp; '
            f'ATS Score: <span style="color:{score_color(ta.get("ats_score",0))};font-weight:700">'
            f'{ta.get("ats_score", 0)}</span></p>'
            f'<p style="color:#C89AF0">{ta.get("summary", "")}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Expandable detail per candidate
    st.markdown('<div class="sec-header">📋 Individual Candidate Details</div>', unsafe_allow_html=True)
    for rank, item in enumerate(sorted_results, 1):
        a = item["analysis"]
        name = a.get("candidate_name", item["filename"]) or item["filename"]
        ats = a.get("ats_score", 0)
        badge = " ⭐ Top Pick" if rank == 1 else f" #{rank}"
        with st.expander(f"{badge} — {name} | ATS: {ats}"):
            col_a, col_b, col_c, col_d = st.columns(4)
            col_a.metric("ATS Score", ats)
            col_b.metric("KW Match", a.get("keyword_match_score", 0))
            col_c.metric("Format", a.get("format_score", 0))
            col_d.metric("Readability", a.get("readability_score", 0))

            st.markdown(f"**Recommendation:** {a.get('overall_recommendation', '')}")
            st.markdown(f"> {a.get('summary', '')}")

            skills_html = " ".join(f'<span class="tag">{s}</span>' for s in a.get("skills", []))
            st.markdown("**Skills:**")
            st.markdown(skills_html or "None", unsafe_allow_html=True)

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    # Batch Export
    st.markdown('<div class="sec-header">📦 Batch Export</div>', unsafe_allow_html=True)
    exp1, exp2 = st.columns(2)

    with exp1:
        csv_bytes = build_batch_csv(batch_results)
        st.download_button(
            "⬇️ Download Summary CSV (All Candidates)",
            data=csv_bytes,
            file_name=f"hireai_batch_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with exp2:
        with st.spinner("📑 Building Top-3 PDF Report…"):
            batch_pdf = build_batch_pdf(batch_results, jd_title)
        st.download_button(
            "⬇️ Download Top-3 Merged PDF Report",
            data=batch_pdf,
            file_name=f"hireai_top3_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown('<div class="brand-header">🔮 HireAi</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-sub">AI Resume Intelligence</div>', unsafe_allow_html=True)
        st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

        if not st.session_state.logged_in:
            auth_mode = st.radio("", ["Sign In", "Create Account"], horizontal=True)

            if auth_mode == "Sign In":
                email = st.text_input("Email", key="login_email", placeholder="you@example.com")
                pw = st.text_input("Password", type="password", key="login_pw")
                if st.button("Sign In", use_container_width=True):
                    if login_user(email, pw):
                        st.rerun()
                    else:
                        st.error("Invalid credentials or account not found.")
                st.caption("Don't have an account? Switch to **Create Account**.")

            else:
                name = st.text_input("Full Name", key="reg_name", placeholder="Jane Doe")
                email = st.text_input("Email", key="reg_email", placeholder="you@example.com")
                pw = st.text_input("Password", type="password", key="reg_pw")
                pw2 = st.text_input("Confirm Password", type="password", key="reg_pw2")
                if st.button("Create Account", use_container_width=True):
                    if pw != pw2:
                        st.error("Passwords do not match.")
                    elif not name or not email:
                        st.error("All fields are required.")
                    elif register_user(name, email, pw):
                        login_user(email, pw)
                        st.success(f"Welcome, {name}!")
                        st.rerun()
                    else:
                        st.error("Email already registered.")

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)
            st.info("Sign in to save analysis history, track job applications, and export PDF reports.")

        else:
            st.markdown(
                f'<div class="user-badge">'
                f'<div class="user-badge-name">👤 {st.session_state.user_name}</div>'
                f'<div class="user-badge-email">{st.session_state.user_email}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            st.markdown("**Navigation**")
            tabs = ["🔍 Analyzer", "📊 Dashboard", "💼 Saved Jobs"]
            for tab in tabs:
                if st.button(tab, use_container_width=True, key=f"nav_{tab}"):
                    st.session_state.current_tab = tab.split(" ", 1)[1]

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

            hist = st.session_state.analysis_history
            jobs = st.session_state.saved_jobs
            batch = st.session_state.batch_results
            col1, col2 = st.columns(2)
            col1.metric("Analyses", len(hist))
            col2.metric("Saved Jobs", len(jobs))
            if batch:
                st.metric("Batch Last Run", len(batch))
            if hist:
                avg = sum(h["ats"] for h in hist) / len(hist)
                st.metric("Avg ATS Score", f"{avg:.0f}")

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)
            st.markdown("**Model**")
            st.caption("LLaMA 3.3-70B via Groq Cloud")

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)
            if st.button("Sign Out", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.user_name = ""
                st.session_state.user_email = ""
                st.session_state.current_tab = "Analyzer"
                st.rerun()


# ─────────────────────────────────────────────
# TAB: ANALYZER
# ─────────────────────────────────────────────
def render_analyzer():
    st.markdown('<div class="brand-header">🔮 HireAi</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="brand-sub">ATS scoring · keyword gap analysis · CV optimization · PDF export · Batch processing</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    # ── Mode Toggle ──
    mode = st.radio(
        "**Processing Mode**",
        ["📄 Single Resume", "📦 Batch / ZIP"],
        horizontal=True,
        label_visibility="visible",
    )

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    # ── Job Description (shared for both modes) ──
    st.markdown('<div class="sec-header">📝 Job Description</div>', unsafe_allow_html=True)
    jd_col1, jd_col2 = st.columns([1, 2])
    with jd_col1:
        jd_title = st.text_input(
            "Job title (for tracking & reports)",
            placeholder="e.g. Senior ML Engineer",
            key="jd_title_input",
        )
    with jd_col2:
        jd_text = st.text_area(
            "Paste full job description",
            height=150,
            placeholder="Paste the job description here…",
            label_visibility="collapsed",
            key="jd_text_input",
        )

    if st.session_state.logged_in and jd_text.strip() and jd_title.strip():
        if st.button("💾 Save This Job", use_container_width=False):
            st.session_state.saved_jobs.append({
                "title": jd_title,
                "jd": jd_text,
                "saved_at": datetime.now().strftime("%b %d, %Y"),
            })
            st.success(f"Saved '{jd_title}' to Saved Jobs!")

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # SINGLE RESUME MODE
    # ══════════════════════════════════════════
    if mode == "📄 Single Resume":
        st.markdown('<div class="sec-header">📤 Upload Resume</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "PDF, DOCX, or TXT",
            type=["pdf", "docx", "txt"],
            label_visibility="collapsed",
            key="single_uploader",
        )

        run = st.button("🚀 Analyze Resume", use_container_width=True, key="single_run")

        if run:
            if not uploaded:
                st.error("Please upload a resume file.")
                return
            if not jd_text.strip():
                st.error("Please paste a job description.")
                return

            with st.spinner("🔍 Extracting text from resume…"):
                file_bytes = uploaded.read()
                resume_text = extract_text(file_bytes, uploaded.name)

            if len(resume_text.strip()) < 50:
                st.error("Could not extract text — is the PDF a scanned image? Try a text-based PDF or DOCX.")
                return

            with st.spinner("🤖 Analyzing with LLaMA 3.3-70B…"):
                analysis = analyze_resume(resume_text, jd_text)

            if not analysis:
                st.error("Analysis failed. Please check your GROQ_API_KEY in Streamlit secrets and retry.")
                return

            # Store in history and session
            if st.session_state.logged_in:
                st.session_state.analysis_history.append({
                    "date": datetime.now().strftime("%b %d, %Y %H:%M"),
                    "name": analysis.get("candidate_name", "Unknown"),
                    "ats": analysis.get("ats_score", 0),
                    "jd_title": jd_title or "Untitled",
                    "data": analysis,
                })

            st.session_state.last_single_analysis = analysis
            st.session_state.last_single_resume_text = resume_text
            st.session_state.last_single_jd_title = jd_title

        # Render results if available
        if st.session_state.last_single_analysis:
            render_single_results(
                st.session_state.last_single_analysis,
                st.session_state.last_single_resume_text,
                jd_text,
                st.session_state.last_single_jd_title,
            )

    # ══════════════════════════════════════════
    # BATCH / ZIP MODE
    # ══════════════════════════════════════════
    else:
        st.markdown('<div class="sec-header">📦 Upload Resumes (ZIP or Multiple Files)</div>', unsafe_allow_html=True)
        st.caption(
            "Upload a single ZIP archive containing PDF/DOCX/TXT resumes, "
            "or select multiple individual files at once."
        )

        batch_upload = st.file_uploader(
            "ZIP archive or multiple PDF/DOCX/TXT files",
            type=["zip", "pdf", "docx", "txt"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key="batch_uploader",
        )

        run_batch = st.button("🚀 Analyze All Resumes", use_container_width=True, key="batch_run")

        if run_batch:
            if not batch_upload:
                st.error("Please upload at least one file or a ZIP archive.")
                return
            if not jd_text.strip():
                st.error("Please paste a job description.")
                return

            # Collect all files to analyze
            files_to_analyze = []  # list of {filename, text}

            for uploaded_file in batch_upload:
                raw = uploaded_file.read()
                fname = uploaded_file.name.lower()

                if fname.endswith(".zip"):
                    extracted = extract_from_zip(raw)
                    files_to_analyze.extend(extracted)
                else:
                    text = extract_text(raw, uploaded_file.name)
                    if len(text.strip()) >= 50:
                        files_to_analyze.append({"filename": uploaded_file.name, "text": text})
                    else:
                        st.warning(f"⚠️ Skipped {uploaded_file.name} — could not extract sufficient text.")

            if not files_to_analyze:
                st.error("No valid resume text found in uploaded files. Check that files are text-based PDFs or DOCX.")
                return

            st.info(f"📂 Found **{len(files_to_analyze)}** resume(s) to analyze. Starting batch processing…")

            progress_bar = st.progress(0, text="Initializing…")
            batch_results = []
            errors = 0

            for idx, file_item in enumerate(files_to_analyze):
                progress_fraction = idx / len(files_to_analyze)
                progress_bar.progress(
                    progress_fraction,
                    text=f"Analyzing {idx + 1}/{len(files_to_analyze)}: {file_item['filename']}…"
                )

                analysis = analyze_resume(file_item["text"], jd_text)

                if analysis:
                    batch_results.append({
                        "filename": file_item["filename"],
                        "analysis": analysis,
                    })
                    # Save each to history
                    if st.session_state.logged_in:
                        st.session_state.analysis_history.append({
                            "date": datetime.now().strftime("%b %d, %Y %H:%M"),
                            "name": analysis.get("candidate_name", file_item["filename"]),
                            "ats": analysis.get("ats_score", 0),
                            "jd_title": jd_title or "Batch Analysis",
                            "data": analysis,
                        })
                else:
                    errors += 1
                    st.warning(f"⚠️ Analysis failed for {file_item['filename']} — skipped.")

            progress_bar.progress(1.0, text="✅ Batch analysis complete!")

            st.session_state.batch_results = batch_results

            if errors:
                st.warning(f"{errors} file(s) failed to analyze. Results below show successful analyses only.")

        # Render batch dashboard if results exist
        if st.session_state.batch_results:
            render_batch_dashboard(st.session_state.batch_results, jd_title)


# ─────────────────────────────────────────────
# TAB: DASHBOARD
# ─────────────────────────────────────────────
def render_dashboard():
    st.markdown('<div class="brand-header">📊 Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-sub">Track your ATS performance over time</div>', unsafe_allow_html=True)
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    hist = st.session_state.analysis_history

    if not hist:
        st.info("No analyses yet. Run your first resume analysis from the **Analyzer** tab.")
        return

    avg_ats = sum(h["ats"] for h in hist) / len(hist)
    best = max(hist, key=lambda h: h["ats"])

    p1, p2, p3 = st.columns(3)
    p1.markdown(
        f'<div class="stat-pill"><div class="stat-pill-num">{avg_ats:.0f}</div>'
        f'<div class="stat-pill-lbl">Avg ATS Score</div></div>',
        unsafe_allow_html=True,
    )
    p2.markdown(
        f'<div class="stat-pill"><div class="stat-pill-num">{best["ats"]}</div>'
        f'<div class="stat-pill-lbl">Best Score</div></div>',
        unsafe_allow_html=True,
    )
    p3.markdown(
        f'<div class="stat-pill"><div class="stat-pill-num">{len(hist)}</div>'
        f'<div class="stat-pill-lbl">Total Analyses</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="sec-header">📈 ATS Score Trend</div>', unsafe_allow_html=True)
    chart_data = pd.DataFrame({
        "Analysis #": list(range(1, len(hist) + 1)),
        "ATS Score": [h["ats"] for h in hist],
    }).set_index("Analysis #")
    st.line_chart(chart_data["ATS Score"])

    st.markdown('<div class="sec-header">🗂 Analysis History</div>', unsafe_allow_html=True)
    for h in reversed(hist):
        col_a, col_b, col_c, col_d = st.columns([3, 2, 1, 1])
        col_a.markdown(f"**{h['name']}**")
        col_b.markdown(f"_{h['jd_title']}_")
        col_c.markdown(
            f'<span style="color:{score_color(h["ats"])};font-weight:700">{h["ats"]}</span>',
            unsafe_allow_html=True,
        )
        col_d.markdown(f"<small>{h['date']}</small>", unsafe_allow_html=True)
        st.divider()

    if len(hist) >= 3:
        latest = hist[-1]
        st.markdown('<div class="sec-header">📊 Score Breakdown (Latest Analysis)</div>', unsafe_allow_html=True)
        latest_data = latest["data"]
        score_df = pd.DataFrame({
            "Category": ["ATS Score", "Keyword Match", "Format", "Readability"],
            "Score": [
                latest_data.get("ats_score", 0),
                latest_data.get("keyword_match_score", 0),
                latest_data.get("format_score", 0),
                latest_data.get("readability_score", 0),
            ],
        }).set_index("Category")
        st.bar_chart(score_df)


# ─────────────────────────────────────────────
# TAB: SAVED JOBS
# ─────────────────────────────────────────────
def render_saved_jobs():
    st.markdown('<div class="brand-header">💼 Saved Jobs</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="brand-sub">Track job descriptions and see how your resume ranks</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    jobs = st.session_state.saved_jobs
    if not jobs:
        st.info("No saved jobs yet. Paste a job description in the **Analyzer** tab and click **Save This Job**.")
        return

    st.markdown(f"**{len(jobs)} saved position{'s' if len(jobs) != 1 else ''}**")
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    for i, job in enumerate(jobs):
        with st.expander(f"🏢 {job['title']} — saved {job['saved_at']}"):
            st.markdown(f'<div class="job-card-meta">Saved on {job["saved_at"]}</div>', unsafe_allow_html=True)
            st.text_area("Job Description", job["jd"], height=200, key=f"jd_view_{i}", disabled=True)

            related = [
                h for h in st.session_state.analysis_history
                if h["jd_title"] == job["title"]
            ]
            if related:
                st.markdown(f"**{len(related)} resume(s) analyzed against this role:**")
                for r in related:
                    c1, c2, c3 = st.columns([3, 1, 2])
                    c1.markdown(f"📄 {r['name']}")
                    c2.markdown(
                        f'<span style="color:{score_color(r["ats"])};font-weight:700">{r["ats"]} ATS</span>',
                        unsafe_allow_html=True,
                    )
                    c3.markdown(f"<small>{r['date']}</small>", unsafe_allow_html=True)
            else:
                st.caption("No resumes analyzed against this role yet.")

            col_del = st.columns([1, 3])[0]
            if col_del.button("🗑 Remove Job", key=f"del_job_{i}"):
                st.session_state.saved_jobs.pop(i)
                st.rerun()


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
render_sidebar()

if not st.session_state.logged_in:
    st.markdown('<div class="brand-header">🔮 HireAi</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="brand-sub">AI-powered resume intelligence. ATS scoring, keyword gap analysis, '
        'batch processing, and optimized CV generation.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    f1.markdown(
        '<div class="hire-card"><h3>🎯 ATS Scoring</h3>'
        '<p style="color:#A090C0">Get a detailed breakdown of resume ATS compatibility, '
        'format quality, and keyword alignment.</p></div>',
        unsafe_allow_html=True,
    )
    f2.markdown(
        '<div class="hire-card"><h3>🔑 Keyword Gap</h3>'
        '<p style="color:#A090C0">Instantly identify missing keywords from any job description '
        'and understand what recruiters are scanning for.</p></div>',
        unsafe_allow_html=True,
    )
    f3.markdown(
        '<div class="hire-card"><h3>📦 Batch Mode</h3>'
        '<p style="color:#A090C0">Upload a ZIP of resumes and rank all candidates side-by-side '
        'in seconds. Export CSV or Top-3 PDF report.</p></div>',
        unsafe_allow_html=True,
    )
    f4.markdown(
        '<div class="hire-card"><h3>📄 PDF Export</h3>'
        '<p style="color:#A090C0">Download polished PDF analysis reports and ATS-optimized '
        'CV templates ready to submit.</p></div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)
    st.info("👈 **Create a free account or sign in** from the sidebar to get started.")

else:
    current = st.session_state.get("current_tab", "Analyzer")
    if current == "Analyzer":
        render_analyzer()
    elif current == "Dashboard":
        render_dashboard()
    elif current == "Saved Jobs":
        render_saved_jobs()
    else:
        render_analyzer()
