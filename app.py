"""
HireAi — AI-Powered Resume Intelligence Platform
Production-ready Streamlit SaaS app for Streamlit Cloud deployment.
"""

import streamlit as st
import json
import re
import io
import os
import hashlib
from datetime import datetime

import pdfplumber
import docx
import pandas as pd
from groq import Groq
from fpdf import FPDF
import unicodedata

    

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

/* Scrollbar */
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: #0C0516; }
::-webkit-scrollbar-thumb { background: #5416B5; border-radius: 3px; }

/* Main area */
.main { background: #0C0516; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0F083B 0%, #0C0516 100%);
    border-right: 1px solid #2A1A4A;
}

/* Headings */
h1, h2, h3 { color: #7F3AA1; font-weight: 700; }

/* Cards */
.hire-card {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 16px;
    padding: 24px;
    margin: 10px 0;
}

/* Score card */
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

/* Tags */
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

/* Improvement items */
.imp-item {
    background: #0F083B;
    border-left: 3px solid #7F3AA1;
    padding: 12px 16px;
    margin: 8px 0;
    border-radius: 0 10px 10px 0;
    font-size: .9rem;
    line-height: 1.5;
}

/* Section headers */
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

/* Brand header */
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

/* Buttons */
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

/* Text area */
.stTextArea textarea {
    background: #0F083B;
    color: #E8EAF0;
    border: 1px solid #3A1A6A;
    border-radius: 10px;
    font-family: 'Space Grotesk', sans-serif;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #0F083B;
    border: 1px dashed #5416B5;
    border-radius: 12px;
    padding: 8px;
}

/* Metrics */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 12px;
    padding: 16px;
}

/* Tabs */
[data-testid="stTab"] { color: #A090C0; }

/* Info / success / error boxes */
.stAlert { border-radius: 10px; }

/* Sidebar user badge */
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

/* Dashboard stat pill */
.stat-pill {
    background: linear-gradient(135deg, #0F083B, #5416B5 250%);
    border: 1px solid #5416B5;
    border-radius: 12px;
    padding: 16px;
    text-align: center;
}
.stat-pill-num { font-size: 2rem; font-weight: 700; color: #C89AF0; }
.stat-pill-lbl { font-size: .75rem; color: #7060A0; text-transform: uppercase; letter-spacing: 1px; font-family: 'DM Mono', monospace; }

/* History table row */
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

/* Saved job card */
.job-card {
    background: linear-gradient(135deg, #0F083B, #1A0F4A);
    border: 1px solid #3A1A6A;
    border-radius: 14px;
    padding: 20px;
    margin: 10px 0;
}
.job-card-title { font-weight: 700; color: #C89AF0; font-size: 1rem; }
.job-card-meta { color: #7060A0; font-size: .8rem; margin-top: 4px; font-family: 'DM Mono', monospace; }

/* divider */
.purple-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #5416B5, transparent);
    margin: 24px 0;
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
        "analysis_history": [],   # list of dicts: {date, name, ats, jd_title, data}
        "saved_jobs": [],          # list of dicts: {title, jd, saved_at}
        "current_tab": "Analyzer",
        "users_db": {},            # mock: {email: {pw_hash, name}}
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
@st.cache_data(show_spinner=False)
def extract_text(file_bytes: bytes, file_name: str) -> str:
    buf = io.BytesIO(file_bytes)
    if file_name.endswith(".pdf"):
        with pdfplumber.open(buf) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    if file_name.endswith(".docx"):
        d = docx.Document(buf)
        return "\n".join(p.text for p in d.paragraphs)
    return file_bytes.decode("utf-8", errors="ignore")


def ask_groq(sys_p: str, usr_p: str, temp: float = 0.3) -> str:
    client = get_groq_client()
    r = client.chat.completions.create(
        model=MODEL,
        temperature=temp,
        max_tokens=4096,
        messages=[{"role": "system", "content": sys_p},
                  {"role": "user", "content": usr_p}],
    )
    return r.choices[0].message.content.strip()


def parse_json(text: str):
    m = re.search(r"```json\s*([\s\S]+?)```", text)
    if m:
        text = m.group(1)
    else:
        s = text.find("{") if "{" in text else text.find("[")
        if s != -1:
            text = text[s:]
    try:
        return json.loads(text)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def analyze_resume(resume: str, jd: str) -> dict:
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
    if data is None:
        try:
            data = json.loads(raw)
        except Exception:
            data = {}
    return data


def gen_cv(resume: str, jd: str, analysis: dict) -> str:
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
    return ask_groq(sys_p, usr_p, temp=0.5)


def score_color(s: int) -> str:
    if s >= 75:
        return "#A0F080"
    if s >= 50:
        return "#F0D070"
    return "#F08090"


# ─────────────────────────────────────────────
# PDF EXPORT
# ─────────────────────────────────────────────
def build_analysis_pdf(analysis: dict, opt_cv: str) -> bytes:
    """Generate a polished PDF report from analysis data."""
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
    pdf.cell(0, 14, "HireAi - Resume Analysis Report", ln=True)

    pdf.set_xy(10, 22)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(112, 96, 160)
    pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}", ln=True)

    pdf.ln(10)

    # Candidate Info block
    pdf.set_fill_color(20, 12, 60)
    pdf.set_draw_color(84, 22, 181)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, "Candidate Profile", ln=True)
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
        val = analysis.get(key, "—") or "—"
        pdf.set_text_color(112, 96, 160)
        pdf.cell(55, 7, f"  {label}:", ln=False)
        pdf.set_text_color(232, 234, 240)
        pdf.cell(0, 7, str(val), ln=True)

    pdf.ln(6)

    # Scores
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, "ATS Score Summary", ln=True)
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
        pdf.cell(90, 8, f"  {label}", ln=False)
        if val >= 75:
            pdf.set_text_color(100, 220, 80)
        elif val >= 50:
            pdf.set_text_color(220, 200, 70)
        else:
            pdf.set_text_color(220, 80, 100)
        pdf.cell(0, 8, f"{val} / 100", ln=True)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(160, 200, 100)
    rec = analysis.get("overall_recommendation", "")
    pdf.cell(0, 8, f"  Recommendation: {rec}", ln=True)

    pdf.ln(3)
    pdf.set_font("Helvetica", "I", 10)
    pdf.set_text_color(180, 170, 210)
    summary = analysis.get("summary", "")
    pdf.multi_cell(0, 6, f"  {summary}")

    pdf.ln(6)

    # Keywords
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, "Keyword Analysis", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(100, 220, 80)
    pdf.cell(0, 7, "  Matched Keywords:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    matched = ", ".join(analysis.get("matched_keywords", []))
    pdf.multi_cell(0, 6, f"  {matched or 'None identified'}")

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(220, 80, 100)
    pdf.cell(0, 7, "  Missing Keywords:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    missing = ", ".join(analysis.get("missing_keywords", []))
    pdf.multi_cell(0, 6, f"  {missing or 'None identified'}")

    pdf.ln(4)

    # Strengths
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, "Strengths", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(232, 234, 240)
    for s in analysis.get("strengths", []):
        pdf.cell(0, 7, f"  ✓  {s}", ln=True)

    pdf.ln(4)

    # Improvements
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(200, 154, 240)
    pdf.cell(0, 8, "Improvement Recommendations", ln=True)
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
        pdf.cell(20, 7, f"  {icons.get(p, '')}", ln=False)
        pdf.set_text_color(232, 234, 240)
        issue = imp.get("issue", "")
        suggestion = imp.get("suggestion", "")
        pdf.multi_cell(0, 7, f"{issue}  →  {suggestion}")
        pdf.ln(1)

    # ── Page 2: Optimized CV ──
    if opt_cv:
        pdf.add_page()
        pdf.set_fill_color(15, 8, 59)
        pdf.rect(0, 0, 210, 30, "F")
        pdf.set_xy(10, 8)
        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(200, 154, 240)
        pdf.cell(0, 14, "HireAi — ATS-Optimized CV", ln=True)

        pdf.set_xy(10, 22)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(112, 96, 160)
        pdf.cell(0, 6, "AI-generated, keyword-optimized resume template", ln=True)

        pdf.ln(8)
        pdf.set_font("Courier", "", 9)
        pdf.set_text_color(220, 214, 240)
        for line in opt_cv.split("\n"):
            pdf.multi_cell(0, 5, line)

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)
    


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown('<div class="brand-header">🔮 HireAi</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-sub">AI Resume Intelligence</div>', unsafe_allow_html=True)
        st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

        # ── Auth ──
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
            # User badge
            st.markdown(
                f'<div class="user-badge">'
                f'<div class="user-badge-name">👤 {st.session_state.user_name}</div>'
                f'<div class="user-badge-email">{st.session_state.user_email}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            # Navigation
            st.markdown("**Navigation**")
            tabs = ["🔍 Analyzer", "📊 Dashboard", "💼 Saved Jobs"]
            for tab in tabs:
                if st.button(tab, use_container_width=True, key=f"nav_{tab}"):
                    st.session_state.current_tab = tab.split(" ", 1)[1]

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

            # Stats
            hist = st.session_state.analysis_history
            jobs = st.session_state.saved_jobs
            col1, col2 = st.columns(2)
            col1.metric("Analyses", len(hist))
            col2.metric("Saved Jobs", len(jobs))

            if hist:
                avg = sum(h["ats"] for h in hist) / len(hist)
                st.metric("Avg ATS Score", f"{avg:.0f}")

            st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

            # Model info
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
        '<div class="brand-sub">ATS scoring · keyword gap analysis · CV optimization · PDF export</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown('<div class="sec-header">📤 Upload Resume</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("PDF, DOCX, or TXT", type=["pdf", "docx", "txt"], label_visibility="collapsed")

    with c2:
        st.markdown('<div class="sec-header">📝 Job Description</div>', unsafe_allow_html=True)
        jd_title = st.text_input("Job title (optional, for tracking)", placeholder="e.g. Senior ML Engineer")
        jd_text = st.text_area("Paste full job description", height=180, placeholder="Paste the job description here…", label_visibility="collapsed")

    # Save Job button
    if st.session_state.logged_in and jd_text.strip() and jd_title.strip():
        if st.button("💾 Save This Job", use_container_width=False):
            st.session_state.saved_jobs.append({
                "title": jd_title,
                "jd": jd_text,
                "saved_at": datetime.now().strftime("%b %d, %Y"),
            })
            st.success(f"Saved '{jd_title}' to Saved Jobs!")

    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)
    run = st.button("🚀 Analyze Resume", use_container_width=True)

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
            st.error("Analysis failed. Check your GROQ_API_KEY in Streamlit secrets.")
            return

        # ── Save to history ──
        if st.session_state.logged_in:
            st.session_state.analysis_history.append({
                "date": datetime.now().strftime("%b %d, %Y %H:%M"),
                "name": analysis.get("candidate_name", "Unknown"),
                "ats": analysis.get("ats_score", 0),
                "jd_title": jd_title or "Untitled",
                "data": analysis,
            })

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
                f'<div class="imp-item">{icons.get(p,"⚪")} <strong>[{p}]</strong> '
                f'{imp.get("issue","")} <br>'
                f'<span style="color:#C89AF0">→ {imp.get("suggestion","")}</span></div>',
                unsafe_allow_html=True,
            )

        # ── Optimized CV ──
        st.markdown('<div class="sec-header">📝 ATS-Optimized CV</div>', unsafe_allow_html=True)
        with st.spinner("✍️ Generating optimized CV…"):
            opt_cv = gen_cv(resume_text, jd_text, analysis)

        st.text_area("ATS-Optimized CV (copy or download)", opt_cv, height=400)

        # ── PDF Export ──
        st.markdown('<div class="sec-header">📄 Export</div>', unsafe_allow_html=True)
        with st.spinner("📑 Building PDF report…"):
            pdf_bytes = build_analysis_pdf(analysis, opt_cv)

        fname = analysis.get("candidate_name", "candidate").replace(" ", "_")
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "⬇️ Download PDF Report",
            data=pdf_bytes,
            file_name=f"hireai_report_{fname}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        dl2.download_button(
            "⬇️ Download Optimized CV (.txt)",
            data=opt_cv.encode("utf-8"),
            file_name=f"hireai_cv_{fname}.txt",
            mime="text/plain",
            use_container_width=True,
        )

        st.success("✅ Analysis complete! Download your PDF report or copy the optimized CV above.")


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

    # ── Summary stats ──
    avg_ats = sum(h["ats"] for h in hist) / len(hist)
    best = max(hist, key=lambda h: h["ats"])
    latest = hist[-1]

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

    # ── ATS trend chart ──
    st.markdown('<div class="sec-header">📈 ATS Score Trend</div>', unsafe_allow_html=True)
    chart_data = pd.DataFrame({
        "Analysis #": list(range(1, len(hist) + 1)),
        "ATS Score": [h["ats"] for h in hist],
        "Job Title": [h["jd_title"] for h in hist],
    }).set_index("Analysis #")
    st.line_chart(chart_data["ATS Score"])

    # ── History list ──
    st.markdown('<div class="sec-header">🗂 Analysis History</div>', unsafe_allow_html=True)
    for i, h in enumerate(reversed(hist), 1):
        col_a, col_b, col_c, col_d = st.columns([3, 2, 1, 1])
        col_a.markdown(f"**{h['name']}**")
        col_b.markdown(f"_{h['jd_title']}_")
        col_c.markdown(
            f'<span style="color:{score_color(h["ats"])};font-weight:700">{h["ats"]}</span>',
            unsafe_allow_html=True,
        )
        col_d.markdown(f"<small>{h['date']}</small>", unsafe_allow_html=True)
        st.divider()

    # ── Score distribution ──
    if len(hist) >= 3:
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

            # Show how analyses ranked against this JD
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
                st.caption("No resumes analyzed against this role yet. Go to Analyzer and use this job description.")

            col_del = st.columns([1, 3])[0]
            if col_del.button("🗑 Remove Job", key=f"del_job_{i}"):
                st.session_state.saved_jobs.pop(i)
                st.rerun()


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
render_sidebar()

if not st.session_state.logged_in:
    # Landing page for unauthenticated users
    st.markdown('<div class="brand-header">🔮 HireAi</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="brand-sub">AI-powered resume intelligence. ATS scoring, keyword gap analysis, and optimized CV generation.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="purple-divider"></div>', unsafe_allow_html=True)

    f1, f2, f3 = st.columns(3)
    f1.markdown(
        '<div class="hire-card"><h3>🎯 ATS Scoring</h3><p style="color:#A090C0">Get a detailed breakdown of your resume\'s ATS compatibility, format quality, and keyword alignment.</p></div>',
        unsafe_allow_html=True,
    )
    f2.markdown(
        '<div class="hire-card"><h3>🔑 Keyword Gap</h3><p style="color:#A090C0">Instantly identify missing keywords from any job description and understand what recruiters are scanning for.</p></div>',
        unsafe_allow_html=True,
    )
    f3.markdown(
        '<div class="hire-card"><h3>📄 PDF Export</h3><p style="color:#A090C0">Download a polished PDF analysis report and an ATS-optimized CV template ready to submit.</p></div>',
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
