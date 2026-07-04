# 🪄 HireAi — AI Resume Intelligence Platform

**HireAi** is a production-ready SaaS application that gives job seekers a deep, AI-powered analysis of their resume against any job description. It surfaces ATS compatibility scores, identifies keyword gaps, generates an optimized CV, and exports a polished PDF report — all powered by LLaMA 3.3-70B via Groq Cloud.

---

## ✨ Features

| Feature | Description |
|---|---|
| **ATS Scoring** | Four-metric score: Overall ATS, Keyword Match, Format Quality, Readability |
| **Keyword Gap Analysis** | Matched vs. missing keywords pulled directly from the JD |
| **AI CV Optimization** | LLM-rewritten resume that naturally weaves in missing keywords |
| **PDF Export** | Branded, two-page PDF report (analysis + optimized CV) |
| **User Accounts** | Session-based login/registration with persistent history |
| **Dashboard** | ATS score trend chart and full analysis history |
| **Saved Jobs** | Save job descriptions and track which resumes rank against them |

---

## 🚀 Deploy to Streamlit Cloud (5 minutes)

### Step 1 — Fork / Push to GitHub

Push this repository (containing `app.py`, `requirements.txt`, and `README.md`) to a GitHub repository.

```
your-repo/
├── app.py
├── requirements.txt
└── README.md
```

### Step 2 — Create a Streamlit Cloud App

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **"New app"**.
3. Select your repository, branch (`main`), and set **Main file path** to `app.py`.
4. Click **"Deploy"**.

### Step 3 — Add Your Groq API Key (Secrets)

This is the most important step. HireAi reads the key from Streamlit's secure secrets store — **never hardcode it**.

1. In your Streamlit Cloud dashboard, open your app and click the **⋮ menu → Settings → Secrets**.
2. Add the following:

```toml
GROQ_API_KEY = "gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

3. Click **"Save"**. The app will automatically restart with the key loaded.

> **Get a free Groq API key:** [console.groq.com](https://console.groq.com)

---

## 💻 Local Development

```bash
# 1. Clone the repo
git clone https://github.com/your-username/hireai.git
cd hireai

# 2. Create and activate a virtual environment
python -m venv venv
source venv/bin/activate      # macOS/Linux
venv\Scripts\activate         # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set your API key
#    Create a file at .streamlit/secrets.toml with:
#    GROQ_API_KEY = "gsk_xxxxxxxxxxxxxxxxxxxx"

mkdir -p .streamlit
echo 'GROQ_API_KEY = "your_key_here"' > .streamlit/secrets.toml

# 5. Run the app
streamlit run app.py
```

---

## 🗂 Project Structure

```
hireai/
├── app.py              # Main application (all modules inline)
├── requirements.txt    # Production dependencies
└── README.md           # This file
```

### Key Modules Inside `app.py`

| Function | Purpose |
|---|---|
| `get_groq_client()` | Cached Groq client initialized from `st.secrets` |
| `extract_text()` | Cached PDF/DOCX/TXT text extraction via `pdfplumber` / `python-docx` |
| `analyze_resume()` | Cached LLM analysis returning structured JSON |
| `gen_cv()` | LLM-powered ATS-optimized CV rewriter |
| `build_analysis_pdf()` | `fpdf2`-powered branded PDF export (2 pages) |
| `render_analyzer()` | Main analysis UI tab |
| `render_dashboard()` | Score history + trend chart tab |
| `render_saved_jobs()` | Job description tracker tab |
| `render_sidebar()` | Auth + navigation sidebar |

---

## 🔐 Security Notes

- **API keys** are read exclusively from `st.secrets["GROQ_API_KEY"]` — never hardcoded.
- **User passwords** are SHA-256 hashed before being stored in session state.
- **Session state** is scoped to each browser session; no data persists between sessions without an external database.

> For a production deployment with persistent storage, replace `st.session_state` history/user stores with a database backend (e.g., Supabase, Firebase, or PostgreSQL via `psycopg2`).

---

## 📦 Dependencies

| Package | Version | Purpose |
|---|---|---|
| `streamlit` | ≥1.35 | Web app framework |
| `groq` | ≥0.9 | LLaMA 3.3-70B via Groq Cloud |
| `pdfplumber` | ≥0.11 | PDF text extraction |
| `python-docx` | ≥1.1 | DOCX text extraction |
| `fpdf2` | ≥2.7.9 | PDF generation / export |
| `pandas` | ≥2.2 | Data handling + charts |

---

## 🎨 Theme

HireAi uses a custom dark purple palette:

| Role | Hex |
|---|---|
| Background | `#0C0516` |
| Primary / Headers | `#7F3AA1` |
| Cards (gradient from) | `#0F083B` |
| Cards (gradient to) | `#5416B5` |
| Body Text | `#E8EAF0` |
| Accent Text | `#C89AF0` |

---

## 📄 License

MIT — free to use, modify, and deploy.

---

*Built with ❤️ using Groq · LLaMA 3.3-70B · Streamlit · fpdf2*
