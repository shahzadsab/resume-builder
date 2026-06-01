# ATS Resume & Cover Letter Builder 📄

A **production-ready, zero-cost** Streamlit web app that generates ATS-optimised
resumes and tailored cover letters from your existing resume and a job description.

---

## Features

| Feature | Without HF Token | With Free HF Token |
|---|---|---|
| Resume parsing (PDF/DOCX) | ✅ | ✅ |
| Keyword extraction from JD | ✅ (regex) | ✅ (spaCy + regex) |
| ATS score before/after | ✅ | ✅ |
| .docx resume download | ✅ | ✅ |
| Company Wikipedia profile | ✅ | ✅ |
| Salary estimation | ✅ (rule-based) | ✅ (rule-based) |
| AI resume rewriting | ❌ (passthrough) | ✅ (Flan-T5) |
| AI cover letter | ❌ (template) | ✅ (AI-generated) |
| Culture & benefits info | ❌ | ✅ (Flan-T5) |

---

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Download spaCy model (optional but recommended)
python -m spacy download en_core_web_sm

# 3. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

---

## Getting a Free Hugging Face Token

1. Create a free account at https://huggingface.co
2. Go to https://huggingface.co/settings/tokens
3. Click "New token" → select "Read" role → copy the token
4. Paste into the sidebar when running the app

Free tier includes ~30,000 tokens/month — more than enough for daily use.

---

## Deploy to Streamlit Cloud (Free)

1. Push this repo to GitHub
2. Go to https://share.streamlit.io
3. Connect your repo, set `app.py` as the entry point
4. Deploy — Streamlit Cloud is free for public apps

**Note:** On Streamlit Cloud, set your HF token as a Secret:
`Settings → Secrets → HF_TOKEN = "hf_xxxxx"`

Then read it in app.py:
```python
import streamlit as st
hf_token = st.secrets.get("HF_TOKEN", "")
```

---

## Architecture

```
app.py
├── parse_pdf() / parse_docx()        – pdfplumber / python-docx
├── extract_resume_sections()          – regex heuristic parser
├── extract_keywords()                 – spaCy or regex fallback
├── score_resume()                     – keyword match score
├── analyse_company()                  – Wikipedia + HF Inference API
├── optimise_resume_text()             – keyword injection + HF rewrite
├── build_docx()                       – ATS-safe .docx output
├── generate_cover_letter()            – HF or template
└── build_cover_letter_docx()          – .docx cover letter
```

---

## Cost Breakdown

| Service | Cost |
|---|---|
| Streamlit (hosting) | Free |
| Wikipedia API | Free |
| spaCy en_core_web_sm | Free |
| Hugging Face Inference API | Free (30k tokens/month) |
| pdfplumber / python-docx | Free (open source) |

**Total: $0**
