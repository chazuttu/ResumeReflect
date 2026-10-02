# ⚡ ResumeReflect – AI Resume Tailor

**Match any job in 30 seconds.** Upload your resume (PDF), paste a job description, and get a one-page resume written in the employer's own language, plus an ATS report.

### 👉 [Try it free: resumereflect.streamlit.app](https://resumereflect.streamlit.app/)

---

## What you get (free)
- **Tailored one-page resume**: PDF and editable Word file
- **Your links stay clickable** (LinkedIn, GitHub, projects, portfolio)
- **ATS report**: resume optimisation score, skill match with the job, and a list of what's missing
- **Honest by design**: it never adds skills you don't have. Missing skills are shown so you know what to learn.

## How it works
1. Upload your resume (PDF)
2. Paste the job description
3. Download your tailored resume and ATS report

## Optional: Personal resume review + career counselling
One-to-one resume review and job-search guidance by email, available from inside the app.

## Privacy
- Your resume and job description are sent to an AI provider ([Groq](https://groq.com)) to generate the result.
- The app does not keep a database of resumes.
- If you book a counselling session, your email and tailored resume are shared with the owner so a review can be done.

## Built with
Python · Streamlit · Groq API · ReportLab · python-docx · pdfplumber

## Run it yourself
```bash
pip install -r requirements.txt
streamlit run app.py
```
Add your keys in `.streamlit/secrets.toml` (never commit this file):
```toml
GROQ_API_KEY = "your-key"
```

Made in India 🇮🇳 · © 2026 ResumeReflect. All rights reserved.
