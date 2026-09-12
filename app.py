import io
import json
import os

import groq
import pdfplumber
import streamlit as st
from docx import Document as DocxDocument
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable
)

# ─────────────────────────────────────────────
# PAGE CONFIG + STYLE
# ─────────────────────────────────────────────
st.set_page_config(page_title="JD → Resume", page_icon="🎯", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600&display=swap');

#MainMenu, footer, header {visibility: hidden;}
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}

.stApp {
    background: radial-gradient(circle at 15% 0%, #fff8ea 0%, transparent 45%),
                radial-gradient(circle at 85% 15%, #eef6f2 0%, transparent 40%),
                #f7f6f2;
}
.block-container {padding-top: 3rem; padding-bottom: 3rem; max-width: 620px;}

.badge-row {display: flex; gap: 8px; margin-bottom: 1.1rem;}
.badge-chip {
    background: rgba(255,255,255,0.7); border: 1px solid #e9e6dc; border-radius: 100px;
    padding: 4px 12px; font-size: 0.72rem; color: #6b6f5e; font-weight: 500;
}

.hero-title {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 2.3rem; font-weight: 800; color: #161811;
    margin-bottom: 0.4rem; letter-spacing: -0.03em; line-height: 1.1;
}
.hero-title span {color: #C99A3E;}
.hero-tag {color: #6b6f5e; font-size: 1rem; margin-bottom: 2rem; line-height: 1.55; max-width: 480px;}

.form-card {
    background: #ffffff; border-radius: 20px; padding: 2rem 2rem 1.6rem;
    border: 1px solid #edebe0; box-shadow: 0 12px 32px rgba(30,28,14,0.06);
    margin-bottom: 1.5rem;
}

.step-label {display: flex; align-items: center; gap: 8px; margin: 0 0 0.6rem;}
.step-num {
    width: 20px; height: 20px; border-radius: 50%; background: #161811; color: #fff;
    font-size: 0.68rem; font-weight: 600; display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
.step-text {font-family: 'Plus Jakarta Sans', sans-serif; font-size: 0.88rem; font-weight: 600; color: #161811;}

textarea {
    font-size: 0.9rem !important;
    border-radius: 12px !important;
    border: 1.5px solid #e5e3d6 !important;
    background: #fbfaf5 !important;
}
textarea:focus {border-color: #C99A3E !important; box-shadow: 0 0 0 3px rgba(201,154,62,0.15) !important;}

[data-testid="stFileUploaderDropzone"] {
    background: #fbfaf5 !important;
    border: 1.5px dashed #d8d5c4 !important;
    border-radius: 12px !important;
}

div[data-testid="stTextArea"] {margin-bottom: 1.5rem;}

div.stButton > button {
    background: linear-gradient(180deg, #E3B65F 0%, #C99A3E 100%);
    color: #1a1706; font-weight: 700; font-size: 0.95rem;
    border: none; border-radius: 12px; padding: 0.8rem 1.2rem; width: 100%;
    margin-top: 0.4rem; box-shadow: 0 4px 14px rgba(201,154,62,0.4);
    transition: transform 0.12s ease, box-shadow 0.12s ease;
}
div.stButton > button:hover {transform: translateY(-2px); color: #1a1706; box-shadow: 0 8px 20px rgba(201,154,62,0.45);}
div.stButton > button:active {transform: translateY(0px);}

.result-card {
    background: linear-gradient(160deg, #1b1d15 0%, #14161B 100%);
    border-radius: 20px; padding: 1.8rem;
    text-align: center; margin: 1.8rem 0 1rem;
    box-shadow: 0 14px 30px rgba(20,22,27,0.25);
}
.score-num {font-family: 'Plus Jakarta Sans', sans-serif; font-size: 2.8rem; font-weight: 800; color: #F5F1E7;}
.score-label {color: #9a9d8f; font-size: 0.85rem; margin-top: 2px; letter-spacing: 0.01em;}
.score-bar-bg {background: #2a2c24; border-radius: 20px; height: 8px; margin-top: 16px; overflow: hidden;}
.score-bar-fill {height: 100%; border-radius: 20px;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="badge-row">
    <div class="badge-chip">AI-matched, not templated</div>
    <div class="badge-chip">Free, no signup</div>
</div>
""", unsafe_allow_html=True)
st.markdown('<div class="hero-title">Tailor your resume<br>to any <span>job</span> in seconds</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-tag">Upload your resume and paste a job description. Get back a tailored, ATS-ready resume as a PDF, scored by AI against the role.</div>', unsafe_allow_html=True)

st.markdown('<div class="form-card">', unsafe_allow_html=True)

try:
    GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", ""))
except Exception:
    GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")

# ─────────────────────────────────────────────
# FILE TEXT EXTRACTION
# ─────────────────────────────────────────────
def extract_text(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    if name.endswith(".pdf"):
        text = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)
        return "\n".join(text)

    if name.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs if p.text)

    return ""


# ─────────────────────────────────────────────
# CORE AI LOGIC
# ─────────────────────────────────────────────
def call_groq(resume_text, jd_text):
    client = groq.Groq(api_key=GROQ_API_KEY)
    prompt = f"""
You are an expert ATS resume specialist. Analyze the resume against the job description
and produce a tailored version. Return ONLY a JSON object, no markdown, no text before or after.

{{
  "candidate_name": "full name from resume",
  "email": "email or empty string",
  "phone": "phone or empty string",
  "location": "city or empty string",
  "match_score": 75,
  "summary": "2-3 sentence professional summary tailored to the job description",
  "work_experience": [{{
    "title": "job title exactly as in resume",
    "company": "company name exactly as in resume",
    "dates": "dates exactly as in resume",
    "bullets": ["bullet rewritten with JD keywords", "bullet", "bullet"]
  }}],
  "education": [{{
    "degree": "degree exactly as in resume",
    "institution": "institution exactly as in resume",
    "year": "year exactly as in resume"
  }}],
  "skills": ["skill1","skill2","skill3","skill4","skill5","skill6"]
}}

STRICT RULES:
- Never fabricate experience, skills, or education not present in the original resume
- Rewrite bullet points using relevant keywords from the job description
- Keep dates, company names, and institutions exactly as original
- match_score must be a number, not a string, reflecting genuine keyword and experience overlap between the resume and job description
- Return ONLY pure JSON

RESUME:
{resume_text}

JOB DESCRIPTION:
{jd_text}
"""
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4,
        max_tokens=1800,
    )
    return response.choices[0].message.content


def parse_json(text):
    text = text.strip()
    if "```" in text:
        text = "\n".join(l for l in text.split("\n") if not l.strip().startswith("```"))
    start, end = text.find("{"), text.rfind("}") + 1
    return json.loads(text[start:end])


def build_pdf(data):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch,
        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
    )
    styles = {
        "name": ParagraphStyle("name", fontName="Helvetica-Bold", fontSize=18, spaceAfter=2, textColor=colors.HexColor("#14161B")),
        "contact": ParagraphStyle("contact", fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#555555"), spaceAfter=10),
        "section": ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=10.5, textColor=colors.HexColor("#B8863A"), spaceBefore=12, spaceAfter=4),
        "summary": ParagraphStyle("summary", fontName="Helvetica-Oblique", fontSize=9.5, leading=13, spaceAfter=4),
        "jobtitle": ParagraphStyle("jobtitle", fontName="Helvetica-Bold", fontSize=10, spaceBefore=6),
        "jobmeta": ParagraphStyle("jobmeta", fontName="Helvetica-Oblique", fontSize=8.5, textColor=colors.HexColor("#666666"), spaceAfter=2),
        "bullet": ParagraphStyle("bullet", fontName="Helvetica", fontSize=9, leftIndent=14, leading=12.5, spaceAfter=2),
        "body": ParagraphStyle("body", fontName="Helvetica", fontSize=9, leading=13),
    }

    story = []
    story.append(Paragraph(data.get("candidate_name", "Candidate"), styles["name"]))
    contact = " · ".join(x for x in [data.get("email"), data.get("phone"), data.get("location")] if x)
    if contact:
        story.append(Paragraph(contact, styles["contact"]))
    story.append(HRFlowable(width="100%", thickness=0.75, color=colors.HexColor("#dddddd")))

    if data.get("summary"):
        story.append(Paragraph("PROFESSIONAL SUMMARY", styles["section"]))
        story.append(Paragraph(data["summary"], styles["summary"]))

    jobs = data.get("work_experience", [])
    if jobs:
        story.append(Paragraph("WORK EXPERIENCE", styles["section"]))
        for job in jobs:
            story.append(Paragraph(f"{job.get('title','')} — {job.get('company','')}", styles["jobtitle"]))
            if job.get("dates"):
                story.append(Paragraph(job["dates"], styles["jobmeta"]))
            for b in job.get("bullets", [])[:4]:
                story.append(Paragraph(f"• {b}", styles["bullet"]))

    edu = data.get("education", [])
    if edu:
        story.append(Paragraph("EDUCATION", styles["section"]))
        for e in edu:
            line = f"{e.get('degree','')} — {e.get('institution','')} ({e.get('year','')})"
            story.append(Paragraph(line, styles["body"]))

    skills = data.get("skills", [])
    if skills:
        story.append(Paragraph("SKILLS", styles["section"]))
        story.append(Paragraph(" · ".join(skills), styles["body"]))

    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#eeeeee")))
    story.append(Paragraph("Tailored with ResumeReflect", ParagraphStyle("footer", fontName="Helvetica-Oblique", fontSize=7, textColor=colors.HexColor("#aaaaaa"), alignment=1, spaceBefore=6)))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# UI FLOW
# ─────────────────────────────────────────────
if "result" not in st.session_state:
    st.session_state.result = None

st.markdown('<div class="step-label"><div class="step-num">1</div><div class="step-text">Upload your resume</div></div>', unsafe_allow_html=True)
resume_file = st.file_uploader(" ", type=["pdf", "docx"], label_visibility="collapsed")

st.markdown('<div class="step-label" style="margin-top:1.4rem;"><div class="step-num">2</div><div class="step-text">Paste the job description</div></div>', unsafe_allow_html=True)
jd = st.text_area(" ", height=170, placeholder="Paste the job posting text here", label_visibility="collapsed")

clicked = st.button("Tailor my resume")
st.markdown('</div>', unsafe_allow_html=True)

if clicked:
    if not GROQ_API_KEY:
        st.error("Server is missing its Groq API key. Contact the site owner.")
    elif resume_file is None:
        st.warning("Please upload your resume first.")
    elif len(jd.strip()) < 30:
        st.warning("Please paste a bit more of the job description.")
    else:
        with st.spinner("Reading your resume and matching it to the role..."):
            try:
                resume_text = extract_text(resume_file)
                if len(resume_text.strip()) < 30:
                    st.error("Couldn't read text from that file. Try a different PDF or Word export.")
                    st.session_state.result = None
                else:
                    raw = call_groq(resume_text, jd)
                    data = parse_json(raw)
                    pdf_bytes = build_pdf(data)
                    st.session_state.result = (data, pdf_bytes)
            except Exception:
                st.error("Something went wrong. Please try again.")
                st.session_state.result = None

if st.session_state.result:
    data, pdf_bytes = st.session_state.result
    score = data.get("match_score", 0)
    bar_color = "#3FA383" if score >= 75 else "#D6A749" if score >= 50 else "#C1573D"

    st.markdown(f"""
    <div class="result-card">
        <div class="score-num">{score}</div>
        <div class="score-label">AI-matched score out of 100</div>
        <div class="score-bar-bg"><div class="score-bar-fill" style="width:{score}%; background:{bar_color};"></div></div>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        "Download tailored resume (PDF)",
        data=pdf_bytes,
        file_name="tailored_resume.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
