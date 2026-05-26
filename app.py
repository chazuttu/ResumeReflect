import streamlit as st
import groq
import pdfplumber
import requests
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
RAZORPAY_BASIC = st.secrets.get("RAZORPAY_BASIC", "#")
RAZORPAY_PRO = st.secrets.get("RAZORPAY_PRO", "#")
RAZORPAY_YEARLY = st.secrets.get("RAZORPAY_YEARLY", "#")
SHEET_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzN99UwHn4Bt1mJ4MMS5ZSV-cysoTC_ac6d6oMNkWB_JAGb1i2vqBX3RmrCDqIsla3G/exec"
TOOL_NAME = "ResumeReflect"

st.set_page_config(
    page_title="ResumeReflect - AI Resume Tailor",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# PROFESSIONAL STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --bg: #f8fafc;
    --surface: #ffffff;
    --border: #e2e8f0;
    --text: #0f172a;
    --text-muted: #64748b;
    --accent: #06b6d4;
    --success: #10b981;
    --warning: #f59e0b;
    --danger: #ef4444;
}

html, body, .stApp {
    background: var(--bg) !important;
    font-family: 'DM Sans', sans-serif !important;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 2rem 2.5rem !important; max-width: 1200px !important; }

h1, h2, h3 { font-family: 'Syne', sans-serif !important; }
h1 { font-size: 2.5rem !important; font-weight: 800 !important; }
h2 { font-size: 1.8rem !important; font-weight: 700 !important; margin: 2rem 0 1.5rem !important; }
h3 { font-size: 1.3rem !important; font-weight: 600 !important; }

/* TABS STYLING */
.stTabs [data-baseweb="tab-list"] {
    border-bottom: 2px solid var(--border) !important;
    gap: 3rem !important;
    padding-bottom: 1rem !important;
}
.stTabs [data-baseweb="tab"] {
    padding: 1rem 1.5rem !important;
    border-bottom: 3px solid transparent !important;
    color: var(--text-muted) !important;
    font-weight: 600 !important;
}
.stTabs [data-baseweb="tab"][aria-selected="true"] {
    border-bottom-color: var(--accent) !important;
    color: var(--accent) !important;
}

/* BUTTONS */
.stButton > button {
    background: linear-gradient(135deg, #06b6d4, #0f766e) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 12px 24px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    width: 100% !important;
    height: 44px !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 20px rgba(6,182,212,0.3) !important;
}

/* FORMS */
.stTextInput input, .stTextArea textarea {
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
    padding: 12px !important;
    font-size: 14px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(6,182,212,0.1) !important;
}

/* CARDS */
.metric-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}

/* ATS SCORE DISPLAY */
.ats-score-box {
    text-align: center;
    padding: 2rem;
    border-radius: 12px;
    margin: 2rem 0;
}

.ats-score-box.high {
    background: rgba(16,185,129,0.1);
    border: 2px solid #10b981;
}
.ats-score-box.medium {
    background: rgba(245,158,11,0.1);
    border: 2px solid #f59e0b;
}
.ats-score-box.low {
    background: rgba(239,68,68,0.1);
    border: 2px solid #ef4444;
}

.ats-number { font-size: 4rem; font-weight: 800; margin-bottom: 0.5rem; }
.ats-label { font-size: 1.2rem; font-weight: 600; }

/* KEYWORD SECTION */
.keyword-section {
    background: white;
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 2rem;
    margin: 1.5rem 0;
}

.keyword-section h3 { margin-top: 0; }

.keyword-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 1rem;
    margin: 1rem 0;
}

.keyword-item {
    padding: 1rem;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    text-align: center;
}

.keyword-matched {
    background: rgba(16,185,129,0.15);
    color: #047857;
    border: 1px solid #10b981;
}

.keyword-missing {
    background: rgba(239,68,68,0.15);
    color: #991b1b;
    border: 1px solid #ef4444;
}

/* PRICING */
.pricing-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 2rem;
    margin: 2rem 0;
}

.plan-card {
    background: white;
    border: 2px solid var(--border);
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
}

.plan-card.featured {
    border-color: var(--accent);
    background: linear-gradient(135deg, rgba(6,182,212,0.05), rgba(15,118,110,0.05));
}

.plan-price { font-size: 2.5rem; font-weight: 800; margin: 1rem 0; }
.plan-features { list-style: none; padding: 0; text-align: left; }
.plan-features li { padding: 0.5rem 0; display: flex; gap: 8px; }
.plan-features li:before { content: "✓"; font-weight: 800; color: var(--success); }

/* DIVIDER */
.divider { height: 1px; background: linear-gradient(90deg, transparent, var(--border), transparent); margin: 2rem 0; }

@media (max-width: 768px) {
    .stTabs [data-baseweb="tab-list"] { gap: 1rem !important; }
    .stTabs [data-baseweb="tab"] { padding: 0.8rem !important; }
    .keyword-grid { grid-template-columns: 1fr; }
    .pricing-grid { grid-template-columns: 1fr; }
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
defaults = {
    "step": 0,
    "resume_text": "",
    "jd_text": "",
    "email": "",
    "analysis": None,
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────
def extract_text_from_pdf(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        return text
    except:
        return None

def extract_text_from_docx(docx_file):
    try:
        doc = Document(docx_file)
        return "\n".join([para.text for para in doc.paragraphs])
    except:
        return None

def analyze_resume_groq(resume_text, jd_text):
    """Get detailed ATS analysis from Groq"""
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        
        prompt = f"""Analyze this resume against the job description. Return ONLY valid JSON.

JOB DESCRIPTION:
{jd_text[:2000]}

RESUME:
{resume_text[:2000]}

Return this exact JSON structure (NO other text):
{{
  "ats_score": 0-100,
  "matched_keywords": ["keyword1", "keyword2", "keyword3"],
  "missing_keywords": ["keyword1", "keyword2", "keyword3"],
  "strengths": ["strength1", "strength2", "strength3"],
  "improvements": ["improvement1", "improvement2", "improvement3"],
  "ats_analysis": "2-3 sentence analysis"
}}"""
        
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=800
        )
        
        text = response.choices[0].message.content.strip()
        if "```" in text:
            text = text.split("```")[1].split("```")[0] if "```json" in text else text.split("```")[1]
        
        return json.loads(text.strip())
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def generate_tailored_resume(resume_text, jd_text, missing_keywords):
    """Generate professionally formatted tailored resume"""
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        
        prompt = f"""Rewrite this resume to match the job description perfectly.

INSTRUCTIONS:
1. Keep the same person, dates, and actual experience
2. Add these keywords naturally: {', '.join(missing_keywords[:15])}
3. Reorder sections by job relevance
4. Add quantified achievements
5. Use action verbs
6. Make it professional and realistic

ORIGINAL RESUME:
{resume_text}

JOB DESCRIPTION:
{jd_text[:1500]}

IMPORTANT: Return ONLY the resume content in professional format. No explanations."""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
            max_tokens=1500
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating resume: {str(e)}"

def create_ats_report_word(analysis, job_description, resume_text):
    """Create professional ATS report in Word format"""
    try:
        doc = Document()
        
        # Title
        title = doc.add_paragraph()
        title_run = title.add_run("ATS COMPATIBILITY ANALYSIS REPORT")
        title_run.font.size = Pt(18)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(15, 23, 42)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Date
        date_para = doc.add_paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}")
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_para_format = date_para.runs[0]
        date_para_format.font.size = Pt(10)
        date_para_format.font.color.rgb = RGBColor(100, 116, 139)
        
        doc.add_paragraph()  # Space
        
        # ATS SCORE
        score = analysis.get("ats_score", 0)
        score_para = doc.add_paragraph()
        score_para.add_run("Overall ATS Score: ").bold = True
        score_run = score_para.add_run(f"{score}%")
        score_run.font.size = Pt(14)
        if score >= 70:
            score_run.font.color.rgb = RGBColor(16, 185, 129)
        elif score >= 50:
            score_run.font.color.rgb = RGBColor(245, 158, 11)
        else:
            score_run.font.color.rgb = RGBColor(239, 68, 68)
        
        doc.add_paragraph("This score indicates how well your resume aligns with ATS systems and the job requirements.")
        doc.add_paragraph()
        
        # MATCHED KEYWORDS
        doc.add_heading("Matched Keywords", level=2)
        matched = analysis.get("matched_keywords", [])
        if matched:
            doc.add_paragraph(f"Found {len(matched)} keywords from the job description in your resume:")
            for kw in matched[:10]:
                doc.add_paragraph(f"✓ {kw}", style='List Bullet')
        else:
            doc.add_paragraph("No matching keywords found.", style='Normal')
        
        doc.add_paragraph()
        
        # MISSING KEYWORDS
        doc.add_heading("Missing Keywords (Critical)", level=2)
        missing = analysis.get("missing_keywords", [])
        if missing:
            doc.add_paragraph(f"These {len(missing)} keywords are present in the job description but missing from your resume:")
            for kw in missing[:10]:
                doc.add_paragraph(f"✗ {kw}", style='List Bullet')
            doc.add_paragraph()
            doc.add_paragraph("RECOMMENDATION: Add these keywords naturally to your resume, especially in the professional summary, experience section, and skills.")
        else:
            doc.add_paragraph("All key keywords are present in your resume.")
        
        doc.add_paragraph()
        
        # STRENGTHS
        doc.add_heading("Resume Strengths", level=2)
        strengths = analysis.get("strengths", [])
        if strengths:
            for strength in strengths[:5]:
                doc.add_paragraph(f"• {strength}", style='List Bullet')
        
        doc.add_paragraph()
        
        # IMPROVEMENTS
        doc.add_heading("Recommended Improvements", level=2)
        improvements = analysis.get("improvements", [])
        if improvements:
            for improvement in improvements[:5]:
                doc.add_paragraph(f"→ {improvement}", style='List Bullet')
        
        doc.add_paragraph()
        doc.add_paragraph()
        
        # FOOTER
        footer = doc.add_paragraph("ResumeReflect - AI Resume Optimization")
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_run = footer.runs[0]
        footer_run.font.size = Pt(9)
        footer_run.font.color.rgb = RGBColor(100, 116, 139)
        footer_run.italic = True
        
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        return doc_io
    except Exception as e:
        st.error(f"Error creating report: {str(e)}")
        return None

def create_resume_word(tailored_resume):
    """Create professional resume in Word format"""
    try:
        doc = Document()
        
        # Parse and format resume
        for line in tailored_resume.split('\n'):
            line = line.strip()
            if not line:
                doc.add_paragraph()
                continue
            
            # Headers (all caps or with colons)
            if line.isupper() or (line.endswith(':') and len(line) < 50):
                p = doc.add_heading(line.rstrip(':'), level=2)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
            # Bullet points
            elif line.startswith('•') or line.startswith('-'):
                p = doc.add_paragraph(line.lstrip('•-').strip(), style='List Bullet')
                p.paragraph_format.left_indent = Inches(0.5)
            # Regular text
            else:
                p = doc.add_paragraph(line)
                p.paragraph_format.space_after = Pt(6)
        
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        return doc_io
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def log_to_sheet(data_type, email):
    try:
        requests.post(SHEET_SCRIPT_URL, json={
            "type": data_type,
            "email": email,
            "tool_name": TOOL_NAME,
            "timestamp": datetime.now().isoformat()
        }, timeout=5)
    except:
        pass

# ─────────────────────────────────────────────
# LANDING PAGE
# ─────────────────────────────────────────────
if st.session_state.step == 0:
    st.markdown('''
    <div style="text-align:center;padding:3rem 0;">
    <h1>⚡ Land More Interview Calls</h1>
    <p style="font-size:1.1rem;color:#64748b;max-width:700px;margin:1rem auto;">
    Get your resume past ATS filters with AI. Understand exactly what keywords are missing and get a professionally tailored resume ready to apply.
    </p>
    </div>
    ''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('## How It Works', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3, gap="large")
    
    with col1:
        st.markdown('''
        <div class="metric-card">
        <h3 style="margin-top:0">📄 Upload Resume</h3>
        <p>PDF, DOCX, or paste text directly</p>
        </div>
        ''', unsafe_allow_html=True)
    
    with col2:
        st.markdown('''
        <div class="metric-card">
        <h3 style="margin-top:0">🎯 Add Job Description</h3>
        <p>Paste the job posting you're targeting</p>
        </div>
        ''', unsafe_allow_html=True)
    
    with col3:
        st.markdown('''
        <div class="metric-card">
        <h3 style="margin-top:0">⚡ Get Results</h3>
        <p>Professional ATS report and tailored resume</p>
        </div>
        ''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('## Pricing', unsafe_allow_html=True)
    st.markdown('''<div class="pricing-grid">
    <div class="plan-card">
    <h3 style="margin-top:0">Free</h3>
    <div class="plan-price">$0</div>
    <ul class="plan-features">
    <li>ATS Score Analysis</li>
    <li>Detailed Keyword Report</li>
    <li>1 Resume Rewrite</li>
    <li>Download ATS Report</li>
    </ul>
    </div>
    
    <div class="plan-card featured">
    <h3 style="margin-top:0">Pro</h3>
    <div class="plan-price">$5</div>
    <ul class="plan-features">
    <li>Everything in Free</li>
    <li>Unlimited Rewrites</li>
    <li>Interview Kit</li>
    <li>Download Resume</li>
    </ul>
    </div>
    
    <div class="plan-card">
    <h3 style="margin-top:0">Premium</h3>
    <div class="plan-price">$12/mo</div>
    <ul class="plan-features">
    <li>Everything in Pro</li>
    <li>LinkedIn Profile</li>
    <li>Job Matching</li>
    <li>Email Support</li>
    </ul>
    </div>
    </div>''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("✨ Start Free Analysis", use_container_width=True, key="start"):
            st.session_state.step = 1
            st.rerun()

# ─────────────────────────────────────────────
# STEP 1: UPLOAD
# ─────────────────────────────────────────────
elif st.session_state.step == 1:
    st.markdown('## Upload Your Resume', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 2], gap="large")
    
    with col1:
        st.markdown('**Email Address**')
        email = st.text_input("", placeholder="your@email.com", label_visibility="collapsed", key="email_in")
    
    with col2:
        st.markdown('**Resume**')
        tab1, tab2 = st.tabs(["📁 Upload File", "📝 Paste Text"])
        
        with tab1:
            uploaded = st.file_uploader("Upload PDF or DOCX", type=["pdf", "docx"], label_visibility="collapsed")
            if uploaded:
                if uploaded.type == "application/pdf":
                    text = extract_text_from_pdf(uploaded)
                else:
                    text = extract_text_from_docx(uploaded)
                if text:
                    st.session_state.resume_text = text
                    st.success("✅ Uploaded!")
        
        with tab2:
            pasted = st.text_area("Paste your resume", placeholder="Paste resume text...", height=200, label_visibility="collapsed")
            if pasted:
                st.session_state.resume_text = pasted

    col1, col2 = st.columns(2, gap="large")
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 0
            st.rerun()
    
    with col2:
        if st.button("Continue →", use_container_width=True):
            if not email or "@" not in email:
                st.error("❌ Invalid email")
            elif not st.session_state.resume_text or len(st.session_state.resume_text) < 100:
                st.error("❌ Resume too short")
            else:
                st.session_state.email = email
                log_to_sheet("upload", email)
                st.session_state.step = 2
                st.rerun()

# ─────────────────────────────────────────────
# STEP 2: JOB DESCRIPTION
# ─────────────────────────────────────────────
elif st.session_state.step == 2:
    st.markdown('## Add Job Description', unsafe_allow_html=True)
    
    jd = st.text_area(
        "Paste the complete job posting",
        placeholder="Paste full job description here...",
        height=350,
        label_visibility="collapsed"
    )
    
    if jd:
        st.session_state.jd_text = jd

    col1, col2 = st.columns(2, gap="large")
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    
    with col2:
        if st.button("Analyze Resume →", use_container_width=True):
            if not st.session_state.jd_text or len(st.session_state.jd_text) < 100:
                st.error("❌ Please paste complete job description")
            else:
                st.session_state.step = 3
                st.rerun()

# ─────────────────────────────────────────────
# STEP 3: ANALYSIS & RESULTS
# ─────────────────────────────────────────────
elif st.session_state.step == 3:
    st.markdown('## Resume Analysis Results', unsafe_allow_html=True)
    
    with st.spinner("🔍 Analyzing your resume..."):
        analysis = analyze_resume_groq(st.session_state.resume_text, st.session_state.jd_text)
        st.session_state.analysis = analysis
    
    if analysis:
        score = analysis.get("ats_score", 0)
        matched = analysis.get("matched_keywords", [])
        missing = analysis.get("missing_keywords", [])
        strengths = analysis.get("strengths", [])
        improvements = analysis.get("improvements", [])
        
        # ATS SCORE
        if score >= 70:
            color_class = "high"
            status = "✓ Good Match"
        elif score >= 50:
            color_class = "medium"
            status = "⚠ Needs Work"
        else:
            color_class = "low"
            status = "✗ Low Match"
        
        st.markdown(f'''
        <div class="ats-score-box {color_class}">
        <div class="ats-number">{score}%</div>
        <div class="ats-label">{status}</div>
        </div>
        ''', unsafe_allow_html=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # KEYWORD ANALYSIS
        col1, col2 = st.columns(2, gap="large")
        
        with col1:
            st.markdown('''
            <div class="keyword-section">
            <h3>✓ Matched Keywords</h3>
            <p style="color:#64748b;font-size:14px">Keywords from the job posting found in your resume:</p>
            <div class="keyword-grid">
            ''', unsafe_allow_html=True)
            
            for kw in matched:
                st.markdown(f'<div class="keyword-item keyword-matched">{kw}</div>', unsafe_allow_html=True)
            
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('''
            <div class="keyword-section">
            <h3>⚠ Missing Keywords</h3>
            <p style="color:#64748b;font-size:14px">Important keywords you should add to your resume:</p>
            <div class="keyword-grid">
            ''', unsafe_allow_html=True)
            
            for kw in missing:
                st.markdown(f'<div class="keyword-item keyword-missing">{kw}</div>', unsafe_allow_html=True)
            
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # INSIGHTS
        col1, col2 = st.columns(2, gap="large")
        
        with col1:
            st.markdown('''<div class="metric-card">
            <h3 style="margin-top:0">💪 Your Strengths</h3>
            ''', unsafe_allow_html=True)
            for s in strengths:
                st.markdown(f"✓ {s}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('''<div class="metric-card">
            <h3 style="margin-top:0">🎯 Improvements Needed</h3>
            ''', unsafe_allow_html=True)
            for imp in improvements:
                st.markdown(f"→ {imp}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # DOWNLOADS & ACTIONS
        st.markdown('## Download Your Reports', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2, gap="large")
        
        with col1:
            # ATS Report
            ats_doc = create_ats_report_word(analysis, st.session_state.jd_text, st.session_state.resume_text)
            if ats_doc:
                st.download_button(
                    "📊 Download ATS Report (Word)",
                    ats_doc,
                    "ATS_Report.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        
        with col2:
            st.info("Generate your AI-tailored resume below →")
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        st.markdown('## Generate Tailored Resume', unsafe_allow_html=True)
        
        if st.button("✨ Create Tailored Resume", use_container_width=True):
            with st.spinner("Creating your professional resume..."):
                tailored = generate_tailored_resume(
                    st.session_state.resume_text,
                    st.session_state.jd_text,
                    missing
                )
            
            if tailored:
                # Preview
                st.markdown('### Resume Preview', unsafe_allow_html=True)
                st.text_area("", value=tailored, height=400, disabled=True, label_visibility="collapsed")
                
                # Downloads
                col1, col2, col3 = st.columns(3, gap="large")
                
                with col1:
                    st.download_button(
                        "⬇ Text Format",
                        tailored,
                        "tailored_resume.txt",
                        "text/plain",
                        use_container_width=True
                    )
                
                with col2:
                    resume_doc = create_resume_word(tailored)
                    if resume_doc:
                        st.download_button(
                            "⬇ Word Format",
                            resume_doc,
                            "tailored_resume.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                
                with col3:
                    st.info("👉 Ready to apply!")
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # NAVIGATION
        col1, col2, col3 = st.columns(3, gap="large")
        
        with col1:
            if st.button("← Back", use_container_width=True):
                st.session_state.step = 2
                st.rerun()
        
        with col2:
            if st.button("🔄 Analyze Another", use_container_width=True):
                for k in defaults:
                    st.session_state[k] = defaults[k]
                st.session_state.step = 1
                st.rerun()
        
        with col3:
            if st.button("⭐ Upgrade", use_container_width=True):
                st.session_state.step = 4
                st.rerun()

# ─────────────────────────────────────────────
# STEP 4: PRICING
# ─────────────────────────────────────────────
elif st.session_state.step == 4:
    st.markdown('## Upgrade Your Plan', unsafe_allow_html=True)
    
    st.markdown('''<div class="pricing-grid">
    <div class="plan-card">
    <h3 style="margin-top:0">Pro</h3>
    <div class="plan-price">$5</div>
    <ul class="plan-features">
    <li>Unlimited Resume Rewrites</li>
    <li>Interview Prep Kit</li>
    <li>One-time Payment</li>
    <li>Email Support</li>
    </ul>
    </div>
    
    <div class="plan-card featured">
    <h3 style="margin-top:0">Premium</h3>
    <div class="plan-price">$12/mo</div>
    <ul class="plan-features">
    <li>Everything in Pro</li>
    <li>LinkedIn Profile Optimization</li>
    <li>Job Matching</li>
    <li>Priority Support</li>
    </ul>
    </div>
    </div>''', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3, gap="large")
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    
    with col2:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_BASIC}" target="_blank"><button style="width:100%;height:44px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:10px;cursor:pointer;font-weight:700;font-family:Syne,sans-serif;font-size:14px">Pay Pro $5</button></a>', unsafe_allow_html=True)
    
    with col3:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_PRO}" target="_blank"><button style="width:100%;height:44px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:10px;cursor:pointer;font-weight:700;font-family:Syne,sans-serif;font-size:14px">Pay Premium $12</button></a>', unsafe_allow_html=True)

# FOOTER
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown('''
<div style="text-align:center;padding:2rem 0;color:#64748b;font-size:13px">
<strong>⚡ ResumeReflect</strong><br>
Professional resume optimization powered by AI<br>
<small>Your data is secure and never stored</small>
</div>
''', unsafe_allow_html=True)
