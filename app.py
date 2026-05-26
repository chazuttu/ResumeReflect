import streamlit as st
import groq
import pdfplumber
import requests
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os
import json
import re
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIG & SECRETS
# ─────────────────────────────────────────────
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
RAZORPAY_BASIC = st.secrets.get("RAZORPAY_BASIC", "#")
RAZORPAY_PRO = st.secrets.get("RAZORPAY_PRO", "#")
RAZORPAY_YEARLY = st.secrets.get("RAZORPAY_YEARLY", "#")
SHEET_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzN99UwHn4Bt1mJ4MMS5ZSV-cysoTC_ac6d6oMNkWB_JAGb1i2vqBX3RmrCDqIsla3G/exec"
TOOL_NAME = "ResumeReflect"

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ResumeReflect - Land More Interview Calls",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# MODERN DESIGN SYSTEM
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --primary: #1f2937;
    --secondary: #0f766e;
    --accent: #06b6d4;
    --success: #10b981;
    --warning: #f59e0b;
    --danger: #ef4444;
    --bg: #f8fafc;
    --surface: #ffffff;
    --border: #e2e8f0;
    --text: #0f172a;
    --text-muted: #64748b;
    --radius: 12px;
}

html, body, .stApp {
    background: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 2rem 1.5rem !important; max-width: 1100px !important; margin: 0 auto; }

h1, h2, h3 { font-family: 'Syne', sans-serif !important; color: var(--primary) !important; }
h1 { font-size: clamp(2rem, 5vw, 3.2rem) !important; font-weight: 800 !important; line-height: 1.1 !important; margin: 0 0 1rem !important; }
h2 { font-size: 1.8rem !important; font-weight: 700 !important; margin: 1.5rem 0 0.8rem !important; }

.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.5rem;
    margin-bottom: 1rem;
    transition: all 0.3s ease;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}

.stButton > button {
    background: linear-gradient(135deg, #06b6d4, #0f766e) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 12px 24px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    width: 100% !important;
    transition: all 0.3s !important;
    min-height: 44px !important;
    cursor: pointer !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 20px rgba(6,182,212,0.3) !important;
}

.stFileUploader > div {
    background: var(--bg) !important;
    border: 2px dashed var(--accent) !important;
    border-radius: 12px !important;
    padding: 2rem !important;
}

.stTextArea textarea, .stTextInput input {
    background: var(--surface) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
    font-size: 14px !important;
}
.stTextArea textarea:focus, .stTextInput input:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(6,182,212,0.1) !important;
}

.pricing-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 1.5rem;
    margin: 2rem 0;
}
.plan-card {
    background: var(--surface);
    border: 2px solid var(--border);
    border-radius: var(--radius);
    padding: 2rem 1.5rem;
    text-align: center;
}
.plan-card.featured {
    border-color: var(--accent);
    background: linear-gradient(135deg, rgba(6,182,212,0.05), rgba(15,118,110,0.05));
}
.plan-name { font-size: 14px; font-weight: 700; text-transform: uppercase; color: var(--text-muted); margin-bottom: 0.5rem; }
.plan-price { font-size: 2.2rem; font-weight: 800; color: var(--primary); margin: 0.5rem 0; }
.plan-period { font-size: 12px; color: var(--text-muted); margin-bottom: 1.2rem; }
.plan-features { list-style: none; padding: 0; margin: 1.5rem 0; text-align: left; }
.plan-features li { font-size: 13px; color: var(--text); padding: 0.5rem 0; display: flex; gap: 8px; }
.plan-features li:before { content: "✓"; color: var(--success); font-weight: 800; }

.stat-box {
    background: linear-gradient(135deg, var(--accent), var(--secondary));
    color: white;
    padding: 1.5rem;
    border-radius: var(--radius);
    text-align: center;
}
.stat-number { font-size: 2rem; font-weight: 800; }
.stat-label { font-size: 12px; opacity: 0.9; text-transform: uppercase; }

.info-box {
    background: rgba(6,182,212,0.1);
    border-left: 4px solid var(--accent);
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
}

.keyword-tag {
    display: inline-block;
    padding: 6px 12px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    margin: 4px;
}
.keyword-matched { background: rgba(16,185,129,0.2); color: var(--success); }
.keyword-missing { background: rgba(239,68,68,0.2); color: var(--danger); }

.divider { height: 1px; background: linear-gradient(90deg, transparent, var(--border), transparent); margin: 2rem 0; }

.footer {
    text-align: center;
    padding: 2rem 0;
    border-top: 1px solid var(--border);
    margin-top: 3rem;
    color: var(--text-muted);
    font-size: 13px;
}

@media (max-width: 768px) {
    h1 { font-size: 1.8rem !important; }
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
    "analysis_data": None,
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

def analyze_resume_with_groq(resume_text, jd_text):
    """Analyze resume and return detailed data"""
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        
        prompt = f"""Analyze this resume against the job description. Return ONLY valid JSON with NO other text.

JOB DESCRIPTION:
{jd_text[:1500]}

RESUME:
{resume_text[:2000]}

Return ONLY this JSON structure (no other text):
{{
  "ats_score": 0-100,
  "matched_keywords": ["keyword1", "keyword2"],
  "missing_keywords": ["keyword1", "keyword2"],
  "strengths": ["strength1", "strength2"],
  "improvements": ["improvement1", "improvement2"],
  "ats_analysis": "Brief 2-3 line summary"
}}"""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=800
        )
        
        response_text = response.choices[0].message.content.strip()
        
        # Clean up response if it has markdown code blocks
        if "```json" in response_text:
            response_text = response_text.split("```json")[1].split("```")[0]
        elif "```" in response_text:
            response_text = response_text.split("```")[1].split("```")[0]
        
        return json.loads(response_text.strip())
    except Exception as e:
        st.error(f"Analysis error: {str(e)}")
        return None

def generate_tailored_resume(resume_text, jd_text, matched_keywords, missing_keywords):
    """Generate tailored resume"""
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        
        prompt = f"""Rewrite this resume to match the job description. Be specific and professional.

KEY INSTRUCTIONS:
1. Keep the same person and dates
2. Add these missing keywords naturally: {', '.join(missing_keywords[:10])}
3. Reorder experience sections by relevance to the job
4. Add metrics and achievements
5. Use action verbs

ORIGINAL RESUME:
{resume_text}

JOB DESCRIPTION:
{jd_text[:1000]}

Provide ONLY the tailored resume text. No headers, no markdown, just clean resume text."""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
            max_tokens=1500
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating resume: {str(e)}"

def create_word_document(tailored_resume):
    """Create downloadable Word document"""
    try:
        doc = Document()
        
        for paragraph in tailored_resume.split('\n'):
            if paragraph.strip():
                p = doc.add_paragraph(paragraph.strip())
                p.style = 'Normal'
                if any(word in paragraph.upper() for word in ['SUMMARY', 'EXPERIENCE', 'EDUCATION', 'SKILLS', 'PROJECTS']):
                    p.style = 'Heading 2'
        
        # Save to bytes
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        return doc_io
    except:
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
    <div style="text-align:center;padding:2rem 0;">
    <h1>⚡ Land More Interview Calls</h1>
    <p style="font-size:1.1rem;color:#64748b;max-width:600px;margin:0 auto 2rem;">
    Get your resume past ATS filters. See exactly which keywords you're missing. Get an AI-tailored resume ready to apply.
    </p>
    </div>
    ''', unsafe_allow_html=True)

    # STATS
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="stat-box"><div class="stat-number">3x</div><div class="stat-label">More Interviews</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="stat-box"><div class="stat-number">2m</div><div class="stat-label">Avg Analysis Time</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # HOW IT WORKS
    st.markdown('## How It Works', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3, gap="large")
    with col1:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">📄</div>
        <h3>Upload Resume</h3>
        <p>PDF, DOCX, or paste text</p>
        </div>''', unsafe_allow_html=True)
    
    with col2:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">🎯</div>
        <h3>Add Job Description</h3>
        <p>Paste the job posting</p>
        </div>''', unsafe_allow_html=True)
    
    with col3:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">⚡</div>
        <h3>Get Tailored Resume</h3>
        <p>ATS-optimized, ready to apply</p>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # PRICING
    st.markdown('## Simple Pricing', unsafe_allow_html=True)
    
    # Use proper columns for pricing cards instead of markdown
    col1, col2, col3 = st.columns(3, gap="large")
    
    with col1:
        st.markdown("""
        <div class="plan-card">
        <div class="plan-name">🚀 Free</div>
        <div class="plan-price">$0</div>
        <div class="plan-period">forever</div>
        <ul class="plan-features">
        <li>ATS score analysis</li>
        <li>Matched keywords</li>
        <li>Missing keywords</li>
        <li>1 resume rewrite</li>
        <li>Download as Text/Word</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="plan-card featured">
        <div class="plan-name">⭐ Pro</div>
        <div class="plan-price">$5</div>
        <div class="plan-period">one-time</div>
        <ul class="plan-features">
        <li>Everything in Free</li>
        <li>Unlimited rewrites</li>
        <li>Interview prep kit</li>
        <li>Cover letter</li>
        <li>Priority support</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="plan-card">
        <div class="plan-name">👑 Premium</div>
        <div class="plan-price">$12</div>
        <div class="plan-period">per month</div>
        <ul class="plan-features">
        <li>Everything in Pro</li>
        <li>LinkedIn rewrite</li>
        <li>Job matching</li>
        <li>Weekly coaching</li>
        <li>1-on-1 chat</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("✨ Start Free Analysis", use_container_width=True):
            st.session_state.step = 1
            st.rerun()

# ─────────────────────────────────────────────
# STEP 1: UPLOAD RESUME & EMAIL
# ─────────────────────────────────────────────
elif st.session_state.step == 1:
    st.markdown('## Upload Your Resume', unsafe_allow_html=True)
    
    email = st.text_input("Email Address", placeholder="your@email.com", key="email_input")
    
    tab1, tab2 = st.tabs(["Upload File", "Paste Text"])
    
    with tab1:
        uploaded_file = st.file_uploader("Choose PDF or DOCX", type=["pdf", "docx"])
        if uploaded_file:
            if uploaded_file.type == "application/pdf":
                text = extract_text_from_pdf(uploaded_file)
            else:
                text = extract_text_from_docx(uploaded_file)
            
            if text:
                st.session_state.resume_text = text
                st.success("✅ Resume uploaded!")
    
    with tab2:
        pasted = st.text_area("Paste your resume", placeholder="Paste resume text here...", height=250)
        if pasted:
            st.session_state.resume_text = pasted

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 0
            st.rerun()
    
    with col2:
        if st.button("Continue →", use_container_width=True):
            if not email or "@" not in email:
                st.error("Please enter valid email")
            elif not st.session_state.resume_text or len(st.session_state.resume_text) < 100:
                st.error("Please provide a resume")
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
        "Paste the job posting",
        placeholder="Paste complete job description...",
        height=350,
        key="jd_input"
    )
    
    if jd:
        st.session_state.jd_text = jd

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    
    with col2:
        if st.button("Analyze Resume →", use_container_width=True):
            if not st.session_state.jd_text or len(st.session_state.jd_text) < 50:
                st.error("Please paste complete job description")
            else:
                st.session_state.step = 3
                st.rerun()

# ─────────────────────────────────────────────
# STEP 3: ANALYSIS RESULTS
# ─────────────────────────────────────────────
elif st.session_state.step == 3:
    st.markdown('## Your Resume Analysis', unsafe_allow_html=True)
    
    with st.spinner("🔍 Analyzing your resume..."):
        analysis = analyze_resume_with_groq(st.session_state.resume_text, st.session_state.jd_text)
        st.session_state.analysis_data = analysis
    
    if analysis:
        ats_score = analysis.get("ats_score", 65)
        matched = analysis.get("matched_keywords", [])
        missing = analysis.get("missing_keywords", [])
        strengths = analysis.get("strengths", [])
        improvements = analysis.get("improvements", [])
        
        # ATS SCORE
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if ats_score >= 70:
                color = "#10b981"
                msg = "Good Match"
            elif ats_score >= 50:
                color = "#f59e0b"
                msg = "Needs Work"
            else:
                color = "#ef4444"
                msg = "Low Match"
            
            st.markdown(f'''<div style="text-align:center;padding:2rem;border:3px solid {color};border-radius:12px;">
            <div style="font-size:3.5rem;font-weight:800;color:{color}">{ats_score}%</div>
            <div style="font-size:1.1rem;color:{color};font-weight:700">{msg}</div>
            </div>''', unsafe_allow_html=True)

        st.markdown('---')

        # MATCHED KEYWORDS
        if matched:
            st.markdown('### ✓ Matched Keywords')
            cols = st.columns(4)
            for i, kw in enumerate(matched):
                with cols[i % 4]:
                    st.markdown(f'<div class="keyword-tag keyword-matched">{kw}</div>', unsafe_allow_html=True)
        
        # MISSING KEYWORDS
        if missing:
            st.markdown('### ⚠️ Missing Keywords (Add These!)')
            cols = st.columns(4)
            for i, kw in enumerate(missing):
                with cols[i % 4]:
                    st.markdown(f'<div class="keyword-tag keyword-missing">{kw}</div>', unsafe_allow_html=True)
        
        # STRENGTHS
        if strengths:
            st.markdown('### 💪 Your Strengths')
            for s in strengths:
                st.markdown(f"✓ {s}")
        
        # IMPROVEMENTS
        if improvements:
            st.markdown('### 🎯 Improvements Needed')
            for imp in improvements:
                st.markdown(f"→ {imp}")

        st.markdown('---')

        # DOWNLOADS
        st.markdown('### Download Your ATS Report')
        report = f"""ATS ANALYSIS REPORT
{'='*50}

ATS SCORE: {ats_score}%

MATCHED KEYWORDS:
{', '.join(matched) if matched else 'None'}

MISSING KEYWORDS:
{', '.join(missing) if missing else 'None'}

STRENGTHS:
{chr(10).join(['• ' + s for s in strengths]) if strengths else 'None'}

IMPROVEMENTS:
{chr(10).join(['• ' + i for i in improvements]) if improvements else 'None'}
"""
        
        st.download_button(
            "📊 Download ATS Report (Text)",
            report,
            "ats_report.txt",
            "text/plain",
            use_container_width=True
        )

        st.markdown('---')

        # TAILORED RESUME
        st.markdown('### Generate AI-Tailored Resume', unsafe_allow_html=True)
        st.info("💡 Your tailored resume will have your missing keywords naturally added and be reordered by job relevance.")
        
        if st.button("✨ Generate Tailored Resume", use_container_width=True):
            with st.spinner("Creating your tailored resume..."):
                tailored = generate_tailored_resume(
                    st.session_state.resume_text,
                    st.session_state.jd_text,
                    matched,
                    missing
                )
            
            if tailored:
                st.success("✅ Resume tailored successfully!")
                
                # Display tailored resume
                st.markdown('### Your Tailored Resume')
                st.text_area("", value=tailored, height=400, disabled=True, key="tailored_display")
                
                # Download options
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        "⬇️ Download as Text",
                        tailored,
                        "tailored_resume.txt",
                        "text/plain",
                        use_container_width=True
                    )
                
                with col2:
                    # Create Word document
                    doc_io = create_word_document(tailored)
                    if doc_io:
                        st.download_button(
                            "⬇️ Download as Word",
                            doc_io,
                            "tailored_resume.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

        st.markdown('---')

        # NEXT STEPS
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("← Back", use_container_width=True):
                st.session_state.step = 2
                st.rerun()
        
        with col2:
            if st.button("🔄 Tailor Another", use_container_width=True):
                st.session_state.step = 0
                st.rerun()
        
        with col3:
            if st.button("⭐ Upgrade for More", use_container_width=True):
                st.session_state.step = 4
                st.rerun()

# ─────────────────────────────────────────────
# STEP 4: PRICING
# ─────────────────────────────────────────────
elif st.session_state.step == 4:
    st.markdown('## Choose Your Plan', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.markdown("""
        <div class="plan-card">
        <div class="plan-name">⭐ Pro</div>
        <div class="plan-price">$5</div>
        <div class="plan-period">one-time</div>
        <ul class="plan-features">
        <li>Unlimited rewrites</li>
        <li>Interview prep kit</li>
        <li>Cover letter</li>
        <li>Email support</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="plan-card featured">
        <div class="plan-name">👑 Premium</div>
        <div class="plan-price">$12</div>
        <div class="plan-period">per month</div>
        <ul class="plan-features">
        <li>Everything in Pro</li>
        <li>LinkedIn rewrite</li>
        <li>Job matching</li>
        <li>1-on-1 chat</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 3
            st.rerun()
    
    with col2:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_BASIC}" target="_blank"><button style="width:100%;padding:12px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:8px;cursor:pointer;font-weight:700;font-size:14px">Pay Pro $5</button></a>', unsafe_allow_html=True)
    
    with col3:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_PRO}" target="_blank"><button style="width:100%;padding:12px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:8px;cursor:pointer;font-weight:700;font-size:14px">Pay Premium $12</button></a>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown('''<div class="footer">
<strong>⚡ ResumeReflect</strong><br>
Land more interviews with AI-powered resume tailoring.<br>
<small>Your resume is not stored. Privacy-first, built for job seekers.</small><br><br>
<small>© 2026 ResumeReflect | Made in India 🇮🇳</small>
</div>''', unsafe_allow_html=True)
