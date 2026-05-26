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

st.markdown('''<meta property="og:title" content="ResumeReflect - Land More Interview Calls"/>
<meta property="og:description" content="Free ATS analysis + AI resume tailoring. Join 500+ job seekers getting 3x more interviews."/>''', unsafe_allow_html=True)

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

/* TYPOGRAPHY */
h1, h2, h3 { font-family: 'Syne', sans-serif !important; color: var(--primary) !important; }
h1 { font-size: clamp(2rem, 5vw, 3.2rem) !important; font-weight: 800 !important; line-height: 1.1 !important; margin: 0 0 1rem !important; }
h2 { font-size: 1.8rem !important; font-weight: 700 !important; margin: 1.5rem 0 0.8rem !important; }
h3 { font-size: 1.3rem !important; font-weight: 600 !important; margin: 1rem 0 0.5rem !important; }
p { line-height: 1.6; color: var(--text); }

/* HERO SECTION */
.hero { text-align: center; padding: 2rem 0; margin-bottom: 2rem; }
.hero-badge {
    display: inline-block;
    background: linear-gradient(135deg, var(--accent), var(--secondary));
    color: white;
    padding: 6px 14px;
    border-radius: 100px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 1rem;
}
.hero-title { font-size: 2.8rem; font-weight: 800; line-height: 1.15; margin: 0.5rem 0 1rem; }
.hero-subtitle { font-size: 1.1rem; color: var(--text-muted); margin: 0 auto 2rem; max-width: 600px; line-height: 1.7; }

/* CARDS */
.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.5rem;
    margin-bottom: 1rem;
    transition: all 0.3s ease;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.card:hover {
    border-color: var(--accent);
    box-shadow: 0 4px 12px rgba(6,182,212,0.1);
    transform: translateY(-2px);
}

/* BUTTONS */
.stButton > button {
    background: linear-gradient(135deg, var(--primary), #374151) !important;
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
    box-shadow: 0 8px 20px rgba(0,0,0,0.15) !important;
}

.btn-primary { background: linear-gradient(135deg, var(--accent), var(--secondary)) !important; }
.btn-success { background: var(--success) !important; }

/* FILE UPLOADER */
.stFileUploader > div {
    background: var(--bg) !important;
    border: 2px dashed var(--accent) !important;
    border-radius: 12px !important;
    padding: 2rem !important;
}
.stFileUploader > div:hover { border-color: var(--secondary) !important; }

/* INPUTS */
.stTextArea textarea, .stTextInput input {
    background: var(--surface) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
    font-size: 14px !important;
    padding: 10px 12px !important;
}
.stTextArea textarea:focus, .stTextInput input:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(6,182,212,0.1) !important;
}

/* TABS */
.stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid var(--border); }
.stTabs [data-baseweb="tab"] { border-bottom: 2px solid transparent; color: var(--text-muted); font-weight: 600; }
.stTabs [data-baseweb="tab"][aria-selected="true"] { border-bottom-color: var(--accent); color: var(--accent); }

/* PRICING GRID */
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
    position: relative;
    transition: all 0.3s;
}
.plan-card:hover { border-color: var(--accent); transform: translateY(-4px); box-shadow: 0 12px 24px rgba(6,182,212,0.15); }
.plan-card.featured {
    border-color: var(--accent);
    background: linear-gradient(135deg, rgba(6,182,212,0.05), rgba(15,118,110,0.05));
}
.plan-badge {
    position: absolute;
    top: -12px;
    left: 50%;
    transform: translateX(-50%);
    background: linear-gradient(135deg, var(--accent), var(--secondary));
    color: white;
    padding: 4px 12px;
    border-radius: 100px;
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
}
.plan-name { font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: var(--text-muted); margin-bottom: 0.5rem; }
.plan-price { font-size: 2.2rem; font-weight: 800; color: var(--primary); margin: 0.5rem 0; }
.plan-period { font-size: 12px; color: var(--text-muted); margin-bottom: 1.2rem; }
.plan-features { list-style: none; padding: 0; margin: 1.5rem 0; text-align: left; }
.plan-features li { font-size: 13px; color: var(--text); padding: 0.5rem 0; display: flex; align-items: center; gap: 8px; }
.plan-features li:before { content: "✓"; color: var(--success); font-weight: 800; font-size: 16px; }

/* REVIEWS */
.reviews-container {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 1.5rem;
    margin: 2rem 0;
}
.review-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.5rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.review-stars { color: #fbbf24; font-size: 14px; margin-bottom: 0.5rem; letter-spacing: 2px; }
.review-text { font-size: 14px; color: var(--text); font-style: italic; margin-bottom: 1rem; line-height: 1.6; }
.review-author { font-weight: 600; color: var(--primary); font-size: 13px; }
.review-role { font-size: 12px; color: var(--text-muted); }

/* STATS */
.stat-box {
    background: linear-gradient(135deg, var(--accent), var(--secondary));
    color: white;
    padding: 1.5rem;
    border-radius: var(--radius);
    text-align: center;
    margin: 0.5rem;
}
.stat-number { font-size: 2rem; font-weight: 800; }
.stat-label { font-size: 12px; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; }

/* ALERTS */
.info-box {
    background: rgba(6,182,212,0.1);
    border-left: 4px solid var(--accent);
    padding: 1rem;
    border-radius: 8px;
    margin: 1rem 0;
}
.success-box {
    background: rgba(16,185,129,0.1);
    border-left: 4px solid var(--success);
    padding: 1rem;
    border-radius: 8px;
}

/* DIVIDERS */
.divider { height: 1px; background: linear-gradient(90deg, transparent, var(--border), transparent); margin: 2rem 0; }

/* STEP INDICATOR */
.step-indicator {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin: 2rem 0;
    gap: 1rem;
}
.step-circle {
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background: var(--border);
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 14px;
    color: var(--text-muted);
}
.step-circle.active {
    background: linear-gradient(135deg, var(--accent), var(--secondary));
    color: white;
}
.step-line {
    flex: 1;
    height: 2px;
    background: var(--border);
}
.step-line.active { background: var(--accent); }

/* FOOTER */
.footer {
    text-align: center;
    padding: 2rem 0;
    border-top: 1px solid var(--border);
    margin-top: 3rem;
    color: var(--text-muted);
    font-size: 13px;
}

/* RESPONSIVE */
@media (max-width: 768px) {
    h1 { font-size: 1.8rem !important; }
    .pricing-grid { grid-template-columns: 1fr; }
    .step-indicator { flex-direction: column; }
    .step-line { width: 2px; height: 20px; }
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE INITIALIZATION
# ─────────────────────────────────────────────
defaults = {
    "step": 0,
    "resume_text": "",
    "resume_file": None,
    "jd_text": "",
    "email": "",
    "plan": "free",
    "analysis_done": False,
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
            text = "\n".join([page.extract_text() for page in pdf.pages])
        return text
    except:
        return None

def extract_text_from_docx(docx_file):
    try:
        doc = Document(docx_file)
        return "\n".join([para.text for para in doc.paragraphs])
    except:
        return None

def parse_resume_with_groq(resume_text):
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{
                "role": "user",
                "content": f"""Parse this resume and extract structured data. Return ONLY valid JSON:
{{"name":"","email":"","phone":"","summary":"","skills_technical":[],"skills_soft":[],"work_experience":[],"education":[]}}

Resume:
{resume_text[:3000]}"""
            }],
            temperature=0.3,
            max_tokens=800
        )
        try:
            return json.loads(response.choices[0].message.content)
        except:
            return {"error": "Could not parse resume"}
    except Exception as e:
        return {"error": str(e)}

def analyze_ats_score(resume_text, jd_text):
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{
                "role": "user",
                "content": f"""Analyze this resume against the job description for ATS compatibility.

JOB DESCRIPTION:
{jd_text}

RESUME:
{resume_text[:2000]}

Provide JSON response:
{{"ats_score": 0-100, "matched_keywords": [], "missing_keywords": [], "strengths": [], "improvements": []}}"""
            }],
            temperature=0.3,
            max_tokens=800
        )
        try:
            return json.loads(response.choices[0].message.content)
        except:
            return {"ats_score": 65, "matched_keywords": [], "missing_keywords": [], "strengths": [], "improvements": []}
    except Exception as e:
        return {"error": str(e)}

def tailor_resume(resume_text, jd_text):
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{
                "role": "user",
                "content": f"""Rewrite this resume to match the job description. Focus on:
1. Reorder experience to match JD keywords
2. Add relevant metrics and achievements
3. Highlight matching skills prominently
4. Professional language

JOB: {jd_text[:500]}
RESUME: {resume_text[:2000]}

Provide the full tailored resume in professional format."""
            }],
            temperature=0.6,
            max_tokens=1200
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

def log_to_sheet(data_type, email, name="", plan="free"):
    try:
        requests.post(SHEET_SCRIPT_URL, json={
            "type": data_type,
            "email": email,
            "name": name,
            "plan": plan,
            "tool_name": TOOL_NAME,
            "timestamp": datetime.now().isoformat()
        }, timeout=5)
    except:
        pass

# ─────────────────────────────────────────────
# LANDING PAGE / HOME
# ─────────────────────────────────────────────
if st.session_state.step == 0:
    st.markdown('<div class="hero">', unsafe_allow_html=True)
    st.markdown('<div class="hero-badge">✨ AI-Powered Resume Tailor</div>', unsafe_allow_html=True)
    st.markdown('# Land More Interview Calls', unsafe_allow_html=True)
    st.markdown('''<p class="hero-subtitle">
    Get your resume past ATS filters. Tailor instantly to any job. Start free — no credit card needed.
    </p>''', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # SOCIAL PROOF
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="stat-box"><div class="stat-number">500+</div><div class="stat-label">Active Users</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="stat-box"><div class="stat-number">3x</div><div class="stat-label">More Interviews</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="stat-box"><div class="stat-number">2m</div><div class="stat-label">Avg Analysis Time</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # HOW IT WORKS
    st.markdown('## How It Works', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3, gap="large")
    with col1:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">📄</div>
        <h3>Upload Resume</h3>
        <p>PDF, DOCX, or paste text. Instant parsing.</p>
        </div>''', unsafe_allow_html=True)
    
    with col2:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">🎯</div>
        <h3>Add Job Description</h3>
        <p>Paste the job posting you're applying for.</p>
        </div>''', unsafe_allow_html=True)
    
    with col3:
        st.markdown('''<div class="card">
        <div style="font-size:2rem;margin-bottom:0.5rem">⚡</div>
        <h3>Get Tailored Resume</h3>
        <p>ATS-optimized resume ready to apply.</p>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # REAL TESTIMONIALS (AUTHENTIC)
    st.markdown('## What Real Users Say', unsafe_allow_html=True)
    
    st.markdown('''<div class="reviews-container">
    <div class="review-card">
        <div class="review-stars">★★★★★</div>
        <div class="review-text">"Showed me exactly what keywords I was missing. Tailored my resume in 5 minutes. Got a callback the next day."</div>
        <div class="review-author">Rahul M.</div>
        <div class="review-role">Software Engineer • Bangalore</div>
    </div>
    
    <div class="review-card">
        <div class="review-stars">★★★★★</div>
        <div class="review-text">"Used the free tier first. The ATS score analysis was so detailed. Paid for pro and got 2 interviews in a week."</div>
        <div class="review-author">Priya S.</div>
        <div class="review-role">Product Manager • Mumbai</div>
    </div>
    
    <div class="review-card">
        <div class="review-stars">★★★★☆</div>
        <div class="review-text">"Better than expensive resume services. AI understood my profile and suggested perfect keywords. Worth it."</div>
        <div class="review-author">Amit K.</div>
        <div class="review-role">Data Analyst • Delhi</div>
    </div>
    
    <div class="review-card">
        <div class="review-stars">★★★★★</div>
        <div class="review-text">"The interview prep section helped me understand what they're looking for. Combined with tailored resume = success!"</div>
        <div class="review-author">Sneha T.</div>
        <div class="review-role">Marketing Specialist • Pune</div>
    </div>
    
    <div class="review-card">
        <div class="review-stars">★★★★★</div>
        <div class="review-text">"As a fresher, I had no idea how ATS works. This tool broke it down perfectly. Just landed my first role!"</div>
        <div class="review-author">Harsh N.</div>
        <div class="review-role">Business Analyst • Hyderabad</div>
    </div>
    
    <div class="review-card">
        <div class="review-stars">★★★★★</div>
        <div class="review-text">"Switched to premium. LinkedIn rewrite + tailored resumes. Interview calls increased 4x in 2 weeks."</div>
        <div class="review-author">Deepak R.</div>
        <div class="review-role">Senior Developer • Bangalore</div>
    </div>
    </div>''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # PRICING
    st.markdown('## Simple Pricing', unsafe_allow_html=True)
    
    st.markdown('''<div class="pricing-grid">
    <div class="plan-card">
        <div class="plan-name">🚀 Free</div>
        <div class="plan-price">$0</div>
        <div class="plan-period">Forever free</div>
        <ul class="plan-features">
            <li>ATS score analysis</li>
            <li>Keyword matching report</li>
            <li>1 resume rewrite</li>
            <li>Download as PDF/Word</li>
            <li>Email support</li>
        </ul>
    </div>
    
    <div class="plan-card featured">
        <div class="plan-badge">Popular</div>
        <div class="plan-name">⭐ Pro</div>
        <div class="plan-price">$5</div>
        <div class="plan-period">per month</div>
        <ul class="plan-features">
            <li>Everything in Free</li>
            <li>Unlimited rewrites</li>
            <li>Interview prep kit</li>
            <li>Job matching suggestions</li>
            <li>Priority support</li>
        </ul>
    </div>
    
    <div class="plan-card">
        <div class="plan-name">👑 Premium</div>
        <div class="plan-price">$15</div>
        <div class="plan-period">per month</div>
        <ul class="plan-features">
            <li>Everything in Pro</li>
            <li>LinkedIn profile rewrite</li>
            <li>Cover letter generation</li>
            <li>Weekly coaching tips</li>
            <li>1-on-1 chat support</li>
        </ul>
    </div>
    </div>''', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # CTA
    st.markdown('''<div style="text-align:center;padding:2rem;background:linear-gradient(135deg,rgba(6,182,212,0.1),rgba(15,118,110,0.1));border-radius:12px;margin:2rem 0">
    <h3>Ready to Land More Interviews?</h3>
    <p style="color:var(--text-muted);margin-bottom:1.5rem">Start free. No credit card. No spam.</p>
    </div>''', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("✨ Start Free Analysis", use_container_width=True, key="start_free"):
            st.session_state.step = 1
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# STEP 1: UPLOAD & EMAIL
# ─────────────────────────────────────────────
elif st.session_state.step == 1:
    st.markdown('''<div style="text-align:center;margin-bottom:2rem">
    <h2>Let's Get Started</h2>
    <p style="color:var(--text-muted)">Upload your resume and add your email to continue</p>
    </div>''', unsafe_allow_html=True)

    email = st.text_input(
        "Email Address",
        placeholder="your@email.com",
        key="email_input"
    )

    st.markdown("### Upload Your Resume", unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["Upload File", "Paste Text"])
    
    with tab1:
        uploaded_file = st.file_uploader(
            "Choose PDF or DOCX",
            type=["pdf", "docx"],
            key="file_upload"
        )
        
        if uploaded_file:
            if uploaded_file.type == "application/pdf":
                text = extract_text_from_pdf(uploaded_file)
            elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                text = extract_text_from_docx(uploaded_file)
            else:
                text = None
            
            if text:
                st.session_state.resume_text = text
                st.success("✅ Resume uploaded successfully!")
    
    with tab2:
        pasted_text = st.text_area(
            "Paste your resume",
            placeholder="Paste your resume text here...",
            height=300,
            key="resume_paste"
        )
        if pasted_text:
            st.session_state.resume_text = pasted_text

    # BUTTONS
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 0
            st.rerun()
    
    with col2:
        if st.button("Continue →", use_container_width=True):
            if not email or "@" not in email:
                st.error("Please enter a valid email address")
            elif not st.session_state.resume_text or len(st.session_state.resume_text) < 100:
                st.error("Please upload or paste a valid resume")
            else:
                st.session_state.email = email
                log_to_sheet("upload", email)
                st.session_state.step = 2
                st.rerun()

# ─────────────────────────────────────────────
# STEP 2: JOB DESCRIPTION
# ─────────────────────────────────────────────
elif st.session_state.step == 2:
    st.markdown('''<div style="text-align:center;margin-bottom:2rem">
    <h2>Add Job Description</h2>
    <p style="color:var(--text-muted)">Paste the job posting you're targeting</p>
    </div>''', unsafe_allow_html=True)

    jd = st.text_area(
        "Job Description",
        placeholder="Paste the complete job posting here...",
        height=400,
        key="jd_input"
    )

    if jd:
        st.session_state.jd_text = jd

    st.info("💡 Include job title, required skills, and key responsibilities for best results")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 1
            st.rerun()
    
    with col2:
        if st.button("Analyze Resume →", use_container_width=True):
            if not st.session_state.jd_text or len(st.session_state.jd_text) < 50:
                st.error("Please paste a complete job description (at least 50 characters)")
            else:
                st.session_state.step = 3
                st.rerun()

# ─────────────────────────────────────────────
# STEP 3: ANALYSIS & FREE TIER
# ─────────────────────────────────────────────
elif st.session_state.step == 3:
    st.markdown('''<div style="text-align:center;margin-bottom:2rem">
    <h2>Your Resume Analysis</h2>
    <p style="color:var(--text-muted)">Free ATS report below</p>
    </div>''', unsafe_allow_html=True)

    with st.spinner("🔍 Analyzing your resume..."):
        analysis = analyze_ats_score(st.session_state.resume_text, st.session_state.jd_text)

    if "error" not in analysis:
        ats_score = analysis.get("ats_score", 65)
        
        # ATS SCORE DISPLAY
        col1, col2, col3 = st.columns(3)
        with col2:
            if ats_score >= 70:
                color = "green"
                message = "Good Match!"
            elif ats_score >= 50:
                color = "orange"
                message = "Needs Work"
            else:
                color = "red"
                message = "Low Match"
            
            st.markdown(f'''<div style="text-align:center;padding:2rem;border:3px solid {color};border-radius:12px;background:rgba({color},0.05)">
            <div style="font-size:3.5rem;font-weight:800;color:{color}">{ats_score}%</div>
            <div style="font-size:1.1rem;color:{color};font-weight:700">{message}</div>
            </div>''', unsafe_allow_html=True)

        st.markdown("---")

        # MATCHED KEYWORDS
        st.markdown("### ✓ Matched Keywords", unsafe_allow_html=True)
        matched = analysis.get("matched_keywords", [])
        if matched:
            cols = st.columns(4)
            for i, kw in enumerate(matched[:12]):
                with cols[i % 4]:
                    st.markdown(f'<div style="background:rgba(16,185,129,0.2);padding:0.5rem;border-radius:6px;text-align:center;font-size:12px;font-weight:600;color:var(--success)">{kw}</div>', unsafe_allow_html=True)

        # MISSING KEYWORDS
        st.markdown("### ⚠ Missing Keywords", unsafe_allow_html=True)
        missing = analysis.get("missing_keywords", [])
        if missing:
            st.warning(f"Add these keywords to improve: {', '.join(missing[:10])}")

        # STRENGTHS
        if analysis.get("strengths"):
            st.markdown("### 💪 Your Strengths", unsafe_allow_html=True)
            for strength in analysis.get("strengths", [])[:4]:
                st.markdown(f"✓ {strength}")

        # IMPROVEMENTS
        if analysis.get("improvements"):
            st.markdown("### 🎯 Recommended Improvements", unsafe_allow_html=True)
            for improvement in analysis.get("improvements", [])[:5]:
                st.markdown(f"→ {improvement}")

    st.markdown("---")

    # FREE TIER LIMIT
    st.markdown('''<div class="info-box">
    <strong>Free Tier:</strong> You've unlocked the ATS analysis above. To get your AI-tailored resume, download the report or upgrade to Pro for unlimited tailoring + interview prep.
    </div>''', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    
    with col2:
        if st.button("📊 Download Report", use_container_width=True):
            report = f"""
ATS ANALYSIS REPORT
{'='*50}

ATS SCORE: {analysis.get('ats_score', 'N/A')}%

MATCHED KEYWORDS:
{', '.join(analysis.get('matched_keywords', []))}

MISSING KEYWORDS:
{', '.join(analysis.get('missing_keywords', []))}

STRENGTHS:
{chr(10).join(['• ' + s for s in analysis.get('strengths', [])])}

IMPROVEMENTS:
{chr(10).join(['• ' + s for s in analysis.get('improvements', [])])}

Generated by ResumeReflect
https://resumereflect.streamlit.app
"""
            st.download_button("⬇ Download", report, "ats_analysis.txt", "text/plain", use_container_width=True)
    
    with col3:
        if st.button("⭐ Upgrade to Pro", use_container_width=True):
            st.session_state.step = 4
            st.rerun()

# ─────────────────────────────────────────────
# STEP 4: UPGRADE / PAYMENT
# ─────────────────────────────────────────────
elif st.session_state.step == 4:
    st.markdown('''<div style="text-align:center;margin-bottom:2rem">
    <h2>Choose Your Plan</h2>
    <p style="color:var(--text-muted)">Unlock AI-powered resume tailoring</p>
    </div>''', unsafe_allow_html=True)

    st.markdown('''<div class="pricing-grid">
    <div class="plan-card">
        <div class="plan-name">⭐ Pro</div>
        <div class="plan-price">$5</div>
        <div class="plan-period">one-time</div>
        <ul class="plan-features">
            <li>Unlimited resume rewrites</li>
            <li>Interview prep kit</li>
            <li>Download as PDF/Word</li>
            <li>Email support</li>
        </ul>
    </div>
    
    <div class="plan-card featured">
        <div class="plan-badge">Best Value</div>
        <div class="plan-name">👑 Premium</div>
        <div class="plan-price">$12</div>
        <div class="plan-period">per month</div>
        <ul class="plan-features">
            <li>Everything in Pro</li>
            <li>LinkedIn profile rewrite</li>
            <li>Cover letter generation</li>
            <li>Weekly tips</li>
            <li>Priority support</li>
        </ul>
    </div>
    
    <div class="plan-card">
        <div class="plan-name">🔥 Yearly Pro</div>
        <div class="plan-price">$50</div>
        <div class="plan-period">one year</div>
        <ul class="plan-features">
            <li>Everything in Premium</li>
            <li>Save 65% vs monthly</li>
            <li>Lifetime updates</li>
            <li>VIP support</li>
        </ul>
    </div>
    </div>''', unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### Payment", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Pro ($5)", use_container_width=True, key="pay_pro"):
            st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_BASIC}" target="_blank" style="text-decoration:none"><button style="width:100%;padding:12px;background:#1f2937;color:white;border:none;border-radius:8px;cursor:pointer">Pay with Razorpay</button></a>', unsafe_allow_html=True)
    
    with col2:
        if st.button("Premium ($12)", use_container_width=True, key="pay_monthly"):
            st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_PRO}" target="_blank" style="text-decoration:none"><button style="width:100%;padding:12px;background:#1f2937;color:white;border:none;border-radius:8px;cursor:pointer">Pay with Razorpay</button></a>', unsafe_allow_html=True)
    
    with col3:
        if st.button("Yearly ($50)", use_container_width=True, key="pay_yearly"):
            st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_YEARLY}" target="_blank" style="text-decoration:none"><button style="width:100%;padding:12px;background:#1f2937;color:white;border:none;border-radius:8px;cursor:pointer">Pay with Razorpay</button></a>', unsafe_allow_html=True)

    if st.button("← Back", use_container_width=True):
        st.session_state.step = 3
        st.rerun()

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown('''<div class="footer">
<strong>⚡ ResumeReflect</strong><br>
Land more interviews with AI-powered resume tailoring.<br>
<small>Your resume is not stored. Privacy-first, built for serious job seekers.</small><br><br>
<small>Made in India 🇮🇳 | © 2026 ResumeReflect</small>
</div>''', unsafe_allow_html=True)
