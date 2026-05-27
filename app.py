import streamlit as st
import groq
import pdfplumber
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import json
from datetime import datetime
import requests

GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
RAZORPAY_BASIC = st.secrets.get("RAZORPAY_BASIC", "#")
RAZORPAY_PRO = st.secrets.get("RAZORPAY_PRO", "#")
SHEET_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzN99UwHn4Bt1mJ4MMS5ZSV-cysoTC_ac6d6oMNkWB_JAGb1i2vqBX3RmrCDqIsla3G/exec"

st.set_page_config(page_title="ResumeReflect - AI Resume Tailor", page_icon="⚡", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""<style>@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');:root{--bg:#f8fafc;--surface:#ffffff;--border:#e2e8f0;--text:#0f172a;--text-muted:#64748b;--accent:#06b6d4;--success:#10b981;--warning:#f59e0b;--danger:#ef4444}html,body,.stApp{background:var(--bg)!important;font-family:'DM Sans',sans-serif!important}#MainMenu,footer,header{visibility:hidden}.block-container{padding:2rem 2.5rem!important;max-width:1200px!important}h1,h2,h3{font-family:'Syne',sans-serif!important}h1{font-size:2.5rem!important;font-weight:800!important}h2{font-size:1.8rem!important;font-weight:700!important;margin:2rem 0 1.5rem!important}.stTabs [data-baseweb="tab-list"]{border-bottom:2px solid var(--border)!important;gap:4rem!important;padding-bottom:1rem!important}.stTabs [data-baseweb="tab"]{padding:1rem 2rem!important;border-bottom:3px solid transparent!important;color:var(--text-muted)!important;font-weight:600!important;font-size:15px!important}.stTabs [data-baseweb="tab"][aria-selected="true"]{border-bottom-color:var(--accent)!important;color:var(--accent)!important}.stButton>button{background:linear-gradient(135deg,#06b6d4,#0f766e)!important;color:white!important;border:none!important;border-radius:10px!important;padding:12px 24px!important;font-family:'Syne',sans-serif!important;font-weight:700!important;width:100%!important;height:44px!important}.stButton>button:hover{transform:translateY(-2px)!important;box-shadow:0 8px 20px rgba(6,182,212,0.3)!important}.stTextInput input,.stTextArea textarea{border:1.5px solid var(--border)!important;border-radius:10px!important;padding:12px!important;font-size:14px!important}.stTextInput input:focus,.stTextArea textarea:focus{border-color:var(--accent)!important;box-shadow:0 0 0 3px rgba(6,182,212,0.1)!important}.metric-card{background:white;border:1px solid var(--border);border-radius:12px;padding:1.5rem;margin-bottom:1rem}.ats-score-box{text-align:center;padding:2rem;border-radius:12px;margin:2rem 0}.ats-score-box.high{background:rgba(16,185,129,0.1);border:2px solid #10b981}.ats-score-box.medium{background:rgba(245,158,11,0.1);border:2px solid #f59e0b}.ats-score-box.low{background:rgba(239,68,68,0.1);border:2px solid #ef4444}.ats-number{font-size:4rem;font-weight:800;margin-bottom:0.5rem}.ats-label{font-size:1.2rem;font-weight:600}.keyword-section{background:white;border:1px solid var(--border);border-radius:12px;padding:2rem;margin:1.5rem 0}.keyword-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:1rem;margin:1rem 0}.keyword-item{padding:1rem;border-radius:8px;font-size:13px;font-weight:600;text-align:center}.keyword-matched{background:rgba(16,185,129,0.15);color:#047857;border:1px solid #10b981}.keyword-missing{background:rgba(239,68,68,0.15);color:#991b1b;border:1px solid #ef4444}.pricing-cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:2rem;margin:2rem 0}.plan-card{background:white;border:2px solid var(--border);border-radius:12px;padding:2rem;text-align:center}.plan-card.featured{border-color:var(--accent);background:linear-gradient(135deg,rgba(6,182,212,0.05),rgba(15,118,110,0.05))}.plan-price{font-size:2.5rem;font-weight:800;margin:1rem 0}.plan-features{list-style:none;padding:0;text-align:left}.plan-features li{padding:0.5rem 0;display:flex;gap:8px}.plan-features li:before{content:"✓";font-weight:800;color:var(--success)}.divider{height:1px;background:linear-gradient(90deg,transparent,var(--border),transparent);margin:2rem 0}</style>""", unsafe_allow_html=True)

if "step" not in st.session_state: st.session_state.step = 0
if "resume_text" not in st.session_state: st.session_state.resume_text = ""
if "jd_text" not in st.session_state: st.session_state.jd_text = ""
if "email" not in st.session_state: st.session_state.email = ""
if "analysis" not in st.session_state: st.session_state.analysis = None

def extract_pdf(f):
    try:
        with pdfplumber.open(f) as pdf: return "\n".join([p.extract_text() or "" for p in pdf.pages])
    except: return None

def extract_docx(f):
    try: return "\n".join([p.text for p in Document(f).paragraphs])
    except: return None

def analyze_now(resume, jd):
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": f"""Analyze resume vs job. Return ONLY JSON no markdown:
{{"ats_score": 75, "matched_keywords": ["Python", "Data Analysis", "SQL"], "missing_keywords": ["Machine Learning", "AWS", "Docker"], "strengths": ["Strong technical background", "Good communication"], "improvements": ["Add cloud certifications", "Highlight team leadership"]}}
JOB: {jd[:800]}
RESUME: {resume[:800]}"""}],
            temperature=0.2,
            max_tokens=400
        )
        text = resp.choices[0].message.content.strip()
        if "```" in text: text = text.split("```")[1] if "```json" not in text else text.split("```json")[1].split("```")[0]
        return json.loads(text)
    except Exception as e:
        st.error(f"Analysis failed: {str(e)}")
        return None

def gen_resume_now(resume, jd):
    try:
        client = groq.Groq(api_key=GROQ_API_KEY)
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": f"""Rewrite resume professionally for this job. Keep real info. Add missing keywords naturally. Format professionally.
JOB: {jd[:600]}
RESUME: {resume[:600]}
Return professional resume only, no explanations."""}],
            temperature=0.5,
            max_tokens=1200
        )
        return resp.choices[0].message.content
    except: return None

def make_word(title, data_dict):
    try:
        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)
        
        t = doc.add_paragraph(title)
        t.runs[0].font.size = Pt(16)
        t.runs[0].font.bold = True
        t.runs[0].font.color.rgb = RGBColor(15, 23, 42)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for k, v in data_dict.items():
            h = doc.add_paragraph(str(k).upper())
            h.runs[0].font.bold = True
            h.runs[0].font.size = Pt(12)
            h.paragraph_format.space_before = Pt(6)
            h.paragraph_format.space_after = Pt(3)
            
            if isinstance(v, list):
                for item in v:
                    p = doc.add_paragraph(str(item), style='List Bullet')
                    p.paragraph_format.space_after = Pt(2)
            else:
                p = doc.add_paragraph(str(v))
                p.paragraph_format.space_after = Pt(6)
        
        b = io.BytesIO()
        doc.save(b)
        b.seek(0)
        return b
    except: return None

def log_email(t, e):
    try: requests.post(SHEET_SCRIPT_URL, json={"type": t, "email": e, "timestamp": datetime.now().isoformat()}, timeout=3)
    except: pass

if st.session_state.step == 0:
    st.markdown('<h1 style="text-align:center;margin:3rem 0">⚡ Land More Interview Calls</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:center;font-size:1.1rem;color:#64748b;margin-bottom:3rem">Get past ATS filters. See what keywords you need. Get a tailored resume.</p>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('## How It Works', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="large")
    with c1: st.markdown('<div class="metric-card"><h3>📄 Upload Resume</h3><p>PDF, DOCX or paste</p></div>', unsafe_allow_html=True)
    with c2: st.markdown('<div class="metric-card"><h3>🎯 Add Job</h3><p>Paste job posting</p></div>', unsafe_allow_html=True)
    with c3: st.markdown('<div class="metric-card"><h3>⚡ Get Results</h3><p>Professional report</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('## Pricing', unsafe_allow_html=True)
    st.markdown('<div class="pricing-cols"><div class="plan-card"><h3>Free</h3><div class="plan-price">$0</div><ul class="plan-features"><li>ATS Analysis</li><li>Keyword Report</li><li>1 Resume</li><li>Download</li></ul></div><div class="plan-card featured"><h3>Pro</h3><div class="plan-price">$5</div><ul class="plan-features"><li>All Free</li><li>Unlimited</li><li>Interview Kit</li><li>Priority</li></ul></div><div class="plan-card"><h3>Premium</h3><div class="plan-price">$12/mo</div><ul class="plan-features"><li>All Pro</li><li>LinkedIn</li><li>Job Match</li><li>Support</li></ul></div></div>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    if st.button("✨ Start Free Analysis", use_container_width=True, key="start"): st.session_state.step = 1; st.rerun()

elif st.session_state.step == 1:
    st.markdown('## Upload Your Resume', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 2], gap="large")
    with c1:
        st.markdown('**Email**')
        email = st.text_input("", placeholder="your@email.com", label_visibility="collapsed", key="e1")
    with c2:
        st.markdown('**Resume**')
        tab1, tab2 = st.tabs(["📁 Upload", "📝 Paste"])
        with tab1:
            u = st.file_uploader("", type=["pdf", "docx"], label_visibility="collapsed", key="f1")
            if u:
                text = extract_pdf(u) if "pdf" in u.type else extract_docx(u)
                if text: st.session_state.resume_text = text; st.success("✅")
        with tab2:
            p = st.text_area("", placeholder="Paste resume...", height=180, label_visibility="collapsed", key="r1")
            if p: st.session_state.resume_text = p
    
    c1, c2 = st.columns(2, gap="large")
    with c1:
        if st.button("← Back", use_container_width=True, key="b1"): st.session_state.step = 0; st.rerun()
    with c2:
        if st.button("Continue →", use_container_width=True, key="c1"):
            if not email or "@" not in email: st.error("Invalid email")
            elif not st.session_state.resume_text or len(st.session_state.resume_text) < 50: st.error("Paste resume")
            else: st.session_state.email = email; log_email("upload", email); st.session_state.step = 2; st.rerun()

elif st.session_state.step == 2:
    st.markdown('## Add Job Description', unsafe_allow_html=True)
    jd = st.text_area("Paste job posting", placeholder="Paste job details...", height=300, label_visibility="collapsed", key="j1")
    if jd: st.session_state.jd_text = jd
    c1, c2 = st.columns(2, gap="large")
    with c1:
        if st.button("← Back", use_container_width=True, key="b2"): st.session_state.step = 1; st.rerun()
    with c2:
        if st.button("Analyze →", use_container_width=True, key="a1"):
            if not st.session_state.jd_text or len(st.session_state.jd_text) < 50: st.error("Paste job description")
            else: st.session_state.step = 3; st.rerun()

elif st.session_state.step == 3:
    st.markdown('## Resume Analysis Results', unsafe_allow_html=True)
    
    if st.session_state.analysis is None:
        with st.spinner("🔍 Analyzing your resume..."):
            st.session_state.analysis = analyze_now(st.session_state.resume_text, st.session_state.jd_text)
    
    if st.session_state.analysis:
        a = st.session_state.analysis
        score = a.get("ats_score", 65)
        matched = a.get("matched_keywords", [])
        missing = a.get("missing_keywords", [])
        strengths = a.get("strengths", [])
        improvements = a.get("improvements", [])
        
        color_class = "high" if score >= 70 else ("medium" if score >= 50 else "low")
        status = "✓ Good Match" if score >= 70 else ("⚠ Needs Work" if score >= 50 else "✗ Low Match")
        
        st.markdown(f'<div class="ats-score-box {color_class}"><div class="ats-number">{score}%</div><div class="ats-label">{status}</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="keyword-section"><h3>✓ Matched Keywords</h3><div class="keyword-grid">', unsafe_allow_html=True)
            for kw in matched: st.markdown(f'<div class="keyword-item keyword-matched">{kw}</div>', unsafe_allow_html=True)
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with c2:
            st.markdown('<div class="keyword-section"><h3>⚠ Missing Keywords</h3><div class="keyword-grid">', unsafe_allow_html=True)
            for kw in missing: st.markdown(f'<div class="keyword-item keyword-missing">{kw}</div>', unsafe_allow_html=True)
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="metric-card"><h3>💪 Strengths</h3>', unsafe_allow_html=True)
            for s in strengths: st.markdown(f"✓ {s}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with c2:
            st.markdown('<div class="metric-card"><h3>🎯 Improvements</h3>', unsafe_allow_html=True)
            for imp in improvements: st.markdown(f"→ {imp}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('## Download ATS Report', unsafe_allow_html=True)
        
        ats_doc = make_word("ATS COMPATIBILITY REPORT", {"ATS Score": f"{score}%", "Matched": matched, "Missing": missing, "Strengths": strengths, "Improvements": improvements})
        if ats_doc:
            st.download_button("📊 Download Report", ats_doc, "ATS_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('## Generate Tailored Resume', unsafe_allow_html=True)
        
        if st.button("✨ Create Professional Resume", use_container_width=True, key="gen"):
            with st.spinner("Creating your resume..."):
                tailored = gen_resume_now(st.session_state.resume_text, st.session_state.jd_text)
            
            if tailored:
                st.markdown('### Your Resume', unsafe_allow_html=True)
                st.text_area("", value=tailored, height=350, disabled=True, label_visibility="collapsed", key="preview")
                
                c1, c2, c3 = st.columns(3, gap="large")
                with c1: st.download_button("⬇ Text", tailored, "resume.txt", "text/plain", use_container_width=True, key="txt")
                with c2:
                    doc = make_word("PROFESSIONAL RESUME", {"Resume": tailored})
                    if doc: st.download_button("⬇ Word", doc, "resume.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="docx")
                with c3: st.info("Ready to apply!")
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3, gap="large")
        with c1:
            if st.button("← Back", use_container_width=True, key="b3"): st.session_state.step = 2; st.rerun()
        with c2:
            if st.button("🔄 Another", use_container_width=True, key="again"): st.session_state.step = 1; st.session_state.resume_text = ""; st.session_state.jd_text = ""; st.session_state.analysis = None; st.rerun()
        with c3:
            if st.button("⭐ Upgrade", use_container_width=True, key="upg"): st.session_state.step = 4; st.rerun()

elif st.session_state.step == 4:
    st.markdown('## Upgrade Your Plan', unsafe_allow_html=True)
    st.markdown('<div class="pricing-cols"><div class="plan-card"><h3>Pro</h3><div class="plan-price">$5</div><ul class="plan-features"><li>Unlimited Rewrites</li><li>Interview Kit</li><li>One-time Payment</li><li>Email Support</li></ul></div><div class="plan-card featured"><h3>Premium</h3><div class="plan-price">$12/mo</div><ul class="plan-features"><li>Everything Pro</li><li>LinkedIn Rewrite</li><li>Job Matching</li><li>Priority Support</li></ul></div></div>', unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        if st.button("← Back", use_container_width=True, key="b4"): st.session_state.step = 3; st.rerun()
    with c2:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_BASIC}" target="_blank"><button style="width:100%;height:44px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:10px;cursor:pointer;font-weight:700;font-family:Syne,sans-serif;font-size:14px">Pay $5</button></a>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<a href="https://buy.razorpay.com/{RAZORPAY_PRO}" target="_blank"><button style="width:100%;height:44px;background:linear-gradient(135deg,#06b6d4,#0f766e);color:white;border:none;border-radius:10px;cursor:pointer;font-weight:700;font-family:Syne,sans-serif;font-size:14px">Pay $12</button></a>', unsafe_allow_html=True)

st.markdown('<div class="divider"></div><div style="text-align:center;padding:2rem 0;color:#64748b;font-size:13px"><strong>⚡ ResumeReflect</strong><br>AI Resume Optimization<br><small>Secure & Private</small></div>', unsafe_allow_html=True)
