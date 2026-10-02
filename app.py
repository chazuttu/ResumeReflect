import streamlit as st
import groq
import pdfplumber
import smtplib, ssl, html, traceback
import requests
from email.message import EmailMessage
import io, json, re, base64, uuid
import qrcode
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ───────────── CONFIG (put keys in .streamlit/secrets.toml) ─────────────
GROQ_API_KEY     = st.secrets.get("GROQ_API_KEY", "")
UPI_ID           = st.secrets.get("UPI_ID", "")        # your PhonePe UPI ID, e.g. name@ybl
UPI_NAME         = st.secrets.get("UPI_NAME", "ResumeReflect")
COUNSELLING_OPEN = st.secrets.get("COUNSELLING_OPEN", True)   # set false to hide the ₹499 card when you're full
TELEGRAM_TOKEN   = st.secrets.get("TELEGRAM_BOT_TOKEN", "")  # optional: instant phone alerts
TELEGRAM_CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID", "")
NOTIFY_EMAIL     = st.secrets.get("NOTIFY_EMAIL", "")   # where booking alerts are sent (your email)
SMTP_USER        = st.secrets.get("SMTP_USER", "")      # Gmail address used to send the alerts
SMTP_PASSWORD    = st.secrets.get("SMTP_PASSWORD", "")  # Gmail "app password" (not your normal password)

st.set_page_config(page_title="ResumeReflect – AI Resume Tailor", page_icon="⚡",
                   layout="centered", initial_sidebar_state="collapsed")

# ───────────── STYLE (responsive: phone + laptop) ─────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap');
html, body, .stApp { font-family:'Poppins',sans-serif !important; color:#f5f3ff;
  background: radial-gradient(1200px 600px at 10% -10%, #5b21b6 0%, transparent 60%),
              radial-gradient(900px 500px at 100% 0%, #be185d 0%, transparent 55%), #0f0a2a !important; }
#MainMenu, footer, header { visibility:hidden; }
.block-container { max-width:760px !important; padding:1.2rem 1rem 3rem !important; }
.hero { text-align:center; padding:2rem .5rem 1rem; }
.badge { display:inline-block; padding:6px 16px; border-radius:99px; font-size:12px; font-weight:600; letter-spacing:1px;
  background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.25); margin-bottom:1rem; }
.hero h1 { font-size:clamp(2rem,7vw,3.4rem); font-weight:800; line-height:1.1; margin:0 0 .8rem; color:#fff; }
.grad { background:linear-gradient(90deg,#22d3ee,#a78bfa,#f472b6); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; }
.hero p { color:#d8d4f2; font-size:clamp(.95rem,3.5vw,1.1rem); max-width:520px; margin:0 auto; line-height:1.6; }
.steps { display:flex; flex-wrap:wrap; gap:.6rem; justify-content:center; margin:1.2rem 0 1.6rem; }
.step { background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.15); border-radius:99px; padding:7px 14px; font-size:13px; font-weight:500; }
.lbl { font-weight:600; font-size:15px; margin:1rem 0 .4rem; color:#fff; }
[data-testid="stFileUploaderDropzone"] { background:#fff !important; border:2px dashed #a78bfa !important; border-radius:16px !important; }
[data-testid="stFileUploaderDropzone"] * { color:#1e1b4b !important; }
[data-testid="stFileUploaderDropzone"] button { background:#6d28d9 !important; color:#fff !important; border-radius:10px !important; border:none !important; }
.stTextArea textarea, .stTextInput input { background:#fff !important; color:#1e1b4b !important; border-radius:14px !important; border:2px solid transparent !important; font-size:15px !important; }
.stTextArea textarea:focus, .stTextInput input:focus { border-color:#a78bfa !important; }
.stButton>button, .stDownloadButton>button, .stLinkButton>a { width:100% !important; min-height:52px; border-radius:14px !important; border:none !important;
  background:linear-gradient(90deg,#7c3aed,#db2777) !important; color:#fff !important; font-weight:700 !important; font-size:16px !important;
  font-family:'Poppins',sans-serif !important; box-shadow:0 8px 24px rgba(124,58,237,.35); transition:transform .15s; text-decoration:none !important; }
.stButton>button:hover, .stDownloadButton>button:hover, .stLinkButton>a:hover { transform:translateY(-2px); color:#fff !important; }
.stButton>button p, .stDownloadButton>button p, .stLinkButton>a p { color:#fff !important; }
.card { background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.16); border-radius:20px; padding:1.3rem; margin:1.1rem 0; backdrop-filter:blur(8px); }
.card h3 { margin:0 0 .4rem; font-size:1.15rem; color:#fff; }
.card p { margin:0; color:#d8d4f2; font-size:14px; line-height:1.6; }
.ring { width:150px; height:150px; border-radius:50%; margin:.5rem auto; display:grid; place-items:center; }
.ring span { width:118px; height:118px; border-radius:50%; background:#1a1340; display:grid; place-items:center; font-size:2.2rem; font-weight:800; color:#fff; }
.center { text-align:center; }
.chip { display:inline-block; margin:4px; padding:6px 13px; border-radius:99px; font-size:13px; font-weight:500; background:rgba(251,191,36,.15); border:1px solid rgba(251,191,36,.5); color:#fde68a; }
.pay { background:linear-gradient(135deg,#7c3aed,#db2777); border-radius:20px; padding:1.4rem; text-align:center; margin:1.2rem 0 .6rem; }
.pay .big { font-size:1.6rem; font-weight:800; color:#fff; } .pay .sm { color:#fce7f3; font-size:14px; margin-top:.3rem; }
.foot { text-align:center; color:#b9b3dd; font-size:13px; margin-top:2.5rem; line-height:1.9; }
@media (max-width:640px){ .block-container{padding:.8rem .8rem 2.5rem !important;} .card{padding:1rem;} }
</style>
""", unsafe_allow_html=True)

# ───────────── HELPERS ─────────────
def read_pdf(file_bytes):
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text.strip()

def track(text):
    """Optional Telegram ping. Sends only anonymous counts, never resumes or emails."""
    if not (TELEGRAM_TOKEN and TELEGRAM_CHAT_ID): return
    try:
        requests.post(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
                      json={"chat_id": TELEGRAM_CHAT_ID, "text": f"ResumeReflect: {text}"}, timeout=5)
    except Exception:
        pass

def notify_owner(subject, body, reply_to="", attachment=None):
    try:
        msg = EmailMessage()
        msg["Subject"], msg["From"], msg["To"] = subject, SMTP_USER, NOTIFY_EMAIL
        if reply_to: msg["Reply-To"] = reply_to
        msg.set_content(body)
        if attachment:
            msg.add_attachment(attachment, maintype="application",
                               subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                               filename="tailored_resume.docx")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=20) as srv:
            srv.login(SMTP_USER, SMTP_PASSWORD)
            srv.send_message(msg)
        return True
    except Exception:
        return False

def call_groq(resume_text, jd_text):
    client = groq.Groq(api_key=GROQ_API_KEY)
    prompt = f"""
You are an expert ATS resume writer. Rewrite the candidate's resume so it MIRRORS the job description.
Return ONLY pure JSON (no markdown, no extra text) in exactly this shape:
{{
  "candidate_name": "", "email": "", "phone": "", "location": "", "linkedin": "",
  "jd_keywords": ["25-40 important skills, tools, technologies, qualifications and role terms from the JD, each 1-4 words"],
  "strong_points": ["5 short points where the candidate already fits the JD"],
  "improvement_tips": ["4 short, practical tips to close the gaps"],
  "summary": "2-3 sentences written in the JD's own language",
  "work_experience": [{{"title":"","company":"","dates":"","location":"","bullets":["",""]}}],
  "projects": [{{"name":"","bullets":[""]}}],
  "education": [{{"degree":"","institution":"","year":"","cgpa":""}}],
  "skills_technical": [], "skills_tools": [], "achievements": [], "certifications": []
}}
RULES:
- Use the JD's exact terminology, responsibilities and keywords throughout the summary, skills and every bullet, wherever the candidate's real background supports it. Reword existing experience in the JD's language; keep every existing skill.
- NEVER invent employers, titles, dates, degrees, certifications, tools or experience the resume does not show. Anything the JD asks for that the resume lacks must NOT appear in the resume (it will be reported as a missing skill).
- Keep company names, dates and institutions exactly as in the original.

RESUME:
{resume_text[:9000]}

JOB DESCRIPTION:
{jd_text[:6000]}
"""
    r = client.chat.completions.create(model="openai/gpt-oss-120b",
        messages=[{"role": "user", "content": prompt}], temperature=0.4, max_tokens=7000)
    return r.choices[0].message.content

def parse_json(text):
    text = "\n".join(l for l in text.strip().split("\n") if not l.strip().startswith("```"))
    s, e = text.find("{"), text.rfind("}") + 1
    return json.loads(text[s:e])

def score_resume(data):
    """Honest score: share of the JD's key terms that actually appear in the tailored resume."""
    parts = [data.get("summary", "")] + data.get("skills_technical", []) + data.get("skills_tools", []) \
          + data.get("achievements", []) + data.get("certifications", [])
    for j in data.get("work_experience", []): parts += [j.get("title", "")] + j.get("bullets", [])
    for p in data.get("projects", []): parts += [p.get("name", "")] + p.get("bullets", [])
    blob = " ".join(map(str, parts)).lower()
    kws = list(dict.fromkeys(k.strip() for k in data.get("jd_keywords", []) if k and k.strip()))
    found = [k for k in kws if k.lower() in blob]
    missing = [k for k in kws if k.lower() not in blob]
    score = int(round(100 * len(found) / len(kws))) if kws else 0
    data.update(match_score=max(5, min(score, 99)), ats_keywords_found=len(found),
                ats_keywords_missing=len(missing), missing_skills=missing)
    return data

# ───────────── UPI QR ─────────────
def upi_link(amount, ref):
    return f"upi://pay?pa={UPI_ID}&pn={UPI_NAME.replace(' ', '%20')}&am={amount}&cu=INR&tn={ref}"

def qr_b64(amount, ref):
    img = qrcode.make(upi_link(amount, ref), box_size=8, border=2)
    buf = io.BytesIO(); img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def upi_block(amount, ref):
    """QR + 'open UPI app' button + payment reference."""
    st.markdown(f"""<div class="center">
      <img src="data:image/png;base64,{qr_b64(amount, ref)}" style="width:210px;max-width:70%;background:#fff;border-radius:16px;padding:8px">
      <div style="margin:.7rem 0 .3rem"><a href="{upi_link(amount, ref)}" style="display:inline-block;padding:10px 20px;border-radius:12px;background:#fff;color:#4c1d95;font-weight:700;text-decoration:none">Open UPI app to pay ₹{amount}</a></div>
      <div style="color:#b9b3dd;font-size:12px">Payment reference: {ref} · Please don't edit the note</div></div>""", unsafe_allow_html=True)

# ───────────── STATE ─────────────
for k, v in {"data": None, "booked": False, "ref499": "RR" + uuid.uuid4().hex[:8].upper()}.items():
    st.session_state.setdefault(k, v)

# ───────────── HERO ─────────────
st.markdown("""
<div class="hero">
  <div class="badge">⚡ AI POWERED · ATS FRIENDLY</div>
  <h1>Match Any Job<br><span class="grad">in 30 Seconds</span></h1>
  <p>Upload your resume, paste the job description, and get a resume written in the employer's own language.</p>
</div>
<div class="steps"><span class="step">1 · Upload PDF</span><span class="step">2 · Paste job</span><span class="step">3 · Download</span></div>
""", unsafe_allow_html=True)

# ───────────── INPUT ─────────────
if not st.session_state.data:
    st.markdown('<div class="lbl">📄 Your resume (PDF)</div>', unsafe_allow_html=True)
    resume_file = st.file_uploader("Resume", type=["pdf"], label_visibility="collapsed")
    st.markdown('<div class="lbl">💼 Job description</div>', unsafe_allow_html=True)
    jd = st.text_area("Job description", height=200, placeholder="Paste the full job description here…", label_visibility="collapsed")
    st.write("")
    if st.button("⚡ Tailor My Resume"):
        if not resume_file:
            st.error("Please upload your resume (PDF).")
        elif resume_file.size > 5 * 1024 * 1024:
            st.error("This file is too large. Please upload a PDF under 5 MB.")
        elif len(jd.strip()) < 80:
            st.error("Please paste the full job description.")
        else:
            try:
                resume_text = read_pdf(resume_file.read())
                if len(resume_text) < 100:
                    st.error("We couldn't read text from this PDF. Please upload a text-based (not scanned) resume.")
                    st.stop()
                with st.spinner("Matching your resume to the job…"):
                    data = None
                    for _ in range(2):
                        try:
                            data = score_resume(parse_json(call_groq(resume_text, jd)))
                            break
                        except json.JSONDecodeError:
                            continue
                    if data is None:
                        raise ValueError("AI reply could not be read")
                st.session_state.data = data
                track(f"resume tailored · match {data['match_score']}% · {len(data['missing_skills'])} missing skills")
                st.rerun()
            except Exception as e:
                traceback.print_exc()
                if not GROQ_API_KEY:
                    st.error("The AI service isn't connected yet. Please try again later.")
                else:
                    st.error("Something went wrong. Please try again in a moment.")
                st.caption(f"Technical note: {type(e).__name__}")

# ───────────── RESULTS ─────────────
else:
    data = st.session_state.data
    score = data["match_score"]

    st.markdown(f"""
    <div class="card center">
      <h3>Your ATS Match</h3>
      <div class="ring" style="background:conic-gradient(#22d3ee {score}%, rgba(255,255,255,.12) 0)"><span>{score}%</span></div>
      <p>{data['ats_keywords_found']} of {data['ats_keywords_found'] + data['ats_keywords_missing']} job keywords are now in your resume</p>
    </div>""", unsafe_allow_html=True)

    miss = data.get("missing_skills", [])
    if miss:
        chips = "".join(f'<span class="chip">{html.escape(str(m))}</span>' for m in miss[:15])
        st.markdown(f"""<div class="card"><h3>🎯 Missing skills</h3>
          <p style="margin-bottom:.6rem">The job asks for these but your resume doesn't show them. Learn them or add them only if you truly have them.</p>{chips}</div>""", unsafe_allow_html=True)

    st.markdown('<div class="lbl">⬇ Your files are ready</div>', unsafe_allow_html=True)
    st.download_button("Download Tailored Resume", data=build_resume(data, watermark=False),
        file_name="resumereflect_resume.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_resume")
    st.write("")
    st.download_button("Download ATS Report", data=build_analysis(data), file_name="resumereflect_ats_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_report")

    # ── ₹499 counselling (UPI QR) ──
    if COUNSELLING_OPEN:
        st.markdown("""<div class="card"><h3>🚀 Personal Resume Review + Career Counselling</h3>
          <p>Get your resume reviewed personally and a one-to-one guidance session for your job search. We'll reply by email. Your tailored resume is shared with us when you book.</p></div>
          <div class="pay"><div class="big">Book your session · ₹499</div>
          <div class="sm">Scan with any UPI app · One-time payment</div></div>""", unsafe_allow_html=True)
        if st.session_state.booked:
            st.success("Thank you! We'll confirm your payment and email you within 24 hours to schedule your session.")
        elif not (UPI_ID and SMTP_USER and SMTP_PASSWORD and NOTIFY_EMAIL):
            st.error("Bookings are opening soon. Please check back shortly.")
        else:
            upi_block(499, st.session_state.ref499)
            c_email = st.text_input("Your email", placeholder="you@email.com", key="c_email").strip()
            if st.button("I've paid ₹499 – book my session", key="c_btn"):
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", c_email):
                    st.error("Enter a valid email so we can reach you.")
                else:
                    ok = notify_owner(
                        subject=f"₹499 booking – {c_email} – {st.session_state.ref499}",
                        body=(f"New counselling booking claim\n\nEmail: {c_email}\nPayment reference: {st.session_state.ref499}\n"
                              f"Candidate: {data.get('candidate_name','')}\nATS match: {data.get('match_score','')}%\n"
                              f"Missing skills: {', '.join(data.get('missing_skills', [])[:10])}\nTime: {datetime.now():%d %b %Y, %I:%M %p}\n\n"
                              f"Check PhonePe for a ₹499 payment with this reference in the note. "
                              f"If it's there, just reply to this email to schedule the session."),
                        reply_to=c_email, attachment=build_resume(data, watermark=False))
                    if ok:
                        track(f"₹499 booking claim · {st.session_state.ref499}")
                        st.session_state.booked = True
                        st.rerun()
                    else:
                        st.error("Couldn't send your booking. Please try again in a moment.")

    st.write("")
    if st.button("🔄 Tailor another resume", key="reset"):
        st.session_state.data = None
        st.session_state.booked = False
        st.rerun()

# ───────────── FOOTER ─────────────
st.markdown("""
<div class="foot">
  <b style="color:#fff;font-size:15px">⚡ ResumeReflect</b> · Made in India 🇮🇳<br>
  🔒 Your resume is processed by AI and not stored on our servers.
</div>
""", unsafe_allow_html=True)
