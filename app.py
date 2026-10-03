import streamlit as st
import groq
import pdfplumber
import smtplib, ssl, html, traceback, copy
import requests
from email.message import EmailMessage
import io, json, re
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ───────────── CONFIG (put keys in .streamlit/secrets.toml) ─────────────
GROQ_API_KEY     = st.secrets.get("GROQ_API_KEY", "")
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

def tg_send(text, doc=None):
    """Private Telegram message to the owner (used for booking alerts if email isn't set up)."""
    if not (TELEGRAM_TOKEN and TELEGRAM_CHAT_ID): return False
    try:
        base = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
        r = requests.post(f"{base}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=10)
        if doc and r.ok:
            requests.post(f"{base}/sendDocument", data={"chat_id": TELEGRAM_CHAT_ID},
                          files={"document": ("tailored_resume.docx", doc)}, timeout=20)
        return r.ok
    except Exception:
        return False

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
- It must fit ONE page: max 4 bullets per job, max 2 per project, max 3 projects, each bullet under 22 words.

RESUME:
{resume_text[:9000]}

JOB DESCRIPTION:
{jd_text[:6000]}
"""
    kw = dict(model="openai/gpt-oss-120b", messages=[{"role": "user", "content": prompt}], temperature=0.4, max_tokens=7000)
    last = None
    for extra in ({"reasoning_effort": "low", "response_format": {"type": "json_object"}}, {"reasoning_effort": "low"}, {}):
        try:   # fastest settings first, safe fallback if the API rejects one
            return client.chat.completions.create(**kw, **extra).choices[0].message.content
        except Exception as e:
            last = e
    raise last

def parse_json(text):
    text = "\n".join(l for l in text.strip().split("\n") if not l.strip().startswith("```"))
    s, e = text.find("{"), text.rfind("}") + 1
    return json.loads(text[s:e])

def score_resume(data, resume_text=""):
    """Two honest numbers:
    optimised = how many of the job keywords YOUR background supports are now used in the tailored resume
    job_match = share of ALL job keywords the tailored resume covers (skills you lack stay missing)."""
    parts = [data.get("summary", "")] + data.get("skills_technical", []) + data.get("skills_tools", []) \
          + data.get("achievements", []) + data.get("certifications", [])
    for j in data.get("work_experience", []): parts += [j.get("title", "")] + j.get("bullets", [])
    for p in data.get("projects", []): parts += [p.get("name", "")] + p.get("bullets", [])
    blob = " ".join(map(str, parts)).lower()
    orig = resume_text.lower()
    kws = list(dict.fromkeys(k.strip() for k in data.get("jd_keywords", []) if k and k.strip()))
    found = [k for k in kws if k.lower() in blob]
    missing = [k for k in kws if k.lower() not in blob]
    dropped = [k for k in missing if k.lower() in orig]
    supported = len(found) + len(dropped)
    optimised = int(round(100 * len(found) / supported)) if supported else 0
    job_match = int(round(100 * len(found) / len(kws))) if kws else 0
    data.update(match_score=max(5, min(optimised, 99)), job_match=min(job_match, 99),
                ats_keywords_found=len(found), ats_keywords_missing=len(missing),
                supported_total=supported, missing_skills=missing)
    return data

def reset_all():
    st.session_state.data = None
    st.session_state.pdf = None
    st.session_state.docx = None
    st.session_state.booked = False

# ─────────────────────────────────────────────
# DOCUMENT HELPERS
# ─────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, **kwargs):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        tag = OxmlElement(f'w:{side}')
        tag.set(qn('w:val'),   kwargs.get(side, 'nil'))
        tag.set(qn('w:sz'),    '0')
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), 'auto')
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def add_para(doc_or_cell, text, bold=False, size=10, color=None,
             italic=False, align=None, space_before=0, space_after=60):
    if hasattr(doc_or_cell, 'add_paragraph'):
        p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()
        p = doc_or_cell.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if align:
        p.alignment = align
    if text:
        run = p.add_run(text)
        run.bold   = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return p

def section_heading(doc, text, color="2E75B6"):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'),   'single')
    bottom.set(qn('w:sz'),    '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    return p

def bullet_para(doc, text, size=9.5):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text.lstrip('•-– '))
    run.font.size = Pt(size)
    return p

# ─────────────────────────────────────────────
# RESUME BUILDER  (one page · clickable links · no branding)
# ─────────────────────────────────────────────
from xml.sax.saxutils import escape as xesc
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from docx.enum.text import WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.shared import Mm

NAVY = "1F3A5F"
LEVELS = [(1.0, 4, 3, 2, 3), (0.96, 4, 3, 2, 3), (0.92, 3, 3, 2, 2), (0.88, 3, 3, 2, 2),
          (0.85, 3, 2, 1, 2), (0.84, 3, 2, 1, 0), (0.82, 2, 2, 1, 0)]   # scale, exp bullets, projects, proj bullets, achievements

def _site(url):
    u = url.lower()
    for k, n in (("linkedin", "LinkedIn"), ("github", "GitHub"), ("behance", "Behance"), ("kaggle", "Kaggle"), ("leetcode", "LeetCode")):
        if k in u: return n
    m = re.search(r"https?://(?:www\.)?([^/]+)", u)
    return m.group(1) if m else "Link"

def extract_links(file_bytes, text=""):
    """Keeps every clickable link from the original PDF (plus plain-text URLs)."""
    found = {}
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                for h in page.hyperlinks or []:
                    uri = (h.get("uri") or "").strip()
                    if not uri or uri.lower().startswith(("mailto:", "tel:")): continue
                    try:
                        box = (max(h["x0"], 0), max(h["top"], 0), min(h["x1"], page.width), min(h["bottom"], page.height))
                        label = " ".join((page.crop(box, strict=False).extract_text() or "").split())
                    except Exception:
                        label = ""
                    if not label or len(label) > 30: label = _site(uri)
                    found.setdefault(uri, label)
    except Exception:
        pass
    for m in re.findall(r"(?:https?://|www\.)[^\s|,;)]+|(?:linkedin\.com|github\.com)/[^\s|,;)]+", text):
        u = m if m.startswith("http") else "https://" + m
        if not any(u.rstrip("/").lower() in k.lower() or k.rstrip("/").lower() in u.lower() for k in found):
            found[u] = _site(u)
    return [{"label": l, "url": u} for u, l in found.items()][:8]

def _trim(data, lv):
    _, eb, mp, pb, ma = lv
    d = copy.deepcopy(data)
    d["work_experience"] = [dict(j, bullets=j.get("bullets", [])[:eb]) for j in d.get("work_experience", [])]
    d["projects"] = [dict(p, bullets=p.get("bullets", [])[:pb]) for p in d.get("projects", [])[:mp]]
    d["achievements"] = d.get("achievements", [])[:ma]
    d["certifications"] = d.get("certifications", [])[:4]
    return d

def _blocks(d, links):
    B, used, matched = [], set(), {}
    B.append(("name", d.get("candidate_name") or "Your Name"))
    skip = {"project", "using", "based", "application", "system", "management", "website", "web"}
    for i, p in enumerate(d.get("projects", [])):
        toks = [t for t in re.findall(r"[a-z0-9]{4,}", p.get("name", "").lower()) if t not in skip]
        for l in links:
            if l["url"] not in used and any(t in (l["url"] + l["label"]).lower() for t in toks):
                matched.setdefault(i, []).append(l); used.add(l["url"])
    parts = []
    if d.get("email"): parts.append((d["email"], "mailto:" + d["email"]))
    parts += [(d[k], None) for k in ("phone", "location") if d.get(k)]
    parts += [(l["label"], l["url"]) for l in links if l["url"] not in used]
    B.append(("contact", parts))
    def head(t): B.append(("head", t))
    if d.get("summary"): head("Summary"); B.append(("para", d["summary"]))
    st_, sk_ = [x for x in d.get("skills_technical", []) if x], [x for x in d.get("skills_tools", []) if x]
    if st_ or sk_:
        head("Skills")
        if st_: B.append(("kv", "Technical", ", ".join(st_)))
        if sk_: B.append(("kv", "Tools", ", ".join(sk_)))
    if d.get("work_experience"):
        head("Experience")
        for j in d["work_experience"]:
            loc = f", {j['location']}" if j.get("location") else ""
            B.append(("role", [(j.get("title", ""), None, True), (f" | {j.get('company', '')}{loc}", None, False)], j.get("dates", "")))
            B += [("bul", b) for b in j.get("bullets", [])]
    if d.get("projects"):
        head("Projects")
        for i, p in enumerate(d["projects"]):
            sp = [(p.get("name", ""), None, True)]
            for l in matched.get(i, [])[:2]: sp += [("  |  ", None, False), (_site(l["url"]), l["url"], False)]
            B.append(("role", sp, ""))
            B += [("bul", b) for b in p.get("bullets", [])]
    if d.get("education"):
        head("Education")
        for e in d["education"]:
            extra = f" (CGPA: {e['cgpa']})" if e.get("cgpa") else ""
            B.append(("role", [(e.get("degree", ""), None, True), (f", {e.get('institution', '')}{extra}", None, False)], e.get("year", "")))
    if d.get("certifications"): head("Certifications"); B.append(("para", "  •  ".join(d["certifications"])))
    if d.get("achievements"):
        head("Achievements"); B += [("bul", a) for a in d["achievements"]]
    return B

def _sty(s):
    body = max(8.0, 9.5 * s)
    return dict(name=19 * s, head=max(9.0, 10 * s), body=body, lead=body * 1.28)

def _clean(t):
    return str(t).replace("₹", "Rs.").encode("cp1252", "ignore").decode("cp1252")

def _xml(spans):
    out = ""
    for t, u, b in spans:
        x = xesc(_clean(t))
        if u: x = f'<a href="{xesc(u, {chr(34): "&quot;"})}" color="#1a56c4"><u>{x}</u></a>'
        out += f"<b>{x}</b>" if b else x
    return out

def _render_pdf(B, s, bottom=22):
    z, W = _sty(s), A4[0] - 32 * mm
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=12 * mm, bottomMargin=bottom * mm)
    def PS(**k):
        k.setdefault("fontName", "Helvetica"); k.setdefault("fontSize", z["body"]); k.setdefault("leading", z["lead"])
        return ParagraphStyle("x", **k)
    navy, story = colors.HexColor("#" + NAVY), []
    for b in B:
        k = b[0]
        if k == "name":
            story.append(Paragraph(f"<b>{xesc(_clean(b[1]))}</b>", PS(fontSize=z["name"], leading=z["name"] * 1.15, alignment=TA_CENTER, textColor=navy)))
        elif k == "contact":
            sp = []
            for i, (t, u) in enumerate(b[1]):
                if i: sp.append(("  |  ", None, False))
                sp.append((t, u, False))
            story.append(Paragraph(_xml(sp), PS(alignment=TA_CENTER)))
        elif k == "head":
            story += [Spacer(1, 5 * s), Paragraph(f"<b>{xesc(b[1].upper())}</b>", PS(fontSize=z["head"], leading=z["head"] * 1.2, textColor=navy)),
                      HRFlowable(width="100%", thickness=0.6, color=navy, spaceBefore=1, spaceAfter=2)]
        elif k == "kv":
            story.append(Paragraph(f"<b>{b[1]}:</b> {xesc(_clean(b[2]))}", PS()))
        elif k == "para":
            story.append(Paragraph(xesc(_clean(b[1])), PS()))
        elif k == "bul":
            story.append(Paragraph(xesc(_clean(b[1])), PS(leftIndent=11, bulletIndent=2), bulletText="•"))
        elif k == "role":
            t = Table([[Paragraph(_xml(b[1]), PS()), Paragraph(f"<i>{xesc(_clean(b[2]))}</i>", PS(alignment=TA_RIGHT))]], colWidths=[W - 38 * mm, 38 * mm])
            t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                   ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.5), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            story.append(t)
    doc.build(story)
    return buf.getvalue(), doc.page

def _hyperlink(p, url, text, size, bold=False):
    rid = p.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
    r, rpr = OxmlElement("w:r"), OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts"); f.set(qn("w:ascii"), "Arial"); f.set(qn("w:hAnsi"), "Arial"); rpr.append(f)
    if bold: rpr.append(OxmlElement("w:b"))
    c = OxmlElement("w:color"); c.set(qn("w:val"), "1A56C4"); rpr.append(c)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(round(size * 2)))); rpr.append(sz)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    r.append(rpr); t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    h.append(r); p._p.append(h)

def _render_docx(B, s):
    z = _sty(s)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(16); sec.top_margin = sec.bottom_margin = Mm(12)
    n = doc.styles["Normal"]; n.font.name = "Arial"; n.font.size = Pt(z["body"])
    n.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    def para(spans, size=None, lead=None, align=None, before=0, after=0, left=0, hang=0, tab=None, color=None, italic=False):
        size, lead = size or z["body"], lead or z["lead"]
        p = doc.add_paragraph(); f = p.paragraph_format
        f.space_before, f.space_after = Pt(before), Pt(after)
        f.line_spacing_rule, f.line_spacing = WD_LINE_SPACING.EXACTLY, Pt(lead)
        if left: f.left_indent = Pt(left)
        if hang: f.first_line_indent = Pt(-hang)
        if align: p.alignment = align
        if tab: f.tab_stops.add_tab_stop(tab[0], tab[1])
        for t, u, b in spans:
            if u: _hyperlink(p, u, t, size, b)
            else:
                r = p.add_run(t); r.bold = b; r.italic = italic; r.font.size = Pt(size); r.font.name = "Arial"
                if color: r.font.color.rgb = RGBColor.from_string(color)
        return p
    for b in B:
        k = b[0]
        if k == "name":
            para([(b[1], None, True)], size=z["name"], lead=z["name"] * 1.15, align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY)
        elif k == "contact":
            sp = []
            for i, (t, u) in enumerate(b[1]):
                if i: sp.append(("  |  ", None, False))
                sp.append((t, u, False))
            para(sp, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif k == "head":
            p = para([(b[1].upper(), None, True)], size=z["head"], lead=z["head"] * 1.2, before=5 * s, after=2, color=NAVY)
            bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
            for a, v in (("val", "single"), ("sz", "4"), ("space", "1"), ("color", NAVY)): bt.set(qn("w:" + a), v)
            bd.append(bt); p._p.get_or_add_pPr().append(bd)
        elif k == "kv":
            para([(b[1] + ": ", None, True), (b[2], None, False)])
        elif k == "para":
            para([(b[1], None, False)])
        elif k == "bul":
            para([("•\t" + b[1], None, False)], left=11, hang=9, tab=(Pt(11), WD_TAB_ALIGNMENT.LEFT))
        elif k == "role":
            para(list(b[1]) + [("\t", None, False), (b[2], None, False)], before=1.5, tab=(Mm(178), WD_TAB_ALIGNMENT.RIGHT))
    buf = io.BytesIO(); doc.save(buf)
    return buf.getvalue()

def build_resume_files(data, links):
    """Shrinks spacing / trims the weakest content until the resume is exactly one page."""
    for lv in LEVELS:
        plan = _trim(data, lv); B = _blocks(plan, links)
        _, pages = _render_pdf(B, lv[0])
        if pages <= 1: break
    pdf, _ = _render_pdf(B, lv[0], bottom=12)
    return plan, pdf, _render_docx(B, lv[0])

# ─────────────────────────────────────────────
# BUILD ANALYSIS DOCX
# ─────────────────────────────────────────────

def build_analysis(data):
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(1.8)
        section.right_margin  = Cm(1.8)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    score = max(5, int(data.get('match_score') or 0))

    jm = int(data.get('job_match') or 0)
    if jm >= 70:
        sc, sl = '00A651', f'STRONG JOB FIT ({jm}% of job skills matched)'
    elif jm >= 40:
        sc, sl = 'E67E22', f'PARTIAL JOB FIT ({jm}% of job skills matched) - learn the missing skills'
    else:
        sc, sl = 'C0392B', f'SKILL GAP ({jm}% of job skills matched) - focus on the missing skills'

    # ── Banner ──
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, '1B4F72')
    set_cell_borders(cell, top='nil', bottom='nil', left='nil', right='nil')

    cell.paragraphs[0].clear()
    bp = cell.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    br1 = bp.add_run("ResumeReflect")
    br1.bold = True
    br1.font.size = Pt(18)
    br1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    br2 = bp.add_run("  ·  ATS RESUME ANALYSIS REPORT")
    br2.font.size = Pt(11)
    br2.font.color.rgb = RGBColor(0x90, 0xC4, 0xF0)
    bp.paragraph_format.space_before = Pt(8)
    bp.paragraph_format.space_after  = Pt(2)

    cp = cell.add_paragraph(f"Candidate: {data.get('candidate_name', '')}")
    cp.runs[0].italic = True
    cp.runs[0].font.size = Pt(9)
    cp.runs[0].font.color.rgb = RGBColor(0xC8, 0xDF, 0xF5)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()

    # ── Score card ──
    stbl = doc.add_table(rows=1, cols=2)
    stbl.style = 'Table Grid'

    lc = stbl.cell(0, 0)
    rc = stbl.cell(0, 1)
    set_cell_bg(lc, 'F5F7FA')
    set_cell_bg(rc, 'FFFFFF')
    set_cell_borders(lc, top='nil', bottom='nil', left='nil', right='nil')
    set_cell_borders(rc, top='nil', bottom='nil', left='nil', right='nil')

    lc.paragraphs[0].clear()
    sp = lc.paragraphs[0]
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run(f"{score}%")
    sr.bold = True
    sr.font.size = Pt(36)
    sr.font.color.rgb = RGBColor(*bytes.fromhex(sc))
    sp.paragraph_format.space_before = Pt(10)
    sp.paragraph_format.space_after  = Pt(2)

    lp2 = lc.add_paragraph("RESUME OPTIMISATION SCORE")
    lp2.runs[0].bold = True
    lp2.runs[0].font.size = Pt(7.5)
    lp2.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    lp2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    lp3 = lc.add_paragraph(sl)
    lp3.runs[0].bold = True
    lp3.runs[0].font.size = Pt(8.5)
    lp3.runs[0].font.color.rgb = RGBColor(*bytes.fromhex(sc))
    lp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp3.paragraph_format.space_after = Pt(10)

    rc.paragraphs[0].clear()
    rp = rc.paragraphs[0]
    rp.paragraph_format.space_before = Pt(10)
    rr1 = rp.add_run(f"Keywords Found: {data.get('ats_keywords_found', '?')}     ")
    rr1.bold = True
    rr1.font.size = Pt(9)
    rr1.font.color.rgb = RGBColor(0x00, 0xA6, 0x51)
    rr2 = rp.add_run(f"Keywords Missing: {data.get('ats_keywords_missing', '?')}     Job skill match: {jm}%")
    rr2.bold = True
    rr2.font.size = Pt(9)
    rr2.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    doc.add_paragraph()

    # ── Strong vs Missing ──
    section_heading(doc, 'Detailed Analysis', '1B4F72')
    atbl = doc.add_table(rows=2, cols=2)
    atbl.style = 'Table Grid'

    hdr_l = atbl.cell(0, 0)
    hdr_r = atbl.cell(0, 1)
    set_cell_bg(hdr_l, 'E8F5E9')
    set_cell_bg(hdr_r, 'FDEDEC')
    set_cell_borders(hdr_l, top='nil', bottom='nil', left='nil', right='nil')
    set_cell_borders(hdr_r, top='nil', bottom='nil', left='nil', right='nil')

    hdr_l.paragraphs[0].clear()
    hl = hdr_l.paragraphs[0]
    hlr = hl.add_run("✓  WHAT MATCHES WELL")
    hlr.bold = True
    hlr.font.size = Pt(9.5)
    hlr.font.color.rgb = RGBColor(0x00, 0xA6, 0x51)
    hl.paragraph_format.space_before = Pt(6)
    hl.paragraph_format.space_after = Pt(4)

    hdr_r.paragraphs[0].clear()
    hr = hdr_r.paragraphs[0]
    hrr = hr.add_run("✗  NEEDS IMPROVEMENT")
    hrr.bold = True
    hrr.font.size = Pt(9.5)
    hrr.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    hr.paragraph_format.space_before = Pt(6)
    hr.paragraph_format.space_after = Pt(4)

    body_l = atbl.cell(1, 0)
    body_r = atbl.cell(1, 1)
    set_cell_bg(body_l, 'E8F5E9')
    set_cell_bg(body_r, 'FDEDEC')
    set_cell_borders(body_l, top='nil', bottom='nil', left='nil', right='nil')
    set_cell_borders(body_r, top='nil', bottom='nil', left='nil', right='nil')

    body_l.paragraphs[0].clear()
    for pt in data.get('strong_points', []):
        p = body_l.add_paragraph(f"• {pt.lstrip('•- ')}")
        p.runs[0].font.size = Pt(9)
        p.paragraph_format.space_after = Pt(3)

    body_r.paragraphs[0].clear()
    for ms in data.get('missing_skills', []):
        p = body_r.add_paragraph(f"• {ms.lstrip('•- ')}")
        p.runs[0].font.size = Pt(9)
        p.paragraph_format.space_after = Pt(3)

    doc.add_paragraph()

    # ── Improvement Roadmap ──
    section_heading(doc, 'Improvement Roadmap', 'E67E22')
    for i, tip in enumerate(data.get('improvement_tips', []), 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        r1 = p.add_run(f"{i}.  ")
        r1.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
        r2 = p.add_run(tip)
        r2.font.size = Pt(9.5)

    # ── Footer ──
    doc.add_paragraph()
    p = doc.add_paragraph("Generated by ResumeReflect  •  AI Resume Tailoring Tool")
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(7.5)
    p.runs[0].font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ───────────── STATE ─────────────
for k, v in {"data": None, "resume_text": "", "pdf": None, "docx": None, "booked": False}.items():
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
                file_bytes = resume_file.read()
                resume_text = read_pdf(file_bytes)
                if len(resume_text) < 100:
                    st.error("We couldn't read text from this PDF. Please upload a text-based (not scanned) resume.")
                    st.stop()
                with st.status("Starting…", expanded=True) as status:
                    st.write("📄 Reading your resume…")
                    links = extract_links(file_bytes, resume_text)
                    st.write("🔍 Matching it to the job and rewriting…")
                    data = None
                    for _ in range(2):
                        try:
                            data = score_resume(parse_json(call_groq(resume_text, jd)), resume_text)
                            break
                        except json.JSONDecodeError:
                            continue
                    if data is None:
                        raise ValueError("AI reply could not be read")
                    st.write("📐 Fitting it on one page…")
                    plan, pdf_bytes, docx_bytes = build_resume_files(data, links)
                    data = score_resume(plan, resume_text)
                    st.session_state.update(data=data, pdf=pdf_bytes, docx=docx_bytes, resume_text=resume_text)
                    status.update(label="Done! Opening your results…", state="complete", expanded=False)
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
      <h3>Resume optimisation score</h3>
      <div class="ring" style="background:conic-gradient(#22d3ee {score}%, rgba(255,255,255,.12) 0)"><span>{score}%</span></div>
      <p>Your new resume uses {data['ats_keywords_found']} of the {data['supported_total']} job keywords your background supports</p>
      <p style="margin-top:.6rem;color:#fde68a">Skill match with this job: <b>{data['job_match']}%</b></p>
    </div>""", unsafe_allow_html=True)

    miss = data.get("missing_skills", [])
    if miss:
        chips = "".join(f'<span class="chip">{html.escape(str(m))}</span>' for m in miss[:15])
        st.markdown(f"""<div class="card"><h3>🎯 What's missing</h3>
          <p style="margin-bottom:.6rem">The job asks for these, but they aren't in your resume. Learning them will raise your chances.</p>{chips}</div>""", unsafe_allow_html=True)

    st.markdown('<div class="lbl">⬇ Your files are ready</div>', unsafe_allow_html=True)
    st.download_button("Download Resume (PDF · 1 page)", data=st.session_state.pdf, file_name="resume.pdf", mime="application/pdf", key="dl_pdf")
    st.write("")
    st.download_button("Download Resume (Word · editable)", data=st.session_state.docx, file_name="resume.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_resume")
    st.write("")
    st.download_button("Download ATS Report", data=build_analysis(data), file_name="resumereflect_ats_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_report")

    # ── ₹499 counselling: book first, we email payment details ──
    if COUNSELLING_OPEN:
        alerts_ok = bool((SMTP_USER and SMTP_PASSWORD and NOTIFY_EMAIL) or (TELEGRAM_TOKEN and TELEGRAM_CHAT_ID))
        st.markdown("""<div class="card"><h3>🚀 Personal Resume Review + Career Counselling</h3>
          <p>Get your resume reviewed personally and a one-to-one guidance session for your job search. We'll reply by email. Your tailored resume is shared with us when you book.</p></div>
          <div class="pay"><div class="big">Personal session · ₹499</div>
          <div class="sm">Book your slot now. We'll email you the payment details.</div></div>""", unsafe_allow_html=True)
        if st.session_state.booked:
            st.success("Thank you! We'll email you within 24 hours with the next steps and payment details.")
        elif not alerts_ok:
            st.error("Bookings are opening soon. Please check back shortly.")
            if st.query_params.get("admin") == "1":
                st.info("Owner only – missing in Streamlit Secrets: SMTP_USER, SMTP_PASSWORD, NOTIFY_EMAIL")
        else:
            c_email = st.text_input("Your email", placeholder="you@email.com", key="c_email").strip()
            if st.button("Book my session", key="c_btn"):
                if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", c_email):
                    st.error("Enter a valid email so we can reach you.")
                else:
                    subject = f"₹499 booking request – {c_email}"
                    body = (f"New counselling booking request\n\nCustomer email: {c_email}\n"
                            f"Candidate: {data.get('candidate_name','')}\nATS match: {data.get('match_score','')}%\n"
                            f"Missing skills: {', '.join(data.get('missing_skills', [])[:10])}\nTime: {datetime.now():%d %b %Y, %I:%M %p}\n\n"
                            f"Next step: reply to this email with your payment details (Rs. 499) and schedule the session "
                            f"after you see the payment in your PhonePe.")
                    ok = False
                    if SMTP_USER and SMTP_PASSWORD and NOTIFY_EMAIL:
                        ok = notify_owner(subject, body, reply_to=c_email, attachment=st.session_state.docx)
                    if not ok:
                        ok = tg_send(subject + "\n\n" + body, st.session_state.docx)
                    if ok:
                        st.session_state.booked = True
                        st.rerun()
                    else:
                        st.error("Couldn't send your booking. Please try again in a moment.")

    st.write("")
    st.button("🔄 Tailor another resume", key="reset", on_click=reset_all)

# ───────────── FOOTER ─────────────
st.markdown("""
<div class="foot">
  <b style="color:#fff;font-size:15px">⚡ ResumeReflect</b> · Made in India 🇮🇳<br>
  🔒 Your resume is processed by AI and not stored on our servers.<br>
  © 2026 ResumeReflect. All rights reserved.
</div>
""", unsafe_allow_html=True)
