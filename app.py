import streamlit as st
import groq
import pdfplumber
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import os
import json
import re
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
GROQ_API_KEY    = st.secrets.get("GROQ_API_KEY", "")
RAZORPAY_BASIC  = st.secrets.get("RAZORPAY_BASIC", "#")
RAZORPAY_PRO    = st.secrets.get("RAZORPAY_PRO", "#")
RAZORPAY_YEARLY = st.secrets.get("RAZORPAY_YEARLY", "#")
EMAIL_DB_FILE   = "used_emails.json"

SHEET_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzN99UwHn4Bt1mJ4MMS5ZSV-cysoTC_ac6d6oMNkWB_JAGb1i2vqBX3RmrCDqIsla3G/exec"

# FIX: Tool name constant - ensures correct name in Google Sheet tracking
TOOL_NAME = "ResumeReflect"

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ResumeReflect - AI Resume Tailor",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Add OG meta tags for proper link sharing (no performance impact)
:0.3rem;">
        Got an interview after using ResumeReflect? Your story helps others — and takes 60 seconds!
    </div>
</div>
""", unsafe_allow_html=True)

r_col1, r_col2 = st.columns(2)
with r_col1:
    reviewer_name = st.text_input(
        "Your Name",
        placeholder="e.g. Rahul S. (or leave blank for Anonymous)",
        key="reviewer_name"
    )
with r_col2:
    reviewer_role = st.text_input(
        "Job Role You Applied For",
        placeholder="e.g. Software Engineer at Infosys",
        key="reviewer_role"
    )

review_text = st.text_area(
    "Your Experience",
    placeholder="How did ResumeReflect help you? Did you get an interview call? What was your ATS score before/after? Any feedback for us?",
    height=110,
    key="review_text"
)

review_rating = st.select_slider(
    "⭐ Your Rating",
    options=["1 Star", "2 Stars", "3 Stars", "4 Stars", "5 Stars"],
    value="5 Stars",
    key="review_rating"
)

if st.button("Submit My Review ✅", key="submit_review", use_container_width=True):
    if review_text and len(review_text.strip()) > 10:
        try:
            requests.post(SHEET_SCRIPT_URL, json={
                "type":      "review",
                "name":      reviewer_name.strip() if reviewer_name.strip() else "Anonymous",
                "role":      reviewer_role or "",
                "review":    review_text.strip(),
                "rating":    review_rating,
                "email":     st.session_state.get("email", ""),
                "tool_name": TOOL_NAME,  # FIX: Added correct tool name
                "timestamp": datetime.now().isoformat()
            }, timeout=5)
        except:
            pass
        st.success("🙏 Thank you! Your review has been submitted. It really helps the community.")
        st.balloons()
    else:
        st.error("Please write at least a sentence about your experience.")

st.markdown('<div style="margin-top:1.5rem"></div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;padding:2.5rem 0 1rem;color:var(--muted);font-size:12px">
    <div style="font-family:'Syne',sans-serif;font-weight:800;font-size:1rem;
        background:linear-gradient(135deg,var(--accent),var(--accent2));
        -webkit-background-clip:text;-webkit-text-fill-color:transparent;
        background-clip:text;margin-bottom:.4rem">⚡ ResumeReflect</div>
    Built for serious job seekers · Made in India 🇮🇳
    <br><br>
    <div style="font-size:11px;margin-bottom:.3rem">🔒 Your resume is processed by AI and not stored on our servers.</div>
    <div style="font-size:11px">© 2026 ResumeReflect. All rights reserved.</div>
</div>
""", unsafe_allow_html=True)

