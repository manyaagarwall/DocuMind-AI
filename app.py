import streamlit as st
from pathlib import Path

css = Path(__file__).resolve().with_name("style.css").read_text(encoding="utf-8")

st.set_page_config(
    page_title="DocuMind AI",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="auto",
)
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True
st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

# ── Imports ───────────────────────────────────────────────────────────────────
from datetime import datetime
from hashlib import sha1
from html import escape
import json
from PIL import Image
import pytesseract
from langchain_core.documents import Document
from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from qdrant_client import QdrantClient
from qdrant_client.http.models import PointStruct, VectorParams, Distance
from docx import Document as DocxDocument
import tempfile, os, subprocess, time
import pandas as pd
import requests
from requests.exceptions import RequestException
import numpy as np

# ═══════════════════════════════════════════════════════════════════════════════
#  CSS — Dark Green + Blue theme
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&display=swap');

*, *::before, *::after { box-sizing: border-box; }

:root {
    --bg: #F5F8F5;
    --bg-panel: #EAF5ED;
    --bg-elevated: #FFFFFF;
    --bg-card: #FFFFFF;
    --bg-hover: #E2F0E6;
    --bg-input: #FFFFFF;

    --purple: #D1FAE5;
    --purple-dim: rgba(22, 131, 93, 0.10);
    --purple-glow: rgba(22, 131, 93, 0.18);

    --cyan: #0F766E;
    --cyan-dim: rgba(15, 118, 110, 0.12);

    --green: #16835D;
    --green-dim: rgba(22, 131, 93, 0.12);

    --text: #24352D;
    --text-2: #52665A;
    --text-3: #718276;
    --text-4: #849287;

    --border: #DCE7DE;
    --accent: #16835D;
    --accent-hover: #116B4B;
}
.chat-input-wrap [data-testid="stChatInput"] {
  border:1px solid rgba(16,185,129,0.32) !important;
  border-radius:15px !important; background:var(--bg-elevated) !important;
  box-shadow:0 6px 24px rgba(0,0,0,0.18);
}
.chat-input-wrap [data-testid="stChatInput"]:focus-within {
  border-color:rgba(16,185,129,0.65) !important;
  box-shadow:0 0 0 3px rgba(16,185,129,0.1),0 6px 24px rgba(0,0,0,0.2) !important;
}
.chat-input-wrap textarea { color:var(--text-1) !important; }
.chat-input-wrap textarea::placeholder { color:var(--text-3) !important; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"],
[data-testid="stMain"] {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
  background: var(--bg) !important;
  color: var(--text-1) !important;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(16,185,129,0.3); border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: rgba(16,185,129,0.5); }

/* ── Layout ── */
.main .block-container,
[data-testid="stMain"] > div { padding: 0 !important; max-width: 100% !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: var(--bg-panel) !important;
  border-right: 1px solid var(--border) !important;
  width: clamp(280px, 19vw, 305px) !important;
  min-width: 280px !important; max-width: 305px !important;
}
[data-testid="stSidebar"] > div:first-child {
  padding: 0 !important;
  height: 100vh;
  overflow-y: auto;
}
[data-testid="stSidebarNav"] { display: none !important; }
[data-testid="stSidebar"] .block-container { padding: 0 !important; }
[data-testid="stButton"] button:focus-visible,
[data-testid="stFileUploader"] button:focus-visible,
[data-testid="stSelectbox"] [role="combobox"]:focus-visible,
[data-testid="stRadio"] input:focus-visible,
[data-testid="stTextInput"] input:focus-visible,
[data-testid="stTextArea"] textarea:focus-visible,
[data-testid="stChatInput"] textarea:focus-visible {
  outline: 2px solid var(--cyan) !important;
  outline-offset: 2px !important;
  box-shadow: 0 0 0 3px rgba(56,189,248,.16) !important;
}
:focus-visible { outline-color: var(--cyan); }
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    scroll-behavior: auto !important;
    animation-duration: .01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: .01ms !important;
  }
}

/* ── Typography ── */
h1,h2,h3,h4,h5,h6 { color: var(--text-1) !important; font-weight: 700 !important; letter-spacing: -0.3px; }
p, .stMarkdown p { color: var(--text-2) !important; }
label { color: var(--text-2) !important; }
span  { color: inherit !important; }

/* ══════════════════════════════════════════
   SIDEBAR
   ══════════════════════════════════════════ */
.sb-brand {
  padding: 22px 18px 18px;
  border-bottom: 1px solid var(--border);
  margin: 0 0 8px;
}
.sb-logo {
  width: 40px; height: 40px; border-radius: 12px; flex-shrink: 0;
  background: linear-gradient(145deg, #10B981 0%, #087E8B 100%);
  display: inline-flex; align-items: center; justify-content: center;
  box-shadow: 0 4px 18px var(--purple-glow);
  border: 1px solid rgba(110,231,183,0.22);
}
.sb-logo svg { width: 22px; height: 22px; stroke: #F8FAFC; }
.sb-name { font-size: 14.5px !important; font-weight: 700 !important; color: var(--text-1) !important; letter-spacing: -0.2px; }
.sb-sub  { font-size: 10px !important; color: var(--text-2) !important; letter-spacing: 0.3px; margin-top: 2px; }

.sb-sec  { font-size: 9.5px !important; font-weight: 600 !important; color: var(--text-4) !important;
  letter-spacing: 1.1px; text-transform: uppercase; padding: 16px 18px 7px; }

/* Sidebar nav buttons */
[data-testid="stSidebar"] .stButton > button {
  width: calc(100% - 28px) !important; text-align: left !important;
  background: transparent !important; border: none !important;
  border-radius: 9px !important; margin: 2px 14px !important;
  padding: 10px 12px !important;
  color: var(--text-2) !important; font-size: 13px !important; font-weight: 500 !important;
  transition: background 0.16s, color 0.16s, border-color 0.16s, transform 0.16s !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: #D1FAE5 !important;
    color: #14532D !important;
    border: 1px solid #A7F3D0 !important;
}
[data-testid="stSidebar"] .st-key-new_chat button {
  width: calc(100% - 28px) !important;
  background: linear-gradient(135deg, #D1FAE5, #CCFBF1) !important;
color: #14532D !important; font-weight: 650 !important;
  margin: 7px 14px 13px !important;
  padding: 11px 14px !important; border-radius: 10px !important;
  border: 1px solid rgba(110,231,183,0.26) !important;
  box-shadow: 0 4px 16px rgba(16,185,129,0.22) !important;
}
[data-testid="stSidebar"] .st-key-new_chat button:hover {
  background: linear-gradient(135deg, #A7F3D0, #99F6E4) !important;
  box-shadow: 0 6px 20px rgba(16,185,129,0.36) !important;
  transform: translateY(-1px) !important;
}
[data-testid="stSidebar"] .st-key-new_chat button:active {
  transform: translateY(0) scale(0.99) !important;
}
[data-testid="stSidebar"] .st-key-nav_chat button[kind="primary"],
[data-testid="stSidebar"] .st-key-nav_documents button[kind="primary"],
[data-testid="stSidebar"] .st-key-nav_history button[kind="primary"],
[data-testid="stSidebar"] .st-key-nav_settings button[kind="primary"] {
  background: rgba(16,185,129,0.14) !important;
  color: var(--text-1) !important;
  border: 1px solid rgba(16,185,129,0.26) !important;
  box-shadow: inset 3px 0 0 var(--purple) !important;
  font-weight: 600 !important;
}

/* Sidebar doc card */
.sdoc {
  margin: 3px 14px 5px; padding: 10px 11px;
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--r-md); transition: border-color 0.14s, background 0.14s;
}
.sdoc:hover { background: var(--bg-hover); border-color: var(--border-p); }
.sdoc-active { border-color: var(--purple) !important; background: rgba(16,185,129,0.1) !important; box-shadow: 0 0 0 1px rgba(16,185,129,0.08); }
.sdoc-row  { display: flex; align-items: center; gap: 8px; }
.sdoc-icon { font-size: 14px; flex-shrink: 0; }
.sdoc-name { font-size: 12px !important; font-weight: 500 !important; color: var(--text-1) !important;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1; min-width: 0; }
.sdoc-date { font-size: 10px !important; color: var(--text-3) !important; margin-top: 2px; }
.sdoc-status { display:flex; align-items:center; gap:6px; margin:8px 0 0 27px;
  font-size:10px !important; color:var(--text-2) !important; }
.sb-empty-docs { margin:4px 18px 12px; color:var(--text-3) !important; font-size:11px !important; }
.sb-docs-head { display:flex; align-items:center; justify-content:space-between; }
.sb-docs-head .sb-sec { padding-right: 0; }
[data-testid="stSidebar"] .st-key-sidebar_add_docs button {
  width: 30px !important; min-height: 30px !important; padding: 0 !important;
  margin: 8px 15px 0 0 !important; border-radius: 8px !important;
  color: var(--text-2) !important; background: var(--bg-elevated) !important;
}
[data-testid="stSidebar"] .st-key-sidebar_add_docs button:hover {
  color: var(--text-1) !important; border-color: var(--border-p) !important;
}
[data-testid="stSidebar"] [class*="st-key-sel_"] button {
  width: calc(100% - 28px) !important; margin: 0 14px 8px !important;
  padding: 5px 10px !important; min-height: 30px !important;
  border: 1px solid transparent !important; background: transparent !important;
  color: var(--text-3) !important; font-size: 11px !important;
}
[data-testid="stSidebar"] [class*="st-key-sel_"] button:hover {
  color: var(--text-1) !important; border-color: var(--border-p) !important;
  background: rgba(16,185,129,0.08) !important;
}

/* Badges */
.badge { display: inline-flex; align-items: center; gap: 3px; padding: 2px 8px;
  border-radius: 99px; font-size: 10px !important; font-weight: 600 !important; }
.bdg-green { background: var(--green-dim);  color: var(--green) !important; border: 1px solid rgba(16,217,138,0.25); }
.bdg-cyan  { background: var(--cyan-dim);   color: var(--cyan)  !important; border: 1px solid rgba(56,189,248,0.25); }
.bdg-purple{ background: var(--purple-dim); color: #6EE7B7      !important; border: 1px solid var(--border-p); }
.bdg-amber { background: rgba(245,158,11,0.1); color: var(--amber) !important; border: 1px solid rgba(245,158,11,0.25); }
.bdg-red   { background: rgba(255,94,94,0.1);  color: var(--red)   !important; border: 1px solid rgba(255,94,94,0.2); }

/* Status dot */
.sdot { width: 6px; height: 6px; border-radius: 50%; display: inline-block; flex-shrink: 0; }
.s-green { background: var(--green); box-shadow: 0 0 6px rgba(16,217,138,0.6); }
.s-red   { background: var(--red);   box-shadow: 0 0 5px rgba(255,94,94,0.5); }
.s-amber { background: var(--amber); box-shadow: 0 0 5px rgba(245,158,11,0.5); }

.sb-hr { height: 1px; background: var(--border); margin: 12px 16px 7px; }

/* Sidebar status */
.sb-stat { margin: 8px 14px 0; padding: 12px; font-size: 11px !important; color: var(--text-2) !important;
  background: rgba(17,31,54,0.72); border: 1px solid var(--border); border-radius: 11px; }
.sb-stat-row { display: flex; align-items: center; gap: 8px; margin-bottom: 7px; }
.sb-model { font-size: 10px !important; color: var(--text-3) !important; padding: 2px 0 0 14px; }
.sb-status-label { margin-left:auto; color:var(--text-3) !important; font-size:10px !important; text-align:right; }
.sb-vector-row { margin-top:9px; margin-bottom:0; }
.sb-profile { display:flex; align-items:center; gap:10px; margin:12px 14px 14px; padding:12px 2px 2px;
  border-top:1px solid var(--border); }
.sb-avatar { width:34px; height:34px; display:flex; align-items:center; justify-content:center;
  border-radius:50%; flex:none; color:#EAE6FF; font-size:11px; font-weight:700;
  background:linear-gradient(145deg,rgba(16,185,129,0.45),rgba(56,189,248,0.2));
  border:1px solid rgba(148,170,205,0.2); }
.sb-profile-name { color:var(--text-1) !important; font-size:11.5px !important; font-weight:600 !important; }
.sb-profile-role { color:var(--text-3) !important; font-size:10px !important; margin-top:2px; }

@media (max-width: 1100px) {
  [data-testid="stSidebar"] { min-width: 260px !important; width: 260px !important; max-width: 260px !important; }
}
@media (max-width: 1100px) {
  [data-testid="stHorizontalBlock"]:has(.app-hdr) {
    display:flex !important; flex-wrap:wrap !important; gap:0 !important;
  }
  [data-testid="stHorizontalBlock"]:has(.app-hdr) > [data-testid="stColumn"]:first-child {
    flex:1 1 100% !important; min-width:0 !important; width:100% !important;
  }
  [data-testid="stHorizontalBlock"]:has(.app-hdr) > [data-testid="stColumn"]:last-child {
    flex:1 1 100% !important; min-width:0 !important; width:100% !important;
  }
  .sp { height:auto; min-height:0; max-height:54vh; border-left:0; border-top:1px solid var(--border); }
}
@media (max-width: 767px) {
  [data-testid="stSidebar"] {
    min-width: min(88vw, 320px) !important;
    width: min(88vw, 320px) !important;
    max-width: min(88vw, 320px) !important;
  }
  [data-testid="stSidebar"] > div:first-child { height: 100dvh; }
  .app-hdr { padding-left: 16px; padding-right: 16px; }
  .msgs { padding-left: 16px; padding-right: 16px; }
  .composer { padding-left: 14px; padding-right: 14px; }
  .sp { height: auto; min-height: 0; max-height: 60vh; border-left: 0; border-top: 1px solid var(--border); }
  .msg-u-inner { max-width: 88%; }
  .wlc-title { font-size:24px !important; }
  .ah-sub { max-width:60vw; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
}

/* ══════════════════════════════════════════
   APP HEADER (chat view top bar)
   ══════════════════════════════════════════ */
.app-hdr {
  padding: 12px 24px;
  border-bottom: 1px solid var(--border);
  background: rgba(14,41,40,0.96);
  backdrop-filter: blur(14px);
  display: flex; align-items: center; justify-content: space-between;
  position: sticky; top: 0; z-index: 100;
}
.ah-l { display: flex; align-items: center; gap: 12px; }
.ah-ficon {
  width: 34px; height: 34px; border-radius: 9px; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center; font-size: 16px;
  border: 1px solid var(--border-p);
}
.ah-title { font-size: 13.5px !important; font-weight: 600 !important;
  color: var(--text-1) !important; margin:0; }
.ah-sub   { font-size: 10.5px !important; color: var(--text-3) !important; margin-top: 1px; }
.ah-r     { display: flex; align-items: center; gap: 8px; font-size: 11px !important; color: var(--text-3) !important; }
.ah-file-meta { color:var(--text-3) !important; font-size:10px !important; }
.doc-ready { display:inline-flex; align-items:center; gap:5px; padding:3px 8px; border-radius:99px;
  background:rgba(16,217,138,0.1); border:1px solid rgba(16,217,138,0.2);
  color:var(--green) !important; font-size:10px !important; font-weight:600 !important; vertical-align:middle; }
.doc-ready .sdot { width:5px; height:5px; }

/* Header glow bar */
.hdr-glow {
  height: 1px;
  background: linear-gradient(90deg, transparent 0%, rgba(16,185,129,0.4) 30%, rgba(56,189,248,0.3) 70%, transparent 100%);
}

/* ══════════════════════════════════════════
   WELCOME SCREEN
   ══════════════════════════════════════════ */
.wlc-wrap {
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; padding: 48px 24px 24px; text-align: center;
}
.wlc-orb {
  width: 76px; height: 76px; border-radius: 22px; margin: 0 auto 26px;
  background: linear-gradient(135deg, rgba(16,185,129,0.2), rgba(56,189,248,0.15), rgba(56,189,248,0.1));
  border: 1px solid rgba(16,185,129,0.4);
  display: flex; align-items: center; justify-content: center; font-size: 34px;
  box-shadow: 0 0 48px rgba(16,185,129,0.18), var(--sh-md);
}
.wlc-title {
  font-size: 27px !important; font-weight: 800 !important; letter-spacing: -0.7px;
  line-height: 1.2; margin:0 0 12px;
  background: linear-gradient(135deg, #FFFFFF 0%, #6EE7B7 45%, #38BDF8 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
}
.wlc-sub {
  font-size: 14px !important; color: var(--text-2) !important;
  max-width: 480px; margin: 0 auto 36px; line-height: 1.7;
}

/* Upload zone */
.upzone {
  border: 2px dashed rgba(16,185,129,0.3); border-radius: var(--r-xl);
  padding: 34px 24px; text-align: center;
  background: linear-gradient(135deg, rgba(16,185,129,0.04), rgba(56,189,248,0.03));
  transition: all 0.2s ease; max-width: 500px; margin: 0 auto 28px; width: 100%;
}
.upzone:hover {
  border-color: rgba(16,185,129,0.6);
  background: linear-gradient(135deg, rgba(16,185,129,0.08), rgba(56,189,248,0.05));
  box-shadow: var(--sh-glow);
}
.upzone-icon  { font-size: 28px; margin-bottom: 10px; display: block; }
.upzone-title { font-size: 14.5px !important; font-weight: 600 !important; color: var(--text-1) !important; }
.upzone-sub   { font-size: 12px !important; color: var(--text-3) !important; margin-top: 4px; }
.upzone-types { font-size: 10.5px !important; color: var(--text-4) !important; margin-top: 6px; }

/* Suggestion cards */
.sug-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; max-width: 540px; width: 100%; }
.sug-card {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--r-lg); padding: 14px 15px; text-align: left;
  transition: all 0.17s ease;
}
.sug-card:hover {
  background: var(--bg-hover); border-color: var(--border-p);
  box-shadow: var(--sh-glow); transform: translateY(-1px);
}
[data-testid="stMain"] [class*="st-key-chat_suggestion_"] [data-testid="stButton"] > button {
  min-height:96px !important; height:100% !important; white-space:normal !important;
  text-align:left !important; padding:13px !important; line-height:1.4 !important;
  color:var(--text-1) !important; background:var(--bg-card) !important;
  border:1px solid var(--border) !important; border-radius:var(--r-lg) !important;
}
[data-testid="stMain"] [class*="st-key-chat_suggestion_"] [data-testid="stButton"] > button:hover {
  background:var(--bg-hover) !important; border-color:var(--border-p) !important;
}
.sug-icon  { font-size: 18px; margin-bottom: 8px; display: block; }
.sug-title { font-size: 12.5px !important; font-weight: 600 !important; color: var(--text-1) !important; }
.sug-desc  { font-size: 11px !important; color: var(--text-3) !important; margin-top: 3px; line-height: 1.45; }

/* ══════════════════════════════════════════
   CHAT MESSAGES
   ══════════════════════════════════════════ */
.msgs {
  padding: 18px 24px 6px;
  max-height: calc(100vh - 270px);
  overflow-y: auto;
  overscroll-behavior: contain;
}

/* User message */
.msg-u-row { display: flex; justify-content: flex-end; margin-bottom: 18px; }
.msg-u-inner { max-width: 66%; }
.msg-u-bub {
  background: linear-gradient(135deg, rgba(16,185,129,0.2), rgba(56,189,248,0.12));
  border: 1px solid rgba(16,185,129,0.28);
  border-radius: 16px 16px 4px 16px; padding: 11px 15px;
  font-size: 14px !important; color: var(--text-1) !important; line-height: 1.65;
  box-shadow: var(--sh-sm);
}
.msg-u-meta {
  font-size: 10px !important; color: var(--text-4) !important;
  text-align: right; margin-top: 4px; padding-right: 2px;
}

/* AI message */
.msg-a-row { display: flex; gap: 11px; align-items: flex-start; margin-bottom: 22px; }
.msg-a-avatar {
  width: 32px; height: 32px; border-radius: 9px; flex-shrink: 0;
  background: linear-gradient(135deg, #10B981, #38BDF8);
  display: flex; align-items: center; justify-content: center; font-size: 14px;
  box-shadow: 0 0 12px var(--purple-glow); margin-top: 2px;
}
.msg-a-inner { flex: 1; min-width: 0; }
.msg-a-label { font-size: 11px !important; font-weight: 700 !important; color: #38BDF8 !important;
  margin-bottom: 7px; letter-spacing: 0.3px; text-transform: uppercase; }
.msg-a-body  {
  font-size: 14px !important; color: var(--text-1) !important;
  line-height: 1.78; max-width: 800px;
}
.msg-a-body p       { margin: 0 0 12px; color: var(--text-1) !important; }
.msg-a-body strong  { color: var(--text-1) !important; font-weight: 650; }
.msg-a-body h1, .msg-a-body h2, .msg-a-body h3, .msg-a-body h4 {
  color: #A7F3D0 !important; font-weight: 700 !important; line-height: 1.35;
  margin: 20px 0 8px;
}
.msg-a-body h1 { font-size: 20px !important; }
.msg-a-body h2 { font-size: 17px !important; }
.msg-a-body h3, .msg-a-body h4 { font-size: 14px !important; color: var(--cyan) !important; }
.msg-a-body ul, .msg-a-body ol { color: var(--text-1) !important; padding-left: 22px; margin: 5px 0 14px; }
.msg-a-body li { margin: 5px 0; color: var(--text-1) !important; }
.msg-a-body blockquote {
  margin: 12px 0; padding: 9px 14px; border-left: 3px solid var(--cyan);
  background: rgba(56,189,248,0.06); color: var(--text-2) !important;
}
.msg-a-body code {
  background: rgba(16,185,129,0.12) !important; color: #A7F3D0 !important;
  border: 1px solid rgba(16,185,129,0.25) !important; border-radius: 4px !important;
  padding: 1px 5px !important; font-size: 12.5px !important;
}
.msg-a-body pre {
  background: var(--bg-elevated) !important; border: 1px solid var(--border) !important;
  border-radius: 10px !important; padding: 13px 15px !important; overflow-x: auto;
}
.msg-a-body table { width:100%; border-collapse:collapse; margin:12px 0; font-size:12.5px; }
.msg-a-body th, .msg-a-body td { padding:8px 10px; border:1px solid var(--border); text-align:left; }
.msg-a-body th { color:#A7F3D0; background:rgba(16,185,129,0.1); }
.msg-a-body hr { border:0; border-top:1px solid var(--border); margin:16px 0; }
[data-testid="stChatMessage"] { padding:0 !important; background:transparent !important; border:0 !important; }
[data-testid="stChatMessageContent"] { color:var(--text-1) !important; line-height:1.72 !important; }
[data-testid="stChatMessageContent"] p { color:var(--text-1) !important; line-height:1.72; margin-bottom:11px; }
[data-testid="stChatMessageContent"] h1,
[data-testid="stChatMessageContent"] h2 { color:#A7F3D0 !important; margin:20px 0 8px; }
[data-testid="stChatMessageContent"] h3,
[data-testid="stChatMessageContent"] h4 { color:var(--cyan) !important; margin:17px 0 7px; }
[data-testid="stChatMessageContent"] ul,
[data-testid="stChatMessageContent"] ol { padding-left:22px; margin:5px 0 14px; }
[data-testid="stChatMessageContent"] li { margin:5px 0; }
[data-testid="stChatMessageContent"] table { width:100%; border-collapse:collapse; font-size:12.5px; }
[data-testid="stChatMessageContent"] th,
[data-testid="stChatMessageContent"] td { padding:8px 10px; border:1px solid var(--border); text-align:left; }
[data-testid="stChatMessageContent"] th { color:#A7F3D0; background:rgba(16,185,129,0.1); }
.chat-ai-label { color:#38BDF8 !important; font-size:10px !important; font-weight:700 !important;
  letter-spacing:1px; margin-bottom:6px; }
.msg-a-meta { font-size: 10px !important; color: var(--text-4) !important; margin-top: 8px; }
.msg-a-actions { display:flex; align-items:center; gap:5px; margin:4px 0 22px 43px; }
.msg-a-actions [data-testid="stButton"] { width:auto !important; }
.msg-a-actions [data-testid="stButton"] button {
  min-height:30px !important; padding:4px 9px !important; font-size:11px !important;
  border:1px solid transparent !important; background:transparent !important;
  color:var(--text-3) !important; box-shadow:none !important;
}
.msg-a-actions [data-testid="stButton"] button:hover {
  color:var(--text-1) !important; background:var(--bg-hover) !important;
  border-color:var(--border) !important; transform:none !important;
}
.msg-a-actions [data-testid="stButton"] button[kind="primary"] {
  color:#A7F3D0 !important; background:rgba(16,185,129,0.13) !important;
  border-color:rgba(16,185,129,0.22) !important;
}
.copy-widget { display:flex; align-items:center; height:32px; }
.copy-widget button {
  display:flex; align-items:center; gap:6px; height:30px; padding:0 9px;
  color:#A5B8BB; font:500 11px Inter, sans-serif; background:transparent;
  border:1px solid transparent; border-radius:8px; cursor:pointer;
}
.copy-widget button:hover { color:#F8FAFC; background:var(--bg-hover); border-color:var(--border); }
.copy-widget svg { width:13px; height:13px; }
.feedback-note { color:var(--green) !important; font-size:10px !important; padding:7px 2px; }

/* Source chips inside message */
.src-row { display: flex; flex-wrap: wrap; gap: 5px; margin-top: 12px;
  padding-top: 11px; border-top: 1px solid var(--border); }
.src-lbl { font-size: 9.5px !important; font-weight: 600 !important; color: var(--text-3) !important;
  letter-spacing: 0.9px; text-transform: uppercase; width: 100%; margin-bottom: 4px; }
.src-chip {
  display: inline-flex; align-items: center; gap: 4px;
  background: var(--cyan-dim); border: 1px solid var(--border-c);
  border-radius: 5px; padding: 3px 8px;
  font-size: 10.5px !important; color: var(--cyan) !important;
  transition: background 0.14s;
}
.src-chip:hover { background: rgba(56,189,248,0.2); }

/* ══════════════════════════════════════════
   CHAT COMPOSER
   ══════════════════════════════════════════ */
.composer {
  padding: 12px 24px 18px;
  border-top: 1px solid var(--border);
  background: linear-gradient(180deg, rgba(11,32,29,0.92), var(--bg-panel) 24%);
  position: sticky; bottom: 0; z-index: 5;
}
.composer-box {
  background: var(--bg-elevated); border: 1px solid var(--border);
  border-radius: 14px; transition: border-color 0.2s, box-shadow 0.2s;
  max-width: 860px;
}
.composer-box:focus-within {
  border-color: var(--border-p);
  box-shadow: 0 0 0 3px rgba(16,185,129,0.1), var(--sh-glow);
}
.composer-hint {
  font-size: 11px !important; color: var(--text-4) !important;
  text-align: center; margin-top: 8px; max-width: 860px;
}

/* ══════════════════════════════════════════
   DOCUMENT CARDS (Documents page)
   ══════════════════════════════════════════ */
.documents-page { max-width: 1440px; margin: 0 auto; padding: 0 28px 44px; }
.docs-heading { display:flex; align-items:flex-start; justify-content:space-between; gap:20px;
  padding:30px 0 22px; }
.docs-title { color:var(--text-1) !important; font-size:26px !important; font-weight:750 !important;
  letter-spacing:-0.7px; margin:0 !important; }
.docs-subtitle { color:var(--text-3) !important; font-size:13px !important; margin-top:6px; }
.docs-count { display:inline-flex; align-items:center; gap:7px; margin-top:4px; padding:7px 11px;
  color:var(--text-2) !important; background:var(--bg-elevated); border:1px solid var(--border);
  border-radius:99px; font-size:11px !important; white-space:nowrap; }
.docs-count strong { color:var(--text-1) !important; }
.upload-panel { display:flex; align-items:center; gap:16px; padding:19px 21px; margin-bottom:20px;
  background:linear-gradient(115deg,rgba(16,185,129,.09),rgba(16,45,50,.92) 52%,rgba(56,189,248,.035));
  border:1px dashed rgba(52,211,153,.48); border-radius:16px; transition:border-color .18s,box-shadow .18s; }
.upload-panel:hover { border-color:rgba(52,211,153,.78); box-shadow:0 0 26px rgba(16,185,129,.1); }
.upload-icon { display:flex; align-items:center; justify-content:center; width:48px; height:48px;
  flex:none; border-radius:50%; background:rgba(16,185,129,.16); border:1px solid rgba(16,185,129,.28);
  font-size:22px; }
.upload-copy { flex:1; min-width:0; }
.upload-title { color:var(--text-1) !important; font-size:14px !important; font-weight:650 !important; }
.upload-desc { color:var(--text-3) !important; font-size:11.5px !important; margin-top:4px; }
.upload-meta { color:var(--text-4) !important; font-size:10px !important; margin-top:7px; }
.st-key-documents_upload_control [data-testid="stFileUploaderDropzone"] {
  min-height:72px; padding:8px 14px !important; background:rgba(7,30,32,.42) !important;
  border:1px dashed rgba(52,211,153,.34) !important; border-radius:12px !important;
}
.st-key-documents_upload_control { padding:0 28px 16px; margin-top:-18px; }
.st-key-documents_upload_control [data-testid="stFileUploader"] { width:100%; }
.st-key-documents_upload_control [data-testid="stFileUploaderDropzoneInstructions"] {
  color:var(--text-2) !important; font-size:11px !important;
}
.st-key-documents_upload_control [data-testid="stFileUploaderDropzone"] button {
  min-height:34px !important; font-size:11px !important;
}
.st-key-documents_upload_control [data-testid="stFileUploaderDropzone"]:hover {
  border-color:var(--border-p) !important;
}
.st-key-documents_upload_control [data-testid="stFileUploaderDropzone"] button {
  border-radius:8px !important; border:1px solid rgba(110,231,183,.26) !important;
  background:linear-gradient(135deg,#10B981,#087E8B) !important; color:#fff !important;
}
.docs-toolbar { align-items:center; padding:15px 16px 6px; margin:0 0 15px;
  background:rgba(14,41,40,.82); border:1px solid var(--border); border-radius:13px; }
.docs-toolbar [data-testid="stTextInput"] input,
.docs-toolbar [data-testid="stSelectbox"] [data-baseweb="select"] > div {
  background:var(--bg-input) !important; border-color:var(--border) !important; border-radius:9px !important;
}
.docs-grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; }
.doc-card { min-width:0; padding:16px; margin-bottom:13px; background:var(--bg-card);
  border:1px solid var(--border); border-radius:14px; box-shadow:var(--sh-card);
  transition:background .16s,border-color .16s,transform .16s; }
.doc-card:hover { background:#163A3D; border-color:rgba(16,185,129,.44); transform:translateY(-1px); }
.doc-card-selected { border-color:rgba(16,185,129,.78); box-shadow:0 0 0 1px rgba(16,185,129,.16),var(--sh-card); }
.doc-card-top { display:flex; align-items:center; gap:11px; min-width:0; }
.doc-card-icon { display:flex; align-items:center; justify-content:center; width:42px; height:42px;
  flex:none; border:1px solid var(--border); border-radius:11px; font-size:20px; }
.doc-card-name { color:var(--text-1) !important; font-size:12.5px !important; font-weight:600 !important;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.doc-card-type { color:var(--text-3) !important; font-size:9.5px !important; text-transform:uppercase;
  letter-spacing:.65px; margin-top:4px; }
.doc-status { display:inline-flex; align-items:center; gap:6px; padding:4px 8px; margin-top:14px;
  border-radius:99px; font-size:10px !important; font-weight:600 !important; }
.doc-status-ready { background:var(--green-dim); color:var(--green) !important;
  border:1px solid rgba(16,217,138,.22); }
.doc-status-error { background:rgba(255,94,94,.09); color:var(--red) !important;
  border:1px solid rgba(255,94,94,.22); }
.doc-card-meta { display:flex; flex-wrap:wrap; gap:7px 12px; padding-top:12px; margin-top:12px;
  border-top:1px solid var(--border); color:var(--text-3) !important; font-size:10px !important; }
.doc-error { padding:9px 10px; margin-top:10px; color:#FF9A9A !important; background:rgba(255,94,94,.06);
  border:1px solid rgba(255,94,94,.13); border-radius:8px; font-size:10.5px !important; line-height:1.45; }
.doc-preview { padding:14px; margin:8px 0 16px; background:rgba(14,41,40,.84);
  border:1px solid var(--border-p); border-radius:12px; }
.doc-preview-title { color:var(--text-1) !important; font-size:12.5px !important; font-weight:650 !important; }
.doc-empty { text-align:center; padding:42px 20px; margin-top:12px; background:rgba(14,41,40,.76);
  border:1px solid var(--border); border-radius:15px; }
.doc-empty-icon { font-size:29px; margin-bottom:10px; }
.doc-empty-title { color:var(--text-1) !important; font-size:14px !important; font-weight:650 !important; }
.doc-empty-copy { color:var(--text-3) !important; font-size:11.5px !important; margin-top:5px; }
.doc-actions [data-testid="stButton"] button { font-size:11px !important; min-height:34px !important; }
@media (max-width:1050px) {
  .docs-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
  .documents-page { padding-left:20px; padding-right:20px; }
  [data-testid="stHorizontalBlock"]:has(.doc-card) { flex-wrap:wrap !important; }
  [data-testid="stHorizontalBlock"]:has(.doc-card) > [data-testid="column"] {
    flex:1 1 48% !important; min-width:0 !important; width:48% !important;
  }
  .st-key-docs_toolbar [data-testid="stHorizontalBlock"] { flex-wrap:wrap !important; }
  .st-key-docs_toolbar [data-testid="stHorizontalBlock"] > [data-testid="column"] {
    flex:1 1 48% !important; min-width:0 !important; width:48% !important;
  }
  .st-key-docs_toolbar [data-testid="stHorizontalBlock"] > [data-testid="column"]:first-child {
    flex:1 1 100% !important; width:100% !important;
  }
}
@media (max-width:680px) {
  .docs-grid { grid-template-columns:minmax(0,1fr); }
  .documents-page { padding-left:14px; padding-right:14px; }
  .st-key-documents_upload_control { padding-left:14px; padding-right:14px; }
  [data-testid="stHorizontalBlock"]:has(.doc-card) > [data-testid="column"] {
    flex:1 1 100% !important; width:100% !important;
  }
  .docs-heading { padding-top:22px; }
  .docs-title { font-size:22px !important; }
  .upload-panel { align-items:flex-start; flex-wrap:wrap; padding:15px; }
  .upload-icon { width:40px; height:40px; font-size:18px; }
  .ah-title { font-size:12px !important; }
  .ah-r { font-size:10px !important; }
}
.dc {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--r-lg); padding: 15px 17px; margin-bottom: 11px;
  transition: all 0.17s ease; box-shadow: var(--sh-card);
  position: relative; overflow: hidden;
}
.dc::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, #10B981, #34D399, #38BDF8);
  opacity: 0; transition: opacity 0.18s;
}
.dc:hover { border-color: var(--border-p); background: var(--bg-hover); }
.dc:hover::before { opacity: 1; }
.dc-top  { display: flex; align-items: center; gap: 11px; }
.dc-ficon { width: 40px; height: 40px; border-radius: 10px; flex-shrink: 0;
  display: flex; align-items: center; justify-content: center;
  font-size: 20px; background: var(--bg-elevated); border: 1px solid var(--border); }
.dc-name  { font-size: 13.5px !important; font-weight: 600 !important;
  color: var(--text-1) !important; flex: 1; min-width: 0;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.dc-pill  { font-size: 9.5px !important; font-weight: 700 !important;
  color: var(--purple) !important; background: var(--purple-dim);
  border: 1px solid var(--border-p); border-radius: 4px; padding: 2px 7px;
  letter-spacing: 0.5px; flex-shrink: 0; }
.dc-meta  { font-size: 11px !important; color: var(--text-3) !important;
  margin-top: 7px; display: flex; gap: 14px; align-items: center; flex-wrap: wrap; }
.dc-prev  { font-size: 11.5px !important; color: var(--text-4) !important;
  line-height: 1.55; margin-top: 9px; padding-top: 9px;
  border-top: 1px solid var(--border); font-style: italic; }

/* ══════════════════════════════════════════
   RETRIEVED SOURCES PANEL
   ══════════════════════════════════════════ */
.sp {
  border-left: 1px solid var(--border);
  background: var(--bg-panel); overflow-y: auto; height: 100vh; max-height: 100vh;
  overscroll-behavior: contain;
}

.sp-hdr {
  padding: 14px 14px 10px;
  border-bottom: 1px solid var(--border);
  background: rgba(13,34,34,0.82);
  backdrop-filter: blur(10px);
  position: sticky; top: 0; z-index: 10;
  display: flex; align-items: center; justify-content: space-between;
}
.sp-title { font-size: 11px !important; font-weight: 700 !important;
  color: var(--text-2) !important; text-transform: uppercase; letter-spacing: 1px;
  display:flex; align-items:center; gap:8px; }
.sp-layers-icon { width:15px; height:15px; color:var(--cyan); flex:none; }
.sp-count { background: var(--bg-card); border: 1px solid var(--border-c);
  border-radius: 99px; padding: 3px 9px;
  font-size: 10px !important; color: var(--cyan) !important; font-weight: 600 !important; }
.sp-body  { padding: 12px 11px; }

/* Source card */
.sc {
  background: var(--bg-card); border: 1px solid var(--border);
  border-radius: 13px; padding: 13px 12px 13px 14px;
  margin-bottom: 9px; position: relative; overflow: hidden;
  transition: border-color 0.14s, background 0.14s;
}
.sc::before {
  content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 3px;
  background: linear-gradient(180deg, var(--purple), #38BDF8);
}
.sc:hover { background: var(--bg-hover); border-color: var(--border-p); }
.sc-row  { display: flex; align-items: flex-start; gap: 9px; }
.sc-num  { width: 20px; height: 20px; border-radius: 5px; flex-shrink: 0;
  background: var(--purple-dim); border: 1px solid var(--border-p);
  display: flex; align-items: center; justify-content: center;
  font-size: 10px !important; font-weight: 700 !important; color: #6EE7B7 !important; }
.sc-body { flex: 1; min-width: 0; }
.sc-file { font-size: 11.5px !important; font-weight: 600 !important;
  color: #7DD3FC !important; margin-bottom: 4px;
  display: flex; align-items: center; gap: 5px; }
.sc-file-name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.sc-id   { font-size: 9.5px !important; color: var(--text-3) !important;
  background: var(--bg-elevated); border: 1px solid var(--border);
  border-radius: 4px; padding: 1px 5px; flex-shrink: 0; }
.sc-text, .sc-preview, .sc-full { font-size: 11.5px !important; color: var(--text-2) !important; line-height: 1.65; overflow-wrap:anywhere; }
.sc-details { margin-top:4px; }
.sc-details summary { list-style:none; cursor:pointer; }
.sc-details summary::-webkit-details-marker { display:none; }
.sc-expand { display:flex; align-items:center; gap:5px; margin-top:8px;
  color:var(--cyan) !important; font-size:10px !important; font-weight:600; }
.sc-expand::after { content:""; width:6px; height:6px; border-right:1.5px solid currentColor;
  border-bottom:1.5px solid currentColor; transform:rotate(45deg); transition:transform .15s; }
.sc-details[open] .sc-expand::after { transform:rotate(225deg); }
.sc-full { max-height:360px; overflow:auto; margin-top:9px; padding:10px;
  white-space:normal; background:rgba(11,32,29,.72); border:1px solid var(--border); border-radius:8px; }
.sp-footer { padding: 10px 14px; border-top: 1px solid var(--border);
  font-size: 10px !important; color: var(--text-3) !important; line-height: 1.6; }
.sp-empty { padding:18px 14px; color:var(--text-3) !important; font-size:11.5px !important; line-height:1.6; }
.sp-empty-title { display:block; color:var(--text-1) !important; font-size:12px !important; font-weight:600; margin-bottom:5px; }
.sp-loading { padding:20px 14px; color:var(--text-2) !important; font-size:11.5px !important;
  display:flex; align-items:center; gap:9px; }
.sp-loading-dot { width:8px; height:8px; border-radius:50%; background:var(--cyan);
  box-shadow:0 0 12px rgba(56,189,248,.65); animation:source-pulse 1s ease-in-out infinite alternate; }
@keyframes source-pulse { to { opacity:.35; transform:scale(.78); } }

/* ══════════════════════════════════════════
   PAGE HEADERS
   ══════════════════════════════════════════ */
.ph { padding: 26px 26px 0; }
.ph-title { font-size: 22px !important; font-weight: 800 !important;
  color: var(--text-1) !important; letter-spacing: -0.5px; margin:0 !important; }
.ph-sub { font-size: 13px !important; color: var(--text-3) !important;
  margin-top: 5px; margin-bottom: 20px; }

/* ══════════════════════════════════════════
   HISTORY CARDS
   ══════════════════════════════════════════ */
.hc { background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--r-md); padding: 13px 15px; margin-bottom: 8px;
  transition: all 0.14s ease; }
.hc:hover { background: var(--bg-hover); border-color: var(--border-p); }
.hc-who  { font-size: 10.5px !important; color: var(--text-3) !important; margin-bottom: 5px; }
.hc-body { font-size: 13.5px !important; color: var(--text-1) !important; line-height: 1.6; }

/* ══════════════════════════════════════════
   SETTINGS CARDS
   ══════════════════════════════════════════ */
.stg-card { background: var(--bg-card); border: 1px solid var(--border);
  border-radius: var(--r-lg); padding: 18px; margin-bottom: 13px;
  box-shadow: var(--sh-card); }
.stg-title { font-size: 13px !important; font-weight: 600 !important;
  color: var(--text-1) !important; margin-bottom: 14px; display: flex; align-items: center; gap: 7px; }
.stg-row { display: flex; align-items: center; gap: 10px;
  padding: 8px 0; border-bottom: 1px solid var(--border); }
.stg-row:last-child { border-bottom: none; }
.stg-key { font-size: 11.5px !important; color: var(--text-3) !important; flex-shrink: 0; width: 100px; }
.stg-val { font-size: 12.5px !important; color: var(--text-1) !important; flex: 1; }

/* ══════════════════════════════════════════
   STREAMLIT WIDGET OVERRIDES
   ══════════════════════════════════════════ */
.stTextInput > div > div > input,
.stTextArea  > div > div > textarea {
  background: var(--bg-input) !important; color: var(--text-1) !important;
  border: 1px solid var(--border) !important; border-radius: var(--r-md) !important;
  font-family: 'Inter', sans-serif !important; font-size: 14px !important;
}
.stTextInput > div > div > input:focus,
.stTextArea  > div > div > textarea:focus {
  border-color: var(--border-p) !important;
  box-shadow: 0 0 0 2px rgba(16,185,129,0.12) !important; outline: none !important;
}
.stTextInput > div > div > input::placeholder,
.stTextArea  > div > div > textarea::placeholder { color: var(--text-3) !important; }
.stTextInput label, .stTextArea label { color: var(--text-3) !important; font-size: 12px !important; }

.main .stButton > button,
[data-testid="stMain"] .stButton > button {
  font-family: 'Inter', sans-serif !important; border-radius: var(--r-md) !important;
  font-size: 13px !important; font-weight: 500 !important; transition: all 0.17s ease !important;
  border: 1px solid var(--border) !important; background: #D1FAE5 !important;
  color: #14532D !important; padding: 7px 13px !important;
}
.main .stButton > button:hover,
[data-testid="stMain"] .stButton > button:hover {
  background: var(--bg-hover) !important; border-color: var(--border-p) !important;
  color: var(--text-1) !important; transform: translateY(-1px) !important;
}
.main .stButton > button[kind="primary"],
[data-testid="stMain"] .stButton > button[kind="primary"],
.stFormSubmitButton > button {
  background: #16835D!important;
  border: none !important; color: #fff !important;
  box-shadow: 0 2px 12px var(--purple-glow) !important;
}
.main .stButton > button[kind="primary"]:hover,
[data-testid="stMain"] .stButton > button[kind="primary"]:hover,
.stFormSubmitButton > button:hover {
  background: linear-gradient(135deg, #087E8B, #10B981) !important;
  box-shadow: 0 4px 18px rgba(16,185,129,0.5) !important;
}

[data-testid="stFileUploader"] {
  background: var(--bg-input) !important;
  border: 2px dashed rgba(16,185,129,0.3) !important;
  border-radius: var(--r-xl) !important;
}
[data-testid="stFileUploader"]:hover { border-color: rgba(16,185,129,0.6) !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { color: var(--text-2) !important; }
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }

.stSelectbox > div > div {
  background: var(--bg-input) !important; color: var(--text-1) !important;
  border: 1px solid var(--border) !important; border-radius: var(--r-md) !important;
}
.stSelectbox label { color: var(--text-3) !important; font-size: 12px !important; }

.streamlit-expanderHeader {
  background: var(--bg-card) !important; border: 1px solid var(--border) !important;
  border-radius: var(--r-md) !important; color: var(--text-2) !important; font-size: 13px !important;
}
.streamlit-expanderContent {
  background: var(--bg) !important; border: 1px solid var(--border) !important;
  border-top: none !important; border-radius: 0 0 var(--r-md) var(--r-md) !important;
}

[data-testid="metric-container"] {
  background: var(--bg-card) !important; border: 1px solid var(--border) !important;
  border-radius: var(--r-lg) !important; padding: 15px !important; box-shadow: var(--sh-card) !important;
}
[data-testid="metric-container"] label { color: var(--text-3) !important; font-size: 11px !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] {
  color: var(--text-1) !important; font-size: 23px !important; font-weight: 700 !important;
}

[data-testid="stDataFrame"]  { border: 1px solid var(--border) !important; border-radius: var(--r-md) !important; }
[data-testid="stForm"]       { background: transparent !important; border: none !important; padding: 0 !important; }
[data-testid="stAlert"]      { border-radius: var(--r-md) !important; font-size: 13px !important; }
.stSpinner > div { border-top-color: var(--purple) !important; }

.stTabs [data-baseweb="tab-list"] { background: transparent !important; border-bottom: 1px solid var(--border) !important; }
.stTabs [data-baseweb="tab"] { background: transparent !important; border: none !important;
  color: var(--text-3) !important; font-size: 13px !important; border-radius: 6px 6px 0 0 !important; }
.stTabs [aria-selected="true"] { background: var(--bg-card) !important; color: #6EE7B7 !important;
  border-bottom: 2px solid var(--purple) !important; }
.stTabs [data-baseweb="tab-panel"] { background: transparent !important; padding: 14px 0 !important; }

code { background: rgba(16,185,129,0.12) !important; color: #A7F3D0 !important;
  border: 1px solid rgba(16,185,129,0.25) !important; border-radius: 4px !important;
  padding: 1px 5px !important; font-size: 12.5px !important; }
pre  { background: var(--bg-elevated) !important; border: 1px solid var(--border) !important;
  border-radius: var(--r-md) !important; padding: 13px !important; }

/* Empty state */
.es { text-align: center; padding: 60px 24px; }
.es-icon  { font-size: 42px; margin-bottom: 13px; }
.es-title { font-size: 16px !important; font-weight: 600 !important; color: var(--text-3) !important; }
.es-sub   { font-size: 12.5px !important; color: var(--text-4) !important; margin-top: 5px; }


<style>
:root {
    --bg: #F5F8F5;
    --bg-elevated: #FFFFFF;
    --surface: #FFFFFF;
    --surface-2: #EAF5ED;
    --text: #24352D;
    --text-2: #52665A;
    --border: #DCE7DE;
    --accent: #16835D;
    --accent-hover: #116B4B;
}

.stApp {
    background: var(--bg) !important;
    color: var(--text) !important;
}

.stApp p,
.stApp label,
.stApp span,
.stApp li,
.stApp h1,
.stApp h2,
.stApp h3,
.stApp h4 {
    color: var(--text);
}

section[data-testid="stSidebar"] {
    background: #EAF5ED !important;
    border-right: 1px solid var(--border);
}

.stApp div[data-testid="stVerticalBlockBorderWrapper"],
.stApp div[data-testid="stExpander"],
.stApp div[data-testid="stFileUploader"] {
    background: #FFFFFF !important;
    border-color: var(--border) !important;
    border-radius: 14px !important;
}

.stApp input,
.stApp textarea {
    background: #FFFFFF !important;
    color: var(--text) !important;
    border-color: #C9D9CD !important;
}

.stApp input::placeholder,
.stApp textarea::placeholder {
    color: #718276 !important;
    opacity: 1 !important;
}

.stApp button {
    border-radius: 10px !important;
    transition: background 0.2s ease, transform 0.2s ease;
}

.stApp div[data-testid="stChatMessage"] {
    background: #FFFFFF !important;
    color: var(--text) !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
    padding: 16px !important;
}

.stApp div[data-testid="stChatInput"] textarea {
    background: #FFFFFF !important;
    color: var(--text) !important;
}

.stApp div[data-testid="stFileUploader"] section {
    background: #F8FBF8 !important;
    border: 1px dashed #A8C8B2 !important;
    border-radius: 12px !important;
}

.stApp div[data-testid="stAlert"] {
    border-radius: 12px !important;
}

.stApp table {
    background: #FFFFFF !important;
    color: var(--text) !important;
}

.stApp th {
    background: #EAF5ED !important;
    color: var(--text) !important;
}

.stApp td {
    color: var(--text-2) !important;
    border-color: var(--border) !important;
}

.stApp pre,
.stApp code {
    color: var(--text) !important;
}

.stApp a {
    color: var(--accent) !important;
}

.stApp,
section[data-testid="stSidebar"] {
    transition: background 0.2s ease;
}
</style>

""", unsafe_allow_html=True)

theme_tokens = {
    "--bg": "#0B1415" if st.session_state.dark_mode else "#F4F8F7",
    "--bg-panel": "#101E1F" if st.session_state.dark_mode else "#E9F5F1",
    "--bg-elevated": "#142526" if st.session_state.dark_mode else "#FFFFFF",
    "--bg-card": "#162829" if st.session_state.dark_mode else "#FFFFFF",
    "--bg-hover": "#203735" if st.session_state.dark_mode else "#E8F5F1",
    "--bg-input": "#142526" if st.session_state.dark_mode else "#FFFFFF",
    "--surface": "#162829" if st.session_state.dark_mode else "#FFFFFF",
    "--surface-2": "#101E1F" if st.session_state.dark_mode else "#E9F5F1",
    "--text": "#E8F2EF" if st.session_state.dark_mode else "#193B35",
    "--text-1": "#E8F2EF" if st.session_state.dark_mode else "#193B35",
    "--text-2": "#BCCDC8" if st.session_state.dark_mode else "#3F5B55",
    "--text-3": "#9AAFA8" if st.session_state.dark_mode else "#536E67",
    "--text-4": "#829A92" if st.session_state.dark_mode else "#607873",
    "--border": "#29413E" if st.session_state.dark_mode else "#D0E2DC",
    "--border-p": "rgba(52,211,153,.36)" if st.session_state.dark_mode else "rgba(8,124,88,.30)",
    "--border-c": "rgba(103,232,249,.32)" if st.session_state.dark_mode else "rgba(15,118,110,.28)",
    "--purple": "#34D399" if st.session_state.dark_mode else "#087C58",
    "--purple-dim": "rgba(52,211,153,.14)" if st.session_state.dark_mode else "rgba(8,124,88,.10)",
    "--purple-glow": "rgba(52,211,153,.18)" if st.session_state.dark_mode else "rgba(8,124,88,.14)",
    "--cyan": "#67E8F9" if st.session_state.dark_mode else "#0F766E",
    "--cyan-dim": "rgba(103,232,249,.12)" if st.session_state.dark_mode else "rgba(15,118,110,.10)",
    "--green": "#34D399" if st.session_state.dark_mode else "#087C58",
    "--green-dim": "rgba(52,211,153,.14)" if st.session_state.dark_mode else "rgba(8,124,88,.10)",
    "--amber": "#FBBF24" if st.session_state.dark_mode else "#A16207",
    "--red": "#F87171" if st.session_state.dark_mode else "#B42318",
    "--accent": "#34D399" if st.session_state.dark_mode else "#087C58",
    "--accent-hover": "#6EE7B7" if st.session_state.dark_mode else "#075F44",
    "--heading-start": "#FFFFFF" if st.session_state.dark_mode else "#193B35",
    "--heading-mid": "#6EE7B7" if st.session_state.dark_mode else "#087C58",
    "--heading-end": "#67E8F9" if st.session_state.dark_mode else "#0F766E",
    "--r-md": "10px",
    "--r-lg": "14px",
    "--r-xl": "18px",
    "--sh-card": "0 8px 28px rgba(0,0,0,.22)" if st.session_state.dark_mode else "0 8px 28px rgba(22,80,63,.07)",
    "--sh-sm": "0 3px 10px rgba(0,0,0,.18)" if st.session_state.dark_mode else "0 3px 10px rgba(22,80,63,.07)",
    "--sh-glow": "0 4px 18px rgba(52,211,153,.14)" if st.session_state.dark_mode else "0 4px 18px rgba(8,124,88,.12)",
}
theme_css = "\n".join(f"{name}: {value};" for name, value in theme_tokens.items())
color_scheme = "dark" if st.session_state.dark_mode else "light"
st.markdown(f"""
<style>
:root {{
  color-scheme: {color_scheme};
  {theme_css}
}}
html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
  background: var(--bg) !important;
  color: var(--text-1) !important;
}}
[data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {{
  background: var(--bg-panel) !important;
  color: var(--text-1) !important;
}}
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 {{
  color: var(--text-1) !important;
}}
.stApp p, .stApp li, .stApp label, [data-testid="stMarkdownContainer"] {{
  color: var(--text-2) !important;
}}
.wlc-title {{
  background-image: linear-gradient(135deg, var(--heading-start), var(--heading-mid) 45%, var(--heading-end)) !important;
}}
.stApp input, .stApp textarea, [data-testid="stChatInput"],
[data-testid="stSelectbox"] [role="combobox"] {{
  background: var(--bg-input) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
.stApp input::placeholder, .stApp textarea::placeholder {{
  color: var(--text-3) !important;
  opacity: 1 !important;
}}
.stApp div.stButton > button:not([kind="primary"]),
.stApp div[data-testid="stDownloadButton"] > button {{
  background: var(--bg-elevated) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
.stApp div.stButton > button:not([kind="primary"]):hover,
.stApp div[data-testid="stDownloadButton"] > button:hover {{
  background: var(--bg-hover) !important;
  color: var(--text-1) !important;
}}
[data-testid="stVerticalBlockBorderWrapper"], [data-testid="stExpander"],
[data-testid="stFileUploader"], [data-testid="stChatMessage"],
[data-testid="metric-container"], .dc, .hc, .sc, .stg-card, .sdoc,
.doc-preview, .upload-panel {{
  background-color: var(--bg-card) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
.stApp div[data-testid="stVerticalBlockBorderWrapper"],
.stApp div[data-testid="stExpander"],
.stApp div[data-testid="stFileUploader"],
.stApp div[data-testid="stChatMessage"] {{
  background: var(--bg-card) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
.stApp div[data-testid="stFileUploader"] section,
.stApp [data-testid="stFileUploaderDropzone"],
[data-testid="stExpander"] details,
[data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
  background: var(--bg-elevated) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
[data-testid="stFileUploader"] button {{
  background: var(--bg-elevated) !important;
  color: var(--text-1) !important;
}}
.stApp [data-testid="stFileUploaderDropzone"] {{
  background: var(--bg-elevated) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
[data-testid="stMain"] [class*="st-key-chat_suggestion_"] [data-testid="stButton"] > button {{
  background: var(--bg-card) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
[data-testid="stMain"] [class*="st-key-chat_suggestion_"] [data-testid="stButton"] > button:hover {{
  background: var(--bg-hover) !important;
}}
.sp {{ background: var(--bg-panel) !important; color: var(--text-1) !important; }}
.sp-hdr {{ background: var(--bg-card) !important; }}
.app-hdr, .docs-toolbar, .doc-preview, .doc-empty, .composer {{
  background: var(--bg-panel) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
.doc-card:hover {{ background: var(--bg-hover) !important; }}
.stg-val {{ color: var(--accent) !important; }}
.sc-file {{ color: var(--cyan) !important; }}
.sc-num {{ color: var(--accent) !important; }}
.sb-stat {{ background: var(--bg-card) !important; }}
.sb-sub, .sb-status-label, .sb-model, .sb-profile-role, .sb-empty-docs,
.doc-card-meta, .docs-subtitle, .upload-desc, .upload-meta {{
  color: var(--text-3) !important;
}}
.msg-a-body h1, .msg-a-body h2, [data-testid="stChatMessageContent"] h1,
[data-testid="stChatMessageContent"] h2, .msg-a-body code, code,
.msg-a-body th, [data-testid="stChatMessageContent"] th {{
  color: var(--text-1) !important;
}}
.msg-a-body pre, pre, .sc-full {{
  background: var(--bg-elevated) !important;
  color: var(--text-1) !important;
  border-color: var(--border) !important;
}}
[data-testid="stMain"] .copy-widget button {{
  color: var(--text-2) !important;
}}
[data-testid="stMain"] .copy-widget button:hover {{
  color: var(--text-1) !important;
  background: var(--bg-hover) !important;
  border-color: var(--border) !important;
}}
.stTabs [aria-selected="true"] {{ color: var(--accent) !important; }}
::-webkit-scrollbar-thumb {{ background: var(--border) !important; }}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  BACKEND — OLLAMA + RAG
# ═══════════════════════════════════════════════════════════════════════════════
OLLAMA_URL             = (os.getenv("OLLAMA_URL") or "http://localhost:11434").rstrip("/")
OLLAMA_GENERATION_MODEL = os.getenv("OLLAMA_GENERATION_MODEL", "llama3.2:1b")
OLLAMA_EMBEDDING_MODEL = os.getenv("OLLAMA_EMBEDDING_MODEL", "nomic-embed-text")
OLLAMA_GENERATE        = f"{OLLAMA_URL}/api/generate"
OLLAMA_TIMEOUT         = int(os.getenv("OLLAMA_TIMEOUT_SECONDS", "15"))

local_vectors   = {}
embedding_cache = {}

def ocr_image(p):   return pytesseract.image_to_string(Image.open(p))
def load_pdf(p):    return PyPDFLoader(p).load()
def load_txt(p):    return [Document(page_content=open(p, encoding="utf-8").read())]
def load_docx(p):
    d = DocxDocument(p)
    return [Document(page_content="\n".join(x.text for x in d.paragraphs))]
def load_excel(p):  return pd.read_excel(p)
def to_docs_from_excel(df):
    return [Document(page_content=" ".join(map(str, row))) for _, row in df.iterrows()]
def split_text(docs, size=500, overlap=50):
    return RecursiveCharacterTextSplitter(chunk_size=size, chunk_overlap=overlap).split_documents(docs)

def ollama_request(url, payload, timeout=None, retries=2):
    timeout = timeout or OLLAMA_TIMEOUT; last = None
    for a in range(retries+1):
        try:
            r = requests.post(url, json=payload, timeout=timeout); r.raise_for_status(); return r.json()
        except RequestException as e:
            last = e
            if a < retries: time.sleep(0.5*(a+1))
    error = RuntimeError(f"Cannot reach Ollama: {last}")
    error.status_code = getattr(getattr(last, "response", None), "status_code", None)
    raise error from last

def compute_embeddings(texts):
    cleaned = [(text or "").strip() for text in texts]
    pending = list(dict.fromkeys(text for text in cleaned if text and text not in embedding_cache))
    if pending:
        try:
            result = ollama_request(
                f"{OLLAMA_URL}/api/embed",
                {"model": OLLAMA_EMBEDDING_MODEL, "input": pending},
                timeout=max(OLLAMA_TIMEOUT, 30),
                retries=0,
            )
            vectors = result.get("embeddings")
            if not vectors or len(vectors) != len(pending):
                raise RuntimeError("Ollama returned an incomplete embedding response.")
        except RuntimeError as exc:
            if getattr(exc, "status_code", None) != 404:
                raise RuntimeError(
                    f"Could not create embeddings with '{OLLAMA_EMBEDDING_MODEL}'. "
                    "Install the embedding model with: ollama pull nomic-embed-text. "
                    f"Details: {exc}"
                ) from exc
            vectors = []
            try:
                for text in pending:
                    result = ollama_request(
                        f"{OLLAMA_URL}/api/embeddings",
                        {"model": OLLAMA_EMBEDDING_MODEL, "prompt": text},
                        timeout=max(OLLAMA_TIMEOUT, 30),
                        retries=0,
                    )
                    vector = result.get("embedding")
                    if not vector:
                        raise RuntimeError("Ollama returned an empty embedding.")
                    vectors.append(vector)
            except RuntimeError as legacy_exc:
                if getattr(legacy_exc, "status_code", None) == 404:
                    detail = (
                        "This Ollama server returned 404 for both embedding API routes. "
                        "Update/restart Ollama and install the model with: "
                        "ollama pull nomic-embed-text."
                    )
                else:
                    detail = f"Install the model with: ollama pull nomic-embed-text. Details: {legacy_exc}"
                raise RuntimeError(
                    f"Could not create embeddings with '{OLLAMA_EMBEDDING_MODEL}'. "
                    f"{detail}"
                ) from legacy_exc
        embedding_cache.update(zip(pending, vectors))
    return [embedding_cache[text] for text in cleaned if text]

def ollama_cli_run(model, prompt, timeout=30):
    try:
        p = subprocess.run(["ollama","run",model,prompt], capture_output=True, text=True, timeout=timeout)
        if p.returncode != 0: raise RuntimeError(p.stderr.strip())
        return p.stdout.strip()
    except Exception as e: raise RuntimeError(f"Ollama CLI: {e}") from e

def _qdrant_up():
    import socket
    try:
        with socket.create_connection(("127.0.0.1",6333), timeout=1): return True
    except Exception: return False

def create_qdrant(name, dim):
    if not _qdrant_up():
        local_vectors[name]={"vectors":[],"texts":[],"dim":dim}
        st.session_state["qdrant_available"]=False; return
    try:
        c=QdrantClient(host="localhost",port=6333)
        try: c.recreate_collection(name, vectors_config=VectorParams(size=dim,distance=Distance.COSINE))
        except Exception:
            if not c.collection_exists(name): c.create_collection(name,vectors_config=VectorParams(size=dim,distance=Distance.COSINE))
        st.session_state.setdefault("qdrant_available",True)
    except Exception:
        local_vectors[name]={"vectors":[],"texts":[],"dim":dim}
        st.session_state["qdrant_available"]=False

def insert_qdrant(name, texts, vectors):
    if not _qdrant_up():
        lv=local_vectors.setdefault(name,{"vectors":[],"texts":[],"dim":len(vectors[0]) if vectors else 0})
        lv["vectors"].extend(vectors); lv["texts"].extend(texts); return
    try:
        c=QdrantClient(host="localhost",port=6333)
        pts=[PointStruct(id=i,vector=vectors[i],payload={"text":texts[i]}) for i in range(len(texts))]
        c.upsert(name,points=pts)
    except Exception:
        lv=local_vectors.setdefault(name,{"vectors":[],"texts":[],"dim":len(vectors[0]) if vectors else 0})
        lv["vectors"].extend(vectors); lv["texts"].extend(texts)

def search_qdrant(name, q_vector):
    if _qdrant_up():
        try:
            res=QdrantClient(host="localhost",port=6333).search(collection_name=name,query_vector=q_vector,limit=5)
            return [r.payload.get("text","") for r in res]
        except Exception: pass
    lv=local_vectors.get(name)
    if not lv or not lv.get("vectors"): return []
    vecs=np.array(lv["vectors"]); q=np.array(q_vector); eps=1e-8
    vn=vecs/(np.linalg.norm(vecs,axis=1,keepdims=True)+eps); qn=q/(np.linalg.norm(q)+eps)
    return [lv["texts"][i] for i in np.argsort(-(vn@qn))[:5]]

def get_doc_cache_key(docs):
    content = f"{OLLAMA_EMBEDDING_MODEL}\n" + "\n".join(d.page_content or "" for d in docs)
    return sha1(content.encode("utf-8","ignore")).hexdigest()

def get_or_index_document(docs):
    st.session_state.setdefault("doc_index_cache",{})
    key=get_doc_cache_key(docs)
    collection_name=f"query_{key}"
    cached=st.session_state.doc_index_cache.get(key)
    if cached:
        if cached.get("collection_name") != collection_name:
            texts=cached.get("texts",[])
            embs=cached.get("embeddings",[])
            if texts and embs:
                create_qdrant(collection_name,len(embs[0])); insert_qdrant(collection_name,texts,embs)
                cached["collection_name"]=collection_name
            else:
                cached=None
        if cached: return cached
    texts=[c.page_content for c in split_text(docs)]
    embs=compute_embeddings(texts)
    if not embs or not embs[0]: raise RuntimeError("Failed to compute embeddings.")
    create_qdrant(collection_name,len(embs[0])); insert_qdrant(collection_name,texts,embs)
    st.session_state.doc_index_cache[key]={"texts":texts,"embeddings":embs,"collection_name":collection_name}
    return st.session_state.doc_index_cache[key]

def run_rag(query, docs):
    st.session_state.setdefault("rag_query_cache",{})
    key=get_doc_cache_key(docs)
    collection_name=f"query_{key}"
    cache_key=(collection_name,query,OLLAMA_GENERATION_MODEL)
    hit=st.session_state.rag_query_cache.get(cache_key)
    if hit: return hit[0],hit[1]
    try: indexed=get_or_index_document(docs)
    except RuntimeError as e: return str(e),[]
    try:
        qemb=compute_embeddings([query])[0]
    except RuntimeError as e:
        return str(e),[]

    retrieved = search_qdrant(indexed["collection_name"], qemb)

    # Fallback to the indexed document's own chunks
    if not retrieved:
        retrieved = indexed.get("texts", [])

    context = "\n".join(
        str(item) for item in retrieved if item
    )

    if not context.strip():
        return (
            "I couldn't find relevant information in the uploaded documents. "
            "Please check that the document was indexed successfully "
            "or try a more specific question."
        ), []
    prompt=(
        "You are DocuMind AI, a helpful document assistant. "
        "Answer the user's question using the supplied document context. "
        "Be accurate, clear, and concise. "
        "For questions about skills, technologies, tools, experience, "
        "education, projects, or qualifications, extract and list every "
        "relevant item explicitly mentioned in the context. "
        "Use bullet points when appropriate. "
        "Do not invent facts. If the context does not contain the answer, "
        "say so clearly.\n\n"
        f"Document context:\n{context}\n\n"
        f"Question: {query}\n\n"
        "Answer:"
    )
    try:
        resp=ollama_request(
            OLLAMA_GENERATE,
            {
                "model": OLLAMA_GENERATION_MODEL,
                "prompt": prompt,
                "stream": False,
                "keep_alive": "10m",
                "options": {"temperature": 0.2, "num_predict": 160, "num_ctx": 2048},
            },
            timeout=60,
        )
        result=resp.get("response",resp.get("text","No response"))
    except RuntimeError as e:
        try:
            result=ollama_cli_run(OLLAMA_GENERATION_MODEL,prompt,timeout=60)
        except RuntimeError as e2:
            return f"Ollama unavailable:\n{e}\n{e2}",[]
    st.session_state.rag_query_cache[cache_key]=(result,retrieved)
    return result,retrieved

# ── Helpers ────────────────────────────────────────────────────────────────────
ACCEPTED = ["pdf","txt","docx","xlsx","png","jpg","jpeg"]
EICONS = {"pdf":"📕","txt":"📄","docx":"📘","xlsx":"📗","png":"🖼️","jpg":"🖼️","jpeg":"🖼️"}
ECOLORS= {"pdf":"rgba(239,68,68,0.18)","docx":"rgba(59,130,246,0.18)",
          "txt":"rgba(100,116,139,0.15)","xlsx":"rgba(16,185,129,0.18)",
          "png":"rgba(139,92,246,0.18)","jpg":"rgba(139,92,246,0.18)","jpeg":"rgba(139,92,246,0.18)"}

def eicon(e):  return EICONS.get(e.lower(),"📎")
def ecolor(e): return ECOLORS.get(e.lower(),"rgba(16,185,129,0.15)")
def hsize(path):
    try:
        b=os.path.getsize(path)
        for u in ["B","KB","MB"]:
            if b<1024: return f"{b:.0f} {u}"
            b/=1024
        return f"{b:.1f} GB"
    except Exception: return ""

def check_ollama():
    try: return requests.get(f"{OLLAMA_URL}/api/tags",timeout=2).status_code==200
    except Exception: return False

def process_upload(file):
    ext=file.name.rsplit(".",1)[-1].lower()
    path=os.path.join(tempfile.gettempdir(),file.name)
    with open(path,"wb") as fp: fp.write(file.read())
    content=""; df_data=None; status="ready"
    try:
        if ext in ["png","jpg","jpeg"]: content=ocr_image(path)
        elif ext=="pdf":  content="\n".join(p.page_content for p in load_pdf(path))
        elif ext=="txt":  content=load_txt(path)[0].page_content
        elif ext=="docx": content=load_docx(path)[0].page_content
        elif ext=="xlsx":
            df_data=load_excel(path)
            content="\n".join(" ".join(map(str,r)) for _,r in df_data.iterrows())
    except Exception: status="error"
    return {"name":file.name,"path":path,"text":content,"ext":ext,
            "date":datetime.now().strftime("%Y-%m-%d %H:%M"),"df":df_data,"status":status}

# ── Session init ───────────────────────────────────────────────────────────────
for k,v in {"files":[],"chat_history":[],"query_results":{},
            "doc_index_cache":{},"rag_query_cache":{},
            "active_view":"chat","selected_file_idx":None,"last_sources":[],
            "doc_upload_generation":0}.items():
    if k not in st.session_state: st.session_state[k]=v

ollama_online = check_ollama()
qdrant_online = _qdrant_up()

# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:

    # Brand
    st.markdown("""
    <div class="sb-brand">
      <div style="display:flex;align-items:center;gap:10px;">
        <div class="sb-logo" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
            <path d="M6 3.75h8l4 4v12.5H6z"/>
            <path d="M14 3.75v4h4M9 12h6M9 15.5h4"/>
            <path d="m18.6 12.4.55 1.45 1.45.55-1.45.55-.55 1.45-.55-1.45-1.45-.55 1.45-.55z" stroke="#38BDF8"/>
          </svg>
        </div>
        <div>
          <div class="sb-name">DocuMind AI</div>
          <div class="sb-sub">AI Document Assistant</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    st.toggle(
        "Dark mode",
        key="dark_mode",
        help="Switch between the dark and light DocuMind themes.",
    )

    # New chat
    if st.button("＋  New chat", key="new_chat", use_container_width=True, type="primary"):
        st.session_state.chat_history=[]; st.session_state.selected_file_idx=None
        st.session_state.last_sources=[]; st.session_state.active_view="chat"
        st.rerun()

    # Navigation
    st.markdown('<div class="sb-sec">WORKSPACE</div>', unsafe_allow_html=True)
    for vid, icon, label in [
        ("chat","◉","Chat"), ("documents","▤","Documents"),
        ("history","◷","Chat history"), ("settings","⚙","Settings"),
    ]:
        if st.button(
            f"{icon}  {label}",
            key=f"nav_{vid}",
            use_container_width=True,
            type="primary" if st.session_state.active_view == vid else "secondary",
        ):
            st.session_state.active_view=vid; st.rerun()

    # Documents
    st.markdown('<div class="sb-hr"></div>', unsafe_allow_html=True)
    doc_head, doc_add = st.columns([5, 1], gap="small")
    with doc_head:
        st.markdown('<div class="sb-sec">UPLOADED DOCUMENTS</div>', unsafe_allow_html=True)
    with doc_add:
        if st.button("＋", key="sidebar_add_docs", help="Add documents", use_container_width=True):
            st.session_state.active_view="documents"
            st.rerun()

    if st.session_state.files:
        for i,f in enumerate(st.session_state.files):
            act="sdoc-active" if st.session_state.selected_file_idx==i else ""
            filename=escape(f.get("name") or "Untitled document")
            upload_date=escape(f.get("date") or "Date unavailable")
            ready=f.get("status","ready")=="ready"
            status_class="s-green" if ready else "s-red"
            status_text="Ready" if ready else "Error"
            st.markdown(f"""
            <div class="sdoc {act}">
              <div class="sdoc-row">
                <span class="sdoc-icon" style="background:{ecolor(f['ext'])};
                  border-radius:5px;padding:3px 5px;">{eicon(f['ext'])}</span>
                <div style="flex:1;min-width:0;">
                  <div class="sdoc-name" title="{filename}">{filename}</div>
                  <div class="sdoc-date">{upload_date}</div>
                </div>
              </div>
              <div class="sdoc-status"><span class="sdot {status_class}"></span>{status_text}</div>
            </div>""", unsafe_allow_html=True)
            if st.button("Open in chat", key=f"sel_{i}", use_container_width=True):
                st.session_state.selected_file_idx=i
                st.session_state.active_view="chat"; st.rerun()
    else:
        st.markdown(
            '<div class="sb-empty-docs">No documents uploaded yet.</div>',
            unsafe_allow_html=True,
        )

    # Status
    st.markdown('<div class="sb-hr"></div>', unsafe_allow_html=True)
    o_cls = "s-green" if ollama_online else "s-red"
    q_cls = "s-green" if qdrant_online else "s-amber"
    o_lbl = "Connected" if ollama_online else "Unavailable"
    q_lbl = "Connected" if qdrant_online else "Local in-memory fallback"
    ollama_model = OLLAMA_GENERATION_MODEL
    st.markdown(f"""
    <div class="sb-stat">
      <div class="sb-stat-row"><span class="sdot {o_cls}"></span><span>Ollama</span>
        <span class="sb-status-label">{o_lbl}</span></div>
      <div class="sb-model">Configured model · {ollama_model}</div>
      <div class="sb-stat-row sb-vector-row"><span class="sdot {q_cls}"></span><span>Qdrant</span>
        <span class="sb-status-label">{q_lbl}</span></div>
    </div>""", unsafe_allow_html=True)
    st.markdown("""
    <div class="sb-profile">
      <div class="sb-avatar" aria-hidden="true">DM</div>
      <div style="min-width:0;">
        <div class="sb-profile-name">Local workspace</div>
        <div class="sb-profile-role">Document assistant</div>
      </div>
    </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN CONTENT
# ═══════════════════════════════════════════════════════════════════════════════
view = st.session_state.active_view

# ─── DOCUMENTS ────────────────────────────────────────────────────────────────
if view == "documents":
    files=st.session_state.files
    upload_limit_mb=st.get_option("server.maxUploadSize")
    st.markdown(f"""
    <div class="documents-page">
      <div class="docs-heading">
        <div>
          <h1 class="docs-title">My Documents</h1>
          <div class="docs-subtitle">Upload, organize, and chat with your documents.</div>
        </div>
        <div class="docs-count" aria-label="{len(files)} documents">
          <strong>{len(files)}</strong> {"document" if len(files)==1 else "documents"}
        </div>
      </div>
      <div class="upload-panel">
        <div class="upload-icon" aria-hidden="true">⇧</div>
        <div class="upload-copy">
          <div class="upload-title">Drop your documents here</div>
          <div class="upload-desc">Upload files to extract text and start asking questions.</div>
          <div class="upload-meta">PDF · DOCX · TXT · XLSX · PNG · JPG · JPEG &nbsp;|&nbsp; Up to {upload_limit_mb} MB per file</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    with st.container(key="documents_upload_control"):
        uploaded=st.file_uploader(
            "Browse Files",
            type=ACCEPTED,
            accept_multiple_files=True,
            key=f"doc_up_{st.session_state.doc_upload_generation}",
            help="Accepted formats: PDF, DOCX, TXT, XLSX, PNG, JPG, and JPEG.",
        )
    if uploaded:
        existing={f["name"] for f in st.session_state.files}
        for upload in uploaded:
            if upload.name in existing:
                continue
            with st.spinner(f"Extracting and indexing {upload.name}…"):
                rec=process_upload(upload)
            st.session_state.files.append(rec)
            existing.add(upload.name)
            st.toast(
                f"{'✓' if rec['status']=='ready' else '⚠'} {upload.name}",
                icon="✅" if rec["status"]=="ready" else "⚠️",
            )
        st.session_state.doc_upload_generation+=1
        st.rerun()

    if not st.session_state.files:
        st.markdown("""
        <div class="documents-page">
          <div class="doc-empty">
            <div class="doc-empty-icon">📂</div>
            <div class="doc-empty-title">No documents yet</div>
            <div class="doc-empty-copy">Choose files above to begin OCR extraction and document search.</div>
          </div>
        </div>""", unsafe_allow_html=True)
    else:
        with st.container(key="docs_toolbar"):
            search_col,type_col,status_col,sort_col,view_col=st.columns([2.2,1,1,1.2,.9],gap="small")
            with search_col:
                search_query=st.text_input(
                    "Search documents",
                    placeholder="Search filename or extracted text…",
                    key="docs_search",
                    label_visibility="collapsed",
                ).strip().casefold()
            with type_col:
                type_filter=st.selectbox(
                    "File type",
                    ["All types"]+[ext.upper() for ext in ACCEPTED],
                    key="docs_type_filter",
                    label_visibility="collapsed",
                )
            with status_col:
                status_filter=st.selectbox(
                    "Processing status",
                    ["All statuses","Completed","Failed"],
                    key="docs_status_filter",
                    label_visibility="collapsed",
                )
            with sort_col:
                sort_order=st.selectbox(
                    "Sort documents",
                    ["Newest first","Oldest first","Name A–Z"],
                    key="docs_sort",
                    label_visibility="collapsed",
                )
            with view_col:
                view_mode=st.radio(
                    "Document display",
                    ["Grid","List"],
                    horizontal=True,
                    key="docs_view_mode",
                    label_visibility="collapsed",
                )

        filtered_files=[]
        for file_index,doc_file in enumerate(st.session_state.files):
            file_status=doc_file.get("status","ready")
            if type_filter!="All types" and doc_file.get("ext","").upper()!=type_filter:
                continue
            if status_filter=="Completed" and file_status!="ready":
                continue
            if status_filter=="Failed" and file_status!="error":
                continue
            if search_query and search_query not in (
                f"{doc_file.get('name','')} {doc_file.get('text','')}".casefold()
            ):
                continue
            filtered_files.append((file_index,doc_file))

        if sort_order=="Newest first":
            filtered_files.sort(key=lambda item:item[1].get("date",""),reverse=True)
        elif sort_order=="Oldest first":
            filtered_files.sort(key=lambda item:item[1].get("date",""))
        else:
            filtered_files.sort(key=lambda item:item[1].get("name","").casefold())

        st.markdown(
            f'<div class="documents-page"><div class="docs-subtitle">'
            f'Showing {len(filtered_files)} of {len(st.session_state.files)} documents'
            f'</div></div>',
            unsafe_allow_html=True,
        )
        if not filtered_files:
            st.markdown("""
            <div class="documents-page"><div class="doc-empty">
              <div class="doc-empty-icon">⌕</div>
              <div class="doc-empty-title">No matching documents</div>
              <div class="doc-empty-copy">Try another search term or adjust your filters.</div>
            </div></div>""", unsafe_allow_html=True)
        else:
            column_count=3 if view_mode=="Grid" else 1
            for row_start in range(0,len(filtered_files),column_count):
                card_columns=st.columns(column_count,gap="medium")
                for card_column,(file_index,doc_file) in zip(
                    card_columns,filtered_files[row_start:row_start+column_count]
                ):
                    with card_column:
                        is_selected=st.session_state.selected_file_idx==file_index
                        ready=doc_file.get("status","ready")=="ready"
                        status_label="Completed" if ready else "Failed"
                        status_class="doc-status-ready" if ready else "doc-status-error"
                        safe_name=escape(doc_file.get("name","Untitled document"))
                        safe_ext=escape(doc_file.get("ext","").upper())
                        uploaded_date=escape(doc_file.get("date","Date unavailable"))
                        file_size=hsize(doc_file.get("path",""))
                        char_count=len(doc_file.get("text") or "")
                        card_class="doc-card doc-card-selected" if is_selected else "doc-card"
                        st.markdown(f"""
                        <article class="{card_class}">
                          <div class="doc-card-top">
                            <div class="doc-card-icon" style="background:{ecolor(doc_file.get('ext',''))};">
                              {eicon(doc_file.get('ext',''))}
                            </div>
                            <div style="min-width:0;flex:1">
                              <div class="doc-card-name" role="heading" aria-level="2" title="{safe_name}">{safe_name}</div>
                              <div class="doc-card-type">{safe_ext} document</div>
                            </div>
                          </div>
                          <div class="doc-status {status_class}">
                            <span class="sdot {'s-green' if ready else 's-red'}"></span>{status_label}
                          </div>
                          <div class="doc-card-meta">
                            <span>Uploaded {uploaded_date}</span>
                            {f'<span>{escape(file_size)}</span>' if file_size else ''}
                            <span>{char_count:,} extracted characters</span>
                          </div>
                          {'' if ready else '<div class="doc-error">Text extraction failed. Remove this file and upload it again to retry.</div>'}
                        </article>""",unsafe_allow_html=True)
                        with st.container(key=f"doc_actions_{file_index}"):
                            chat_col,preview_col,remove_col=st.columns(3,gap="small")
                            with chat_col:
                                if st.button(
                                    "Chat",
                                    key=f"dcht_{file_index}",
                                    use_container_width=True,
                                    type="primary",
                                    disabled=not ready,
                                ):
                                    st.session_state.selected_file_idx=file_index
                                    st.session_state.active_view="chat"
                                    st.rerun()
                            with preview_col:
                                preview_open=st.session_state.get(f"sp_{file_index}",False)
                                if st.button(
                                    "Close preview" if preview_open else "Preview",
                                    key=f"dpv_{file_index}",
                                    use_container_width=True,
                                ):
                                    st.session_state[f"sp_{file_index}"]=not preview_open
                                    st.rerun()
                            with remove_col:
                                if st.button(
                                    "Remove",
                                    key=f"drm_{file_index}",
                                    use_container_width=True,
                                ):
                                    st.session_state.files.pop(file_index)
                                    if st.session_state.selected_file_idx==file_index:
                                        st.session_state.selected_file_idx=None
                                    elif (st.session_state.selected_file_idx or 0)>file_index:
                                        st.session_state.selected_file_idx-=1
                                    st.session_state.pop(f"sp_{file_index}",None)
                                    st.rerun()

                        if st.session_state.get(f"sp_{file_index}",False):
                            st.markdown(f"""
                            <div class="doc-preview">
                              <div class="doc-preview-title">Document preview · {safe_name}</div>
                              <div class="docs-subtitle">{safe_ext} · {uploaded_date}</div>
                            </div>""",unsafe_allow_html=True)
                            if doc_file.get("ext")=="xlsx" and doc_file.get("df") is not None:
                                st.dataframe(doc_file["df"],use_container_width=True)
                            else:
                                extracted_text=doc_file.get("text") or "No text was extracted from this file."
                                st.text_area(
                                    "Extracted OCR and document text",
                                    value=extracted_text,
                                    height=260,
                                    disabled=True,
                                    key=f"doc_preview_text_{file_index}",
                                )

# ─── HISTORY ──────────────────────────────────────────────────────────────────
elif view == "history":
    st.markdown("""
    <div class="ph">
      <h1 class="ph-title">Chat History</h1>
      <div class="ph-sub">All messages from this session.</div>
    </div>""", unsafe_allow_html=True)

    if not st.session_state.chat_history:
        st.markdown("""<div class="es"><div class="es-icon">🕐</div>
        <div class="es-title">No history yet</div>
        <div class="es-sub">Conversations appear here once you start chatting.</div>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown('<div style="padding:4px 26px 0;">', unsafe_allow_html=True)
        if st.button("🗑  Clear All History", type="primary"):
            st.session_state.chat_history=[]; st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        for msg in reversed(st.session_state.chat_history):
            is_u=msg["role"]=="user"
            acc="#10B981" if is_u else "#38BDF8"
            who=(f"👤 You · {msg.get('ts','')} · {msg.get('file_name','')}" if is_u
                 else f"🧠 DocuMind AI · {msg.get('ts','')}")
            body=msg["content"] or ""
            if not is_u and len(body)>340: body=body[:340]+"…"
            st.markdown(f"""
            <div class="hc" style="border-left:3px solid {acc};">
              <div class="hc-who">{who}</div>
              <div class="hc-body">{body}</div>
            </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ─── SETTINGS ─────────────────────────────────────────────────────────────────
elif view == "settings":
    st.markdown("""
    <div class="ph">
      <h1 class="ph-title">Settings</h1>
      <div class="ph-sub">Backend connections, model config, and session management.</div>
    </div>""", unsafe_allow_html=True)
    st.markdown('<div style="padding:6px 26px 0;">', unsafe_allow_html=True)

    od="s-green" if ollama_online else "s-red"
    os_="Online" if ollama_online else "Offline"
    st.markdown(f"""
    <div class="stg-card">
      <div class="stg-title">🦙 Ollama LLM</div>
      <div class="stg-row"><span class="stg-key">Status</span>
        <span class="stg-val" style="display:flex;align-items:center;gap:7px;">
          <span class="sdot {od}"></span>{os_}</span></div>
      <div class="stg-row"><span class="stg-key">Endpoint</span>
        <span class="stg-val" style="color:#6EE7B7;">{OLLAMA_URL}</span></div>
      <div class="stg-row"><span class="stg-key">Answer model</span>
        <span class="stg-val" style="color:#A7F3D0;">{OLLAMA_GENERATION_MODEL}</span></div>
      <div class="stg-row"><span class="stg-key">Embedding model</span>
        <span class="stg-val" style="color:#7DD3FC;">{OLLAMA_EMBEDDING_MODEL}</span></div>
      <div class="stg-row"><span class="stg-key">Timeout</span>
        <span class="stg-val">{OLLAMA_TIMEOUT}s</span></div>
    </div>""", unsafe_allow_html=True)

    qd="s-green" if qdrant_online else "s-amber"
    qs="Online (localhost:6333)" if qdrant_online else "Offline — in-memory fallback"
    st.markdown(f"""
    <div class="stg-card">
      <div class="stg-title">🗄️ Qdrant Vector Store</div>
      <div class="stg-row"><span class="stg-key">Status</span>
        <span class="stg-val" style="display:flex;align-items:center;gap:7px;">
          <span class="sdot {qd}"></span>{qs}</span></div>
      <div class="stg-row"><span class="stg-key">Collection</span>
        <span class="stg-val">query_collection</span></div>
      <div class="stg-row"><span class="stg-key">Metric</span>
        <span class="stg-val">Cosine similarity</span></div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="stg-card"><div class="stg-title">📊 Session</div>', unsafe_allow_html=True)
    s1,s2,s3=st.columns(3)
    s1.metric("Documents", len(st.session_state.files))
    s2.metric("Messages",  len(st.session_state.chat_history))
    s3.metric("Cached",    len(st.session_state.rag_query_cache))
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🗑  Clear Cache", use_container_width=True):
        st.session_state.doc_index_cache={}; st.session_state.rag_query_cache={}
        embedding_cache.clear(); st.toast("Cache cleared ✓",icon="✅"); st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ─── CHAT ─────────────────────────────────────────────────────────────────────
else:
    chat_col, ctx_col = st.columns([2.25, 1], gap="small")
    with ctx_col:
        source_panel=st.empty()

    with chat_col:
        # Header
        idx = st.session_state.selected_file_idx
        if idx is not None and idx < len(st.session_state.files):
            sf=st.session_state.files[idx]
            hi=eicon(sf['ext']); hbg=ecolor(sf['ext'])
            ht=escape(sf["name"])
            file_size=hsize(sf["path"])
            hs=escape(f"{sf['ext'].upper()} · {file_size}" if file_size else sf["ext"].upper())
            hb='<span class="doc-ready"><span class="sdot s-green"></span>Document ready</span>' \
                if sf.get("status","ready")=="ready" else '<span class="badge bdg-red">Processing error</span>'
        else:
            hi="🧠"; hbg="rgba(16,185,129,0.15)"
            ht="DocuMind AI"; hs="Select or upload a document to begin"; hb=""

        od_h="s-green" if ollama_online else "s-red"
        ol_h="Ollama ready" if ollama_online else "Ollama offline"
        n_src=len(st.session_state.last_sources)
        src_badge=f'<span class="badge bdg-cyan" style="margin-left:8px;">⚡ {n_src} sources</span>'

        st.markdown(f"""
        <div class="app-hdr">
          <div class="ah-l">
            <div class="ah-ficon" style="background:{hbg};">{hi}</div>
            <div>
              <div class="ah-title" role="heading" aria-level="1">{ht}&nbsp;{hb}{src_badge}</div>
              <div class="ah-sub">{hs}</div>
            </div>
          </div>
          <div class="ah-r">
            <span class="sdot {od_h}"></span>&nbsp;{ol_h}
          </div>
        </div>
        <div class="hdr-glow"></div>""", unsafe_allow_html=True)

        # Welcome
        suggested_query = None
        if not st.session_state.chat_history:
            st.markdown("""
            <div class="wlc-wrap">
              <div class="wlc-orb">🧠</div>
              <h1 class="wlc-title">Your documents. One intelligent conversation.</h1>
              <div class="wlc-sub">Upload your files, ask questions, and uncover insights with
                AI-powered document search.</div>
            </div>""", unsafe_allow_html=True)

            if not st.session_state.files:
                st.markdown('<div style="display:flex;justify-content:center;padding:0 50px 18px;">', unsafe_allow_html=True)
                qu=st.file_uploader("📎  Add documents · PDF · DOCX · TXT · XLSX · PNG · JPG",
                                    type=ACCEPTED, accept_multiple_files=True,
                                    key=f"welcome_up_{st.session_state.doc_upload_generation}")
                st.markdown("</div>", unsafe_allow_html=True)
                if qu:
                    existing=[f["name"] for f in st.session_state.files]
                    for file in qu:
                        if file.name in existing: continue
                        with st.spinner(f"Processing {file.name}…"):
                            rec=process_upload(file)
                        st.session_state.files.append(rec)
                    st.session_state.selected_file_idx=0
                    st.session_state.doc_upload_generation+=1
                    st.rerun()
            st.markdown('<div style="display:flex;justify-content:center;padding:0 24px 18px;">', unsafe_allow_html=True)
            prompt_columns = st.columns(2, gap="small")
            for prompt_idx, (icon, title, desc) in enumerate([
                ("📋", "Summarize this document.", "Get a concise overview of the key ideas."),
                ("🔑", "Extract key skills and facts.", "Find important details in the text."),
                ("📅", "Find important dates.", "Locate dates and events in your files."),
                ("🔍", "Compare information across documents.", "Find similarities and differences."),
            ]):
                with prompt_columns[prompt_idx % 2]:
                    st.markdown(
                        f'<div class="sug-desc" style="padding:0 4px 5px;">{escape(desc)}</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown('<div class="suggestion-button">', unsafe_allow_html=True)
                    if st.button(
                        f"{icon}  {title}",
                        key=f"chat_suggestion_{prompt_idx}",
                        use_container_width=True,
                    ):
                        suggested_query = title
                    st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # Messages
        else:
            for message_idx, msg in enumerate(st.session_state.chat_history):
                if msg["role"]=="user":
                    meta="  ·  ".join(filter(None,[msg.get("ts",""),msg.get("file_name","")]))
                    body=escape(msg.get("content") or "").replace("\n","<br>")
                    st.markdown(f"""
                    <div class="msg-u-row">
                      <div class="msg-u-inner">
                        <div class="msg-u-bub">{body}</div>
                        <div class="msg-u-meta">{escape(meta)}</div>
                      </div>
                    </div>""", unsafe_allow_html=True)
                else:
                    with st.chat_message("assistant", avatar="🧠"):
                        st.markdown(
                            '<div class="chat-ai-label">DOCUMIND AI</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown(msg.get("content") or "")
                        sources=msg.get("sources",[])
                        if sources:
                            chips="".join(
                                f'<span class="src-chip">📄 Chunk {source_idx+1}</span>'
                                for source_idx in range(len(sources))
                            )
                            st.markdown(
                                f'<div class="src-row"><div class="src-lbl">Retrieved Sources</div>{chips}</div>',
                                unsafe_allow_html=True,
                            )
                        st.markdown(
                            f'<div class="msg-a-meta">{escape(msg.get("ts",""))}</div>',
                            unsafe_allow_html=True,
                        )

                        query=msg.get("query")
                        if not query:
                            query=next(
                                (prior.get("content","") for prior in reversed(
                                    st.session_state.chat_history[:message_idx]
                                ) if prior.get("role")=="user"),
                                "",
                            )
                        feedback=st.session_state.setdefault("chat_feedback",{})
                        feedback_value=feedback.get(str(message_idx))
                        copy_text=json.dumps(msg.get("content") or "").replace("<","\\u003c")
                        action_cols=st.columns([1.1,0.65,0.65,1.2,3.4], gap="small")
                        with action_cols[0]:
                            st.html(f"""
                              <style>
                                .copy-widget {{ display:flex; align-items:center; height:34px; }}
                                .copy-widget button {{ display:flex; align-items:center; gap:6px; height:30px;
                                  padding:0 9px; color:#A8B6CD; font:500 11px Inter, sans-serif;
                                  background:transparent; border:1px solid transparent;
                                  border-radius:8px; cursor:pointer; }}
                                .copy-widget button:hover {{ color:#F8FAFC; background:#2B5559;
                                  border-color:rgba(174,205,204,.18); }}
                                .copy-widget button:focus-visible {{ outline:2px solid #38BDF8;
                                  outline-offset:2px; }}
                                .copy-widget svg {{ width:13px; height:13px; }}
                              </style>
                              <div class="copy-widget">
                                <button type="button" aria-label="Copy response" title="Copy response"
                                  onclick="copyResponse(this)">
                                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                    stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"
                                    aria-hidden="true"><rect x="8" y="8" width="12" height="12" rx="2"/>
                                    <path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/></svg>
                                  <span>Copy</span>
                                </button>
                              </div>
                              <script>
                                const responseText = {copy_text};
                                async function copyResponse(button) {{
                                  try {{
                                    await navigator.clipboard.writeText(responseText);
                                  }} catch (error) {{
                                    const temporary = document.createElement("textarea");
                                    temporary.value = responseText;
                                    temporary.style.position = "fixed";
                                    temporary.style.opacity = "0";
                                    document.body.appendChild(temporary);
                                    temporary.select();
                                    const copied = document.execCommand("copy");
                                    temporary.remove();
                                    if (!copied) {{
                                      button.querySelector("span").textContent = "Select to copy";
                                      return;
                                    }}
                                  }}
                                  button.querySelector("span").textContent = "Copied";
                                }}
                              </script>
                            """, unsafe_allow_javascript=True)
                        with action_cols[1]:
                            if st.button(
                                "👍", key=f"feedback_up_{message_idx}",
                                help="Mark this response as helpful",
                                type="primary" if feedback_value=="up" else "secondary",
                            ):
                                feedback[str(message_idx)]="up"
                        with action_cols[2]:
                            if st.button(
                                "👎", key=f"feedback_down_{message_idx}",
                                help="Mark this response as not helpful",
                                type="primary" if feedback_value=="down" else "secondary",
                            ):
                                feedback[str(message_idx)]="down"
                        with action_cols[3]:
                            regenerate=st.button(
                                "↻  Regenerate", key=f"regenerate_{message_idx}",
                                help="Generate a new response to the same question",
                                use_container_width=True,
                            )
                        if feedback_value:
                            st.markdown(
                                '<div class="feedback-note">Thanks for your feedback.</div>',
                                unsafe_allow_html=True,
                            )
                        if regenerate:
                            source_file=next(
                                (item for item in st.session_state.files
                                 if item.get("name")==msg.get("file_name")),
                                None,
                            )
                            if source_file is None or not query:
                                st.warning("The original document or question is no longer available.")
                            else:
                                source_ext=source_file["ext"]
                                if source_ext=="pdf":
                                    regen_docs=load_pdf(source_file["path"])
                                elif source_ext=="txt":
                                    regen_docs=load_txt(source_file["path"])
                                elif source_ext=="docx":
                                    regen_docs=load_docx(source_file["path"])
                                elif source_ext=="xlsx":
                                    regen_docs=to_docs_from_excel(source_file["df"])
                                elif source_ext in ["png","jpg","jpeg"]:
                                    regen_docs=[Document(page_content=source_file["text"])]
                                else:
                                    regen_docs=[]
                                if regen_docs:
                                    regen_key=(f"query_{get_doc_cache_key(regen_docs)}",query)
                                    st.session_state.rag_query_cache.pop(regen_key,None)
                                    with st.spinner("Regenerating response…"):
                                        new_answer,new_sources=run_rag(query,regen_docs)
                                    msg["content"]=new_answer
                                    msg["sources"]=new_sources
                                    msg["ts"]=datetime.now().strftime("%H:%M")
                                    st.session_state.last_sources=new_sources
                                    st.rerun()

        # Doc selector
        if st.session_state.files:
            st.markdown('<div style="padding:4px 24px 0;">', unsafe_allow_html=True)
            names=[f["name"] for f in st.session_state.files]
            ci=st.session_state.selected_file_idx or 0
            chosen=st.selectbox("Active document",names,index=min(ci,len(names)-1),key="doc_sel")
            ni=names.index(chosen)
            if ni!=st.session_state.selected_file_idx:
                st.session_state.selected_file_idx=ni; st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # Streamlit's native chat input submits on Enter and supports Shift+Enter for newlines.
        st.markdown('<div class="composer-hint">Answers are generated from your documents using RAG + OCR.</div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="chat-input-wrap">', unsafe_allow_html=True)
        chat_input=st.chat_input(
            "Ask anything about your documents...",
            key="chat_composer",
        )
        st.markdown("</div>", unsafe_allow_html=True)
        qin=suggested_query or chat_input

        # Process
        if qin:
            si=st.session_state.selected_file_idx
            if si is None or si>=len(st.session_state.files):
                st.warning("⚠ Select or upload a document first.")
            else:
                file=st.session_state.files[si]; ts=datetime.now().strftime("%H:%M"); ext=file["ext"]
                docs=None
                if ext=="pdf":               docs=load_pdf(file["path"])
                elif ext=="txt":             docs=load_txt(file["path"])
                elif ext=="docx":            docs=load_docx(file["path"])
                elif ext=="xlsx":            docs=to_docs_from_excel(file["df"])
                elif ext in ["png","jpg","jpeg"]: docs=[Document(page_content=file["text"])]
                else: st.warning("Unsupported type.")
                if docs:
                    st.session_state.chat_history.append(
                        {"role":"user","content":qin,"file_name":file["name"],"ts":ts,"sources":[]})
                    source_panel.markdown("""
                    <aside class="sp" aria-label="Retrieved sources" aria-live="polite">
                      <div class="sp-hdr">
                        <div class="sp-title">
                          <svg class="sp-layers-icon" viewBox="0 0 24 24" fill="none" aria-hidden="true"
                            stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
                            <path d="m12 3 9 5-9 5-9-5 9-5Z"/>
                            <path d="m3 12 9 5 9-5M3 16l9 5 9-5"/>
                          </svg>
                          Retrieved Sources
                        </div>
                        <span class="sp-count">…</span>
                      </div>
                      <div class="sp-loading" role="status">
                        <span class="sp-loading-dot"></span>
                        Retrieving passages for this question…
                      </div>
                    </aside>""", unsafe_allow_html=True)
                    with st.spinner("🧠 Retrieving passages and generating answer…"):
                        answer,sources=run_rag(qin,docs)
                    st.session_state.chat_history.append(
                        {"role":"assistant","content":answer,"file_name":file["name"],
                         "query":qin,"ts":datetime.now().strftime("%H:%M"),"sources":sources})
                    st.session_state.last_sources=sources; st.rerun()

    # Sources panel
    with source_panel.container():
        srcs=st.session_state.last_sources; n=len(srcs)
        source_owner=next(
            (message for message in reversed(st.session_state.chat_history)
             if message.get("role")=="assistant" and message.get("sources",[])==srcs),
            None,
        )
        fname=source_owner.get("file_name","") if source_owner else ""
        if not fname and st.session_state.selected_file_idx is not None:
            selected_idx=st.session_state.selected_file_idx
            if selected_idx<len(st.session_state.files):
                fname=st.session_state.files[selected_idx].get("name","")
        safe_fname=escape(fname or "Selected document")
        latest_answer=next(
            (message.get("content","") for message in reversed(st.session_state.chat_history)
             if message.get("role")=="assistant"),
            "",
        )
        retrieval_failed=not srcs and latest_answer.startswith((
            "Ollama unavailable:",
            "Embeddings unavailable.",
            "Failed to compute embeddings.",
            "Cannot reach Ollama:",
        ))

        source_cards=[]
        for j,snip in enumerate(srcs):
            full_text=snip or ""
            has_more=len(full_text)>360
            preview=full_text[:360] + ("…" if has_more else "")
            fi_ext = fname.rsplit(".",1)[-1] if "." in fname else "txt"
            excerpt_html=escape(preview).replace("\n","<br>")
            if has_more:
                content_html=f"""
                  <details class="sc-details">
                    <summary>
                      <span class="sc-preview">{excerpt_html}</span>
                      <span class="sc-expand">Expand passage</span>
                    </summary>
                    <div class="sc-full">{escape(full_text).replace(chr(10),"<br>")}</div>
                  </details>"""
            else:
                content_html=f'<div class="sc-text">{escape(full_text).replace(chr(10),"<br>")}</div>'
            source_cards.append(f"""
              <article class="sc">
                <div class="sc-row">
                  <div class="sc-num" aria-label="Retrieved result {j+1}">{j+1}</div>
                  <div class="sc-body">
                    <div class="sc-file">
                      <span aria-hidden="true">{eicon(fi_ext)}</span>
                      <span class="sc-file-name" title="{safe_fname}">{safe_fname}</span>
                    </div>
                    <div class="sc-id">Retrieved result {j+1}</div>
                    {content_html}
                  </div>
                </div>
              </article>""")

        if srcs:
            panel_content="".join(source_cards)
            panel_footer="Passages are shown in retrieval order. Chunk IDs, page numbers, and relevance scores are not provided in the current response."
        elif retrieval_failed:
            panel_content='<div class="sp-empty"><span class="sp-empty-title">Retrieval unavailable</span>The latest answer reports a retrieval or generation error. Check the response in the chat for details.</div>'
            panel_footer="No source passages were returned for the latest response."
        elif latest_answer:
            panel_content='<div class="sp-empty"><span class="sp-empty-title">No passages found</span>The latest response did not include retrieved source passages.</div>'
            panel_footer="No source passages were returned for the latest response."
        else:
            panel_content='<div class="sp-empty"><span class="sp-empty-title">Sources will appear here</span>Ask a question about an uploaded document to inspect the passages used by the RAG pipeline.</div>'
            panel_footer="Source details appear here after a response is generated."

        st.markdown(f"""
        <aside class="sp" aria-label="Retrieved sources">
          <div class="sp-hdr">
            <div class="sp-title">
              <svg class="sp-layers-icon" viewBox="0 0 24 24" fill="none" aria-hidden="true"
                stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
                <path d="m12 3 9 5-9 5-9-5 9-5Z"/>
                <path d="m3 12 9 5 9-5M3 16l9 5 9-5"/>
              </svg>
              Retrieved Sources
            </div>
            <span class="sp-count" aria-label="{n} retrieved sources">{n}</span>
          </div>
          <div class="sp-body">{panel_content}</div>
          <div class="sp-footer">{panel_footer}</div>
        </aside>""", unsafe_allow_html=True)
