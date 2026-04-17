"""
PDF 工程整合系統 — Streamlit Edition
將 V03 桌面程式 (win32com + tkinter) 重寫為可在 Linux 雲端運行的 Streamlit 網頁程式。
使用 LibreOffice CLI 取代 win32com.client 進行 Word → PDF 轉換。
"""

import io
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
from pypdf import PdfReader, PdfWriter

# ──────────────────────────────────────────────────────────────────────────────
# Page Config  ← 必須是第一個 Streamlit 呼叫
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="⚡ PDF 工程整合系統",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────────────────────
# Cyberpunk CSS Theme
# ──────────────────────────────────────────────────────────────────────────────
CYBER_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Share+Tech+Mono&family=Rajdhani:wght@300;400;500;700&display=swap');

/* ── Variables ─────────────────────────────────────────── */
:root {
    --bg0:      #020209;
    --bg1:      #06061a;
    --bg2:      #0a0a22;
    --card:     #0d0d2a;
    --cyan:     #00f5d4;
    --cyan2:    #00b8ff;
    --magenta:  #ff2d78;
    --yellow:   #f0e000;
    --green:    #00e87a;
    --text:     #c8d8f0;
    --dim:      #4a6080;
    --border:   rgba(0,245,212,0.18);
    --border2:  rgba(255,45,120,0.35);
    --glow-c:   rgba(0,245,212,0.35);
    --glow-m:   rgba(255,45,120,0.35);
}

/* ── Global ─────────────────────────────────────────────── */
* { box-sizing: border-box; }

.stApp {
    background-color: var(--bg0) !important;
    background-image:
        repeating-linear-gradient(
            0deg, transparent, transparent 3px,
            rgba(0,245,212,0.012) 3px, rgba(0,245,212,0.012) 4px
        );
    color: var(--text) !important;
    font-family: 'Rajdhani', sans-serif !important;
    font-size: 1rem;
}

/* hide default chrome */
#MainMenu, footer, header { visibility: hidden !important; }

.block-container {
    padding-top: 1rem !important;
    max-width: 1100px !important;
}

/* ── Header ──────────────────────────────────────────────── */
.cyber-header {
    font-family: 'Orbitron', monospace;
    font-size: 2.2rem;
    font-weight: 900;
    text-align: center;
    letter-spacing: 0.25em;
    padding: 1.2rem 0 0.2rem;
    background: linear-gradient(90deg, var(--cyan), var(--magenta), var(--cyan2), var(--magenta), var(--cyan));
    background-size: 300% auto;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    animation: shimmer 4s linear infinite;
    filter: drop-shadow(0 0 30px rgba(0,245,212,0.3));
}

.cyber-subtitle {
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.72rem;
    color: var(--dim);
    text-align: center;
    letter-spacing: 0.35em;
    text-transform: uppercase;
    margin-bottom: 1.8rem;
}

@keyframes shimmer {
    to { background-position: 300% center; }
}

/* ── Tabs ────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--bg1) !important;
    border-bottom: 1px solid var(--border) !important;
    gap: 0 !important;
}

.stTabs [data-baseweb="tab"] {
    font-family: 'Orbitron', monospace !important;
    font-size: 0.72rem !important;
    letter-spacing: 0.12em !important;
    color: var(--dim) !important;
    padding: 0.85rem 2rem !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
    transition: all 0.3s !important;
}

.stTabs [aria-selected="true"] {
    color: var(--cyan) !important;
    border-bottom: 2px solid var(--cyan) !important;
    background: rgba(0,245,212,0.06) !important;
    text-shadow: 0 0 18px var(--cyan) !important;
}

.stTabs [data-baseweb="tab-panel"] {
    background: transparent !important;
    padding-top: 1.5rem !important;
}

/* ── Buttons ──────────────────────────────────────────────── */
.stButton > button {
    font-family: 'Orbitron', monospace !important;
    font-size: 0.68rem !important;
    letter-spacing: 0.1em !important;
    background: transparent !important;
    border: 1px solid var(--cyan) !important;
    color: var(--cyan) !important;
    border-radius: 2px !important;
    transition: all 0.25s !important;
    box-shadow: inset 0 0 12px rgba(0,245,212,0.08) !important;
    padding: 0.4rem 0.8rem !important;
}

.stButton > button:hover:not(:disabled) {
    background: rgba(0,245,212,0.12) !important;
    box-shadow: 0 0 20px var(--glow-c), inset 0 0 15px rgba(0,245,212,0.1) !important;
    color: #fff !important;
}

.stButton > button:disabled {
    opacity: 0.25 !important;
    cursor: not-allowed !important;
}

/* Primary (execute) button */
.stButton > button[kind="primary"] {
    border-color: var(--magenta) !important;
    color: var(--magenta) !important;
    font-size: 0.78rem !important;
    padding: 0.65rem 1rem !important;
    box-shadow: inset 0 0 20px rgba(255,45,120,0.08) !important;
    letter-spacing: 0.2em !important;
}

.stButton > button[kind="primary"]:hover:not(:disabled) {
    background: rgba(255,45,120,0.12) !important;
    box-shadow: 0 0 25px var(--glow-m), inset 0 0 20px rgba(255,45,120,0.1) !important;
    color: #fff !important;
}

/* Download button */
.stDownloadButton > button {
    font-family: 'Orbitron', monospace !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.15em !important;
    background: rgba(0,232,122,0.06) !important;
    border: 1px solid var(--green) !important;
    color: var(--green) !important;
    border-radius: 2px !important;
    padding: 0.65rem 1rem !important;
    transition: all 0.3s !important;
    box-shadow: inset 0 0 20px rgba(0,232,122,0.05) !important;
    animation: pulse-green 2.5s ease-in-out infinite !important;
}

.stDownloadButton > button:hover {
    background: rgba(0,232,122,0.18) !important;
    box-shadow: 0 0 30px rgba(0,232,122,0.5) !important;
    color: #fff !important;
    animation: none !important;
}

@keyframes pulse-green {
    0%, 100% { box-shadow: inset 0 0 20px rgba(0,232,122,0.05), 0 0 8px rgba(0,232,122,0.2); }
    50%       { box-shadow: inset 0 0 20px rgba(0,232,122,0.05), 0 0 20px rgba(0,232,122,0.45); }
}

/* ── File Uploader ────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    background: var(--card) !important;
    border: 1px dashed rgba(0,245,212,0.25) !important;
    border-radius: 3px !important;
    padding: 0.5rem !important;
    transition: border-color 0.3s !important;
}

[data-testid="stFileUploader"]:hover {
    border-color: var(--cyan) !important;
    box-shadow: 0 0 15px rgba(0,245,212,0.1) !important;
}

[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] p {
    color: var(--dim) !important;
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 0.82rem !important;
}

[data-testid="stFileUploader"] button {
    font-family: 'Orbitron', monospace !important;
    font-size: 0.65rem !important;
    border: 1px solid var(--cyan) !important;
    color: var(--cyan) !important;
    background: transparent !important;
}

/* ── Text Inputs ──────────────────────────────────────────── */
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    color: var(--cyan) !important;
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 0.85rem !important;
    border-radius: 2px !important;
    caret-color: var(--cyan);
}

.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: var(--cyan) !important;
    box-shadow: 0 0 0 1px var(--cyan), 0 0 15px rgba(0,245,212,0.2) !important;
    outline: none !important;
}

.stTextInput label, .stNumberInput label {
    font-family: 'Rajdhani', sans-serif !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.08em !important;
    color: var(--dim) !important;
}

/* ── Checkbox ──────────────────────────────────────────────── */
.stCheckbox label span {
    color: var(--text) !important;
    font-family: 'Rajdhani', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.05em !important;
}

/* ── Progress ──────────────────────────────────────────────── */
.stProgress {
    margin: 0.8rem 0 !important;
}

.stProgress > div {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 2px !important;
    height: 8px !important;
    overflow: hidden;
}

.stProgress > div > div {
    background: linear-gradient(90deg, var(--cyan2), var(--cyan), var(--magenta)) !important;
    background-size: 200% 100% !important;
    animation: progress-glow 1.5s ease-in-out infinite !important;
    box-shadow: 0 0 12px var(--glow-c) !important;
    border-radius: 2px !important;
}

@keyframes progress-glow {
    0%, 100% { box-shadow: 0 0 8px var(--glow-c); }
    50%       { box-shadow: 0 0 20px var(--glow-c), 0 0 35px rgba(0,245,212,0.15); }
}

/* ── Metrics ──────────────────────────────────────────────── */
[data-testid="stMetric"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-top: 2px solid var(--cyan) !important;
    padding: 0.8rem 1rem !important;
    border-radius: 2px !important;
}

[data-testid="stMetricLabel"] {
    font-family: 'Orbitron', monospace !important;
    font-size: 0.62rem !important;
    letter-spacing: 0.18em !important;
    color: var(--dim) !important;
}

[data-testid="stMetricValue"] {
    font-family: 'Orbitron', monospace !important;
    font-size: 1.6rem !important;
    color: var(--cyan) !important;
    text-shadow: 0 0 15px var(--glow-c) !important;
}

/* ── Alerts ───────────────────────────────────────────────── */
[data-testid="stAlert"] {
    background: var(--card) !important;
    border-radius: 2px !important;
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 0.82rem !important;
    border-left: 3px solid !important;
}

/* ── Custom Section Elements ─────────────────────────────── */
.section-label {
    font-family: 'Orbitron', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.22em;
    color: var(--cyan);
    text-transform: uppercase;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid var(--border);
    margin-bottom: 0.7rem;
    margin-top: 0.3rem;
}

.cyber-divider {
    border: none;
    border-top: 1px solid var(--border);
    margin: 1.4rem 0;
}

.empty-hint {
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.82rem;
    color: var(--dim);
    text-align: center;
    padding: 2.5rem 1rem;
    border: 1px dashed rgba(74,96,128,0.3);
    border-radius: 2px;
    margin-top: 1rem;
}

/* ── File List Rows ───────────────────────────────────────── */
.file-list-header {
    display: grid;
    grid-template-columns: 3rem 1fr;
    gap: 0.5rem;
    padding: 0.35rem 0.6rem;
    font-family: 'Orbitron', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.15em;
    color: var(--dim);
    border-bottom: 1px solid var(--border);
    margin-bottom: 0.3rem;
}

.file-row-item {
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.82rem;
    color: var(--text);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    padding: 0.3rem 0;
    line-height: 1.8;
}

.row-num {
    font-family: 'Orbitron', monospace;
    font-size: 0.72rem;
    color: var(--magenta);
    font-weight: 700;
    text-align: center;
    padding-top: 0.5rem;
    text-shadow: 0 0 10px var(--glow-m);
}

.ext-badge {
    display: inline-block;
    font-size: 0.65rem;
    padding: 0.1em 0.45em;
    border-radius: 2px;
    margin-right: 0.4em;
    font-family: 'Orbitron', monospace;
    letter-spacing: 0.08em;
    vertical-align: middle;
}

.ext-docx { background: rgba(0,184,255,0.15); color: var(--cyan2); border: 1px solid rgba(0,184,255,0.3); }
.ext-doc  { background: rgba(255,45,120,0.12); color: var(--magenta); border: 1px solid var(--border2); }

/* ── Log Box ──────────────────────────────────────────────── */
.log-box {
    background: #000;
    border: 1px solid var(--border);
    border-left: 3px solid var(--cyan);
    padding: 0.8rem 1rem;
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.78rem;
    max-height: 280px;
    overflow-y: auto;
    border-radius: 2px;
    line-height: 1.7;
}

.log-line { padding: 0.05rem 0; }
.log-ok  { color: var(--green); }
.log-err { color: var(--magenta); }
.log-inf { color: var(--cyan); }
.log-wrn { color: var(--yellow); }

/* ── Preview Box ──────────────────────────────────────────── */
.preview-box {
    background: var(--card);
    border: 1px solid var(--border);
    border-left: 3px solid var(--yellow);
    padding: 0.65rem 1rem;
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.82rem;
    border-radius: 2px;
    color: var(--text);
    margin-bottom: 0.5rem;
}

.hl-cyan   { color: var(--cyan);    font-weight: bold; }
.hl-mag    { color: var(--magenta); font-weight: bold; }
.hl-yellow { color: var(--yellow);  font-weight: bold; }
.hl-green  { color: var(--green);   font-weight: bold; }

/* ── Metric card (custom) ──────────────────────────────────── */
.cyber-metric {
    background: var(--card);
    border: 1px solid var(--border);
    border-top: 2px solid var(--cyan);
    padding: 1rem;
    text-align: center;
    border-radius: 2px;
    height: 100%;
}

.cyber-metric .value {
    font-family: 'Orbitron', monospace;
    font-size: 2rem;
    color: var(--cyan);
    font-weight: 700;
    text-shadow: 0 0 20px var(--glow-c);
    line-height: 1.1;
}

.cyber-metric .label {
    font-family: 'Orbitron', monospace;
    font-size: 0.58rem;
    letter-spacing: 0.2em;
    color: var(--dim);
    margin-top: 0.3rem;
}

/* ── Status line ──────────────────────────────────────────── */
.status-line {
    font-family: 'Share Tech Mono', monospace;
    font-size: 0.82rem;
    color: var(--dim);
    padding: 0.3rem 0;
}

/* ── scrollbar ────────────────────────────────────────────── */
::-webkit-scrollbar              { width: 6px; height: 6px; }
::-webkit-scrollbar-track        { background: var(--bg0); }
::-webkit-scrollbar-thumb        { background: var(--border); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover  { background: var(--cyan); }
</style>
"""

st.markdown(CYBER_CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# 排序邏輯 (沿用 V03 的 initial_sort_key)
# ──────────────────────────────────────────────────────────────────────────────
_SEQ_RE = re.compile(r"^(\d{1,4})")


def _parse_prefix_seq(stem: str) -> int:
    m = _SEQ_RE.match(stem)
    return int(m.group(1)) if m else 999_999


def _is_toc(stem: str) -> bool:
    return "目錄" in stem


def initial_sort_key(filename: str) -> tuple:
    stem = Path(filename).stem
    return (0 if _is_toc(stem) else 1, _parse_prefix_seq(stem), stem.lower())


# ──────────────────────────────────────────────────────────────────────────────
# 核心工具函數
# ──────────────────────────────────────────────────────────────────────────────
def parse_page_ranges(range_str: str, max_pages: int) -> list:
    """將 '1, 3-5, 8' 解析為 0-based page index list。"""
    indices = []
    for part in range_str.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                s, e = part.split("-", 1)
                start, end = max(1, int(s.strip())), min(max_pages, int(e.strip()))
                if start <= end:
                    indices.extend(range(start - 1, end))
            else:
                p = int(part)
                if 1 <= p <= max_pages:
                    indices.append(p - 1)
        except ValueError:
            pass
    return indices


def convert_word_to_pdf_via_libreoffice(
    file_bytes: bytes, filename: str, tmpdir: str
) -> tuple:
    """
    使用 LibreOffice CLI 將 .doc/.docx 轉換為 PDF。
    回傳 (pdf_bytes | None, log_message)。
    """
    src = Path(tmpdir) / filename
    src.write_bytes(file_bytes)
    try:
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                str(src),
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
        pdf_path = Path(tmpdir) / (src.stem + ".pdf")
        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            return pdf_path.read_bytes(), f"✅ {filename}"
        else:
            stderr_snippet = result.stderr[:300] if result.stderr else "(no stderr)"
            return None, f"❌ {filename} — LibreOffice 無產出。\n   stderr: {stderr_snippet}"
    except subprocess.TimeoutExpired:
        return None, f"⏱ {filename} — 轉換逾時 (120 秒)，請確認檔案未損毀。"
    except FileNotFoundError:
        return (
            None,
            "❌ 找不到 LibreOffice！\n"
            "   請在 packages.txt 加入 'libreoffice' 並重新部署。",
        )
    except Exception as exc:
        return None, f"❌ {filename} — 未預期錯誤：{exc}"


# ──────────────────────────────────────────────────────────────────────────────
# Session State 管理
# ──────────────────────────────────────────────────────────────────────────────
def _reset_tab1_results():
    st.session_state.pop("tab1_result", None)


def _sync_file_order(uploaded_files: list) -> list:
    """
    偵測上傳清單變化，若有新上傳則以 initial_sort_key 重新初始化順序。
    回傳目前有效的 file_order (list of str)。
    """
    names = [f.name for f in uploaded_files]
    current_set = frozenset(names)

    if st.session_state.get("upload_set") != current_set:
        sorted_names = sorted(names, key=initial_sort_key)
        st.session_state.file_order = sorted_names
        st.session_state.upload_set = current_set
        _reset_tab1_results()

    return st.session_state.get("file_order", [])


# ──────────────────────────────────────────────────────────────────────────────
# UI 元件：可排序的檔案清單
# ──────────────────────────────────────────────────────────────────────────────
def render_file_table(file_order: list):
    """顯示帶有上移 / 下移 / 刪除按鈕的檔案清單。"""
    st.markdown('<div class="file-list-header"><span>#</span><span>FILENAME</span></div>', unsafe_allow_html=True)

    for i, name in enumerate(file_order):
        ext = Path(name).suffix.upper().lstrip(".")
        col_num, col_name, col_up, col_dn, col_del = st.columns(
            [0.055, 0.68, 0.075, 0.075, 0.115]
        )

        with col_num:
            st.markdown(f'<div class="row-num">{i + 1:02d}</div>', unsafe_allow_html=True)

        with col_name:
            st.markdown(
                f'<div class="file-row-item">'
                f'<span class="ext-badge ext-{ext.lower()}">{ext}</span>{name}'
                f'</div>',
                unsafe_allow_html=True,
            )

        with col_up:
            if st.button("▲", key=f"up_{i}", disabled=(i == 0), help="上移"):
                lst = st.session_state.file_order
                lst[i], lst[i - 1] = lst[i - 1], lst[i]
                _reset_tab1_results()
                st.rerun()

        with col_dn:
            if st.button("▼", key=f"dn_{i}", disabled=(i == len(file_order) - 1), help="下移"):
                lst = st.session_state.file_order
                lst[i], lst[i + 1] = lst[i + 1], lst[i]
                _reset_tab1_results()
                st.rerun()

        with col_del:
            if st.button("✕ 移除", key=f"del_{i}", help="從清單移除（不影響原始上傳）"):
                st.session_state.file_order.pop(i)
                _reset_tab1_results()
                st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# 核心轉換邏輯
# ──────────────────────────────────────────────────────────────────────────────
def run_conversion(file_order: list, file_map: dict, do_merge: bool, merge_name: str):
    """執行 Word → PDF 批次轉換（並依需求合併）。結果存入 session_state。"""
    total = len(file_order)
    logs = []
    pdfs_in_order = []          # list of (name, bytes)

    progress_bar = st.progress(0, text="⚙  系統初始化中...")
    status_box   = st.empty()

    with tempfile.TemporaryDirectory() as tmpdir:
        # ── 轉換階段 ──────────────────────────────────────────
        for i, name in enumerate(file_order):
            frac = i / total
            progress_bar.progress(frac, text=f"⚙  轉換中 ({i + 1}/{total})：{name}")
            status_box.markdown(
                f'<div class="status-line">▷ 處理：<span class="hl-cyan">{name}</span></div>',
                unsafe_allow_html=True,
            )

            f = file_map[name]
            # 需 seek(0) 以防止重複讀取時位置錯誤
            f.seek(0)
            pdf_bytes, log_msg = convert_word_to_pdf_via_libreoffice(
                f.read(), f.name, tmpdir
            )

            if pdf_bytes:
                pdfs_in_order.append((name, pdf_bytes))
                logs.append(("ok", log_msg))
            else:
                logs.append(("err", log_msg))

        # ── 輸出階段 ──────────────────────────────────────────
        progress_bar.progress(0.95, text="⚙  組裝輸出檔案...")
        download_bytes    = None
        download_filename = ""
        download_mime     = ""

        if pdfs_in_order:
            if do_merge:
                status_box.markdown(
                    '<div class="status-line">🧩 合併 PDF 中，依清單順序...</div>',
                    unsafe_allow_html=True,
                )
                merger = PdfWriter()
                for _, pdf_b in pdfs_in_order:
                    merger.append(io.BytesIO(pdf_b))
                out = io.BytesIO()
                merger.write(out)
                download_bytes    = out.getvalue()
                download_filename = (
                    merge_name if merge_name.endswith(".pdf") else merge_name + ".pdf"
                )
                download_mime = "application/pdf"
                logs.append(("inf", f"🎉 合併完成（{len(pdfs_in_order)} 個）→ {download_filename}"))
            else:
                status_box.markdown(
                    '<div class="status-line">📦 打包為 ZIP...</div>',
                    unsafe_allow_html=True,
                )
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, pdf_b in pdfs_in_order:
                        zf.writestr(Path(name).stem + ".pdf", pdf_b)
                download_bytes    = zip_buf.getvalue()
                download_filename = "converted_pdfs.zip"
                download_mime     = "application/zip"
                logs.append(("inf", f"📦 已打包 {len(pdfs_in_order)} 個 PDF → {download_filename}"))
        else:
            logs.append(("err", "⚠  沒有任何成功轉換的檔案，無法產生輸出。"))

        progress_bar.progress(1.0, text="✔  完成")
        status_box.empty()

    ok_cnt   = sum(1 for kind, _ in logs if kind == "ok")
    fail_cnt = sum(1 for kind, _ in logs if kind == "err")

    st.session_state.tab1_result = {
        "total": total,
        "ok":    ok_cnt,
        "fail":  fail_cnt,
        "logs":  logs,
        "download_bytes":    download_bytes,
        "download_filename": download_filename,
        "download_mime":     download_mime,
    }
    st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# Tab 1：批次轉檔與合併
# ──────────────────────────────────────────────────────────────────────────────
def tab_convert():
    st.markdown('<div class="section-label">📤 上傳 Word 文件</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "支援 .doc / .docx，可一次選取多個檔案",
        type=["doc", "docx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="tab1_uploader",
    )

    if not uploaded:
        st.markdown(
            '<div class="empty-hint">'
            '上傳 .doc / .docx 文件後，系統將以<span class="hl-cyan">智慧排序</span>自動整理清單。<br>'
            '你可使用 ▲▼ 按鈕調整合併順序，接著按下執行按鈕。'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    file_order = _sync_file_order(uploaded)
    file_map   = {f.name: f for f in uploaded}

    # 過濾掉已被移除（不在 file_map 中）的項目
    file_order = [n for n in file_order if n in file_map]
    st.session_state.file_order = file_order

    if not file_order:
        st.warning("⚠  清單已清空，請重新上傳或重整頁面。")
        return

    # ── 檔案清單 ──────────────────────────────────────────────
    render_file_table(file_order)

    st.markdown('<hr class="cyber-divider">', unsafe_allow_html=True)

    # ── 選項 ──────────────────────────────────────────────────
    st.markdown('<div class="section-label">⚙ 轉換選項</div>', unsafe_allow_html=True)
    col_chk, col_name, _ = st.columns([0.22, 0.32, 0.46])

    with col_chk:
        do_merge = st.checkbox("合併為單一 PDF", value=True, key="do_merge")

    with col_name:
        merge_name = st.text_input(
            "合併後檔名",
            value="merged.pdf",
            disabled=not do_merge,
            label_visibility="collapsed",
            key="merge_name",
            placeholder="merged.pdf",
        )

    st.markdown('<hr class="cyber-divider">', unsafe_allow_html=True)

    # ── 執行按鈕 ──────────────────────────────────────────────
    if st.button(
        "▶▶  開始批次轉換  ◀◀",
        type="primary",
        use_container_width=True,
        key="exec_btn",
    ):
        run_conversion(file_order, file_map, do_merge, merge_name)

    # ── 結果顯示 ──────────────────────────────────────────────
    if "tab1_result" not in st.session_state:
        return

    res = st.session_state.tab1_result
    st.markdown('<hr class="cyber-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">📊 執行結果</div>', unsafe_allow_html=True)

    # Metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("總計", res["total"])
    m2.metric("✅ 成功", res["ok"])
    m3.metric("❌ 失敗", res["fail"])
    m4.metric("輸出", "1 PDF" if do_merge else f"{res['ok']} 個")

    # Log
    st.markdown('<div class="section-label" style="margin-top:1rem;">📟 作業日誌</div>', unsafe_allow_html=True)
    css_class_map = {"ok": "log-ok", "err": "log-err", "inf": "log-inf", "wrn": "log-wrn"}
    log_html = "".join(
        f'<div class="log-line {css_class_map.get(kind, "")}">{msg}</div>'
        for kind, msg in res["logs"]
    )
    st.markdown(f'<div class="log-box">{log_html}</div>', unsafe_allow_html=True)

    # Download
    if res.get("download_bytes"):
        st.markdown('<hr class="cyber-divider">', unsafe_allow_html=True)
        st.markdown('<div class="section-label">⬇ 下載區</div>', unsafe_allow_html=True)
        st.download_button(
            label=f"⬇  下載  {res['download_filename']}",
            data=res["download_bytes"],
            file_name=res["download_filename"],
            mime=res["download_mime"],
            use_container_width=True,
            key="download_btn",
        )


# ──────────────────────────────────────────────────────────────────────────────
# Tab 2：PDF 擷取 / 重組
# ──────────────────────────────────────────────────────────────────────────────
def tab_extract():
    st.markdown('<div class="section-label">📤 上傳來源 PDF</div>', unsafe_allow_html=True)

    pdf_file = st.file_uploader(
        "上傳單一 PDF 進行頁面擷取與重組",
        type=["pdf"],
        key="tab2_pdf",
        label_visibility="collapsed",
    )

    if not pdf_file:
        st.markdown(
            '<div class="empty-hint">'
            '上傳 PDF 後，輸入頁碼範圍進行<span class="hl-cyan">擷取</span>與<span class="hl-mag">重組</span>。<br>'
            '格式範例：<span class="hl-yellow">1, 3-5, 8, 10-12</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    # ── 讀取 PDF 基本資訊 ──────────────────────────────────────
    try:
        pdf_file.seek(0)
        reader      = PdfReader(pdf_file)
        total_pages = len(reader.pages)
    except Exception as exc:
        st.error(f"❌ 無法讀取 PDF：{exc}")
        return

    # ── 資訊顯示 + 輸入區 ────────────────────────────────────
    col_info, col_inputs = st.columns([0.25, 0.75])

    with col_info:
        st.markdown(
            f'<div class="cyber-metric">'
            f'<div class="value">{total_pages}</div>'
            f'<div class="label">TOTAL PAGES</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col_inputs:
        range_str = st.text_input(
            "頁碼範圍",
            placeholder="例：1, 3-5, 8, 10-12",
            help="以逗號分隔頁碼或範圍（用連字號），超出總頁數的頁碼將自動忽略。",
            key="range_str",
        )
        output_name = st.text_input(
            "輸出檔名",
            value="extracted.pdf",
            key="out_name",
        )

    # ── 即時預覽 ──────────────────────────────────────────────
    if range_str.strip():
        indices = parse_page_ranges(range_str, total_pages)
        if indices:
            pages_preview = ", ".join(str(i + 1) for i in indices)
            st.markdown(
                f'<div class="preview-box">'
                f'📄 將擷取 <span class="hl-cyan">{len(indices)}</span> 頁 '
                f'<span class="hl-yellow">→</span> {pages_preview}'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning("⚠  目前輸入無有效頁碼，請確認格式（例：1, 3-5, 8）。")

    st.markdown('<hr class="cyber-divider">', unsafe_allow_html=True)

    # ── 執行 ──────────────────────────────────────────────────
    if st.button("✂  執行擷取與重組", type="primary", use_container_width=True, key="extract_btn"):
        if not range_str.strip():
            st.error("請先輸入頁碼範圍。")
            st.stop()

        indices = parse_page_ranges(range_str, total_pages)
        if not indices:
            st.error("無有效頁碼，請確認格式（例：1, 3-5, 8）。")
            st.stop()

        prog = st.progress(0, text="擷取中...")
        writer = PdfWriter()

        pdf_file.seek(0)
        reader = PdfReader(pdf_file)  # 重讀（seek 後）

        for j, idx in enumerate(indices):
            writer.add_page(reader.pages[idx])
            prog.progress((j + 1) / len(indices), text=f"加入第 {idx + 1} 頁...")

        out = io.BytesIO()
        writer.write(out)
        prog.empty()

        filename = (
            output_name if output_name.strip().endswith(".pdf")
            else (output_name.strip() or "extracted") + ".pdf"
        )

        st.success(f"✅ 擷取完成！共 {len(indices)} 頁 → {filename}")
        st.markdown('<div class="section-label">⬇ 下載區</div>', unsafe_allow_html=True)
        st.download_button(
            label=f"⬇  下載  {filename}",
            data=out.getvalue(),
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
            key="extract_download_btn",
        )


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
def main():
    st.markdown(
        '<div class="cyber-header">⚡ PDF ENGINEERING SYSTEM</div>'
        '<div class="cyber-subtitle">'
        'WORD → PDF &nbsp;·&nbsp; BATCH MERGE &nbsp;·&nbsp; PAGE EXTRACT'
        ' &nbsp;—&nbsp; STREAMLIT EDITION'
        '</div>',
        unsafe_allow_html=True,
    )

    tab1, tab2 = st.tabs(["⚡  批次轉檔與合併", "✂  PDF 擷取 / 重組"])

    with tab1:
        tab_convert()

    with tab2:
        tab_extract()


if __name__ == "__main__":
    main()
