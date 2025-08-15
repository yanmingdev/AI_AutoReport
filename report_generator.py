# =============================================================================
# report_generator.py
# -----------------------------------------------------------------------------
# Streamlit「需求/結案報告 AI 產生器」— Cloud-ready（*.streamlit.app）
# - 相對路徑載入 templates/
# - 金鑰：st.secrets["GEMINI_API_KEY"] 優先，否則退回 .env
# - 下載檔名：優先用側欄「專案名稱」；其次由 AI 內容解析；最後用時間戳
# - Sidebar 寬度可調（預設 360px）
# - PPT 依章節自動分頁：支援 **粗體章節**、數字/中文數字編號、Markdown # 標題
# =============================================================================

from __future__ import annotations

import io
import os
import re
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple

import streamlit as st
from dotenv import load_dotenv

# Google Generative AI（新版 SDK）
from google import genai
from google.genai import types

# 檔案輸出
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# -----------------------------------------------------------------------------
# 0) 基本設定
# -----------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent.resolve()   # 專案根目錄
SIDEBAR_WIDTH_PX = 360                       # 側邊欄寬度（可改 320~420）

# Logging（Cloud 檔案系統為暫存）
log_dir = BASE_DIR / "logs"
log_dir.mkdir(exist_ok=True, parents=True)
log_path = log_dir / f"log_{datetime.now():%m%d}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)
logger.info("=== App start ===")

# -----------------------------------------------------------------------------
# 1) 讀取金鑰：st.secrets 優先，其次 .env
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Gemini 文件產生器", page_icon="✨", layout="wide")

load_dotenv(BASE_DIR / ".env")  # 本機開發可用
API_KEY = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
if not API_KEY:
    st.error("❌ 找不到 GEMINI_API_KEY，請在 Streamlit Secrets（或本機 .env）設定")
    st.stop()

# -----------------------------------------------------------------------------
# 2) 基礎樣式（加寬 Sidebar + 主色）
# -----------------------------------------------------------------------------
st.markdown(
    f"""
<style>
/* === 側邊欄加寬 === */
[data-testid="stSidebar"] {{
  width: {SIDEBAR_WIDTH_PX}px !important;
  min-width: {SIDEBAR_WIDTH_PX}px !important;
  max-width: {SIDEBAR_WIDTH_PX}px !important;
}}
/* multiselect 已選標籤可換行 */
[data-testid="stSidebar"] [data-baseweb="select"] div[role="combobox"] {{
  flex-wrap: wrap;
}}
/* 常用字級/行距調整 */
.block-title {{ font-size: 20px; margin: 8px 0 6px; }}
.stTextArea textarea {{ height: 200px !important; }}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 3) UI：Sidebar（報告類型、欄位、溫度）
# -----------------------------------------------------------------------------
st.sidebar.markdown("**生成報告格式：**")
doc_type = st.sidebar.selectbox("", ["結案報告", "需求文件"], index=0, label_visibility="collapsed")

st.sidebar.markdown("**選擇要生成的內容區塊：**")
BLOCKS = ["專案名稱", "專案目標", "專案效益", "開發流程", "作業時程", "專案分工"]
selected_blocks = st.sidebar.multiselect("區塊", BLOCKS, default=[], label_visibility="collapsed")

st.sidebar.markdown("**創意溫度**（0.0＝保守 ↔ 1.0＝創意）")
temperature = st.sidebar.slider("", 0.0, 1.0, 0.50, 0.05)

# 依類型決定主色（再注入樣式）
PRIMARY = "#FF8C00" if doc_type == "結案報告" else "#1E90FF"
PRIMARY_LIGHT = "#FF8C0333" if doc_type == "結案報告" else "#1E90FF33"
st.markdown(
    f"""
<style>
:root {{
  --primary: {PRIMARY};
  --primary-light: {PRIMARY_LIGHT};
}}
.stButton>button {{ background-color: var(--primary) !important; color: #fff !important; border: none !important; }}
.rc-slider-rail {{ background: var(--primary-light) !important; }}
.rc-slider-track, .rc-slider-handle {{ background: var(--primary) !important; border-color: var(--primary) !important; }}
[data-testid="stSidebar"] [data-baseweb="tag"] {{ background: var(--primary) !important; }}
[data-testid="stSidebar"] [data-baseweb="tag"] span,
[data-testid="stSidebar"] [data-baseweb="tag"] svg {{ color: #fff !important; }}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# 4) Header
# -----------------------------------------------------------------------------
st.markdown(
    f"""
<div style="margin-top:-2rem">
  <h1 style="margin:0">🚀 Gemini {doc_type} 產生器</h1>
  <p style="color:#bbb;margin:.25rem 0 0 0">輸入口語化內容，AI 產出專業 {doc_type}（可下載 Word / PPT）</p>
</div>
""",
    unsafe_allow_html=True,
)
st.write("---")

# -----------------------------------------------------------------------------
# 5) 載入模板（相對路徑）
# -----------------------------------------------------------------------------
def load_template(doc_type: str) -> str:
    """讀取 templates/ 下的模板文字。"""
    name = "prompt_template.txt" if doc_type == "結案報告" else "requirement_template.txt"
    path = BASE_DIR / "templates" / name
    if not path.exists():
        st.error(f"❌ 找不到範本：{path}")
        st.stop()
    return path.read_text(encoding="utf-8")

# -----------------------------------------------------------------------------
# 6) 呼叫 Gemini 產生內容
# -----------------------------------------------------------------------------
def generate_content(*, title: str, goal: str, benefit: str,
                     process: str, schedule: str, assignment: str) -> str:
    """用模板組合 Prompt，呼叫 Gemini 產生文字內容。"""
    prompt = load_template(doc_type).format(
        title=title,
        goal=goal,
        benefit=benefit,
        process=process,
        schedule=schedule,
        assignment=assignment,
    )
    client = genai.Client(api_key=API_KEY)
    cfg = types.GenerateContentConfig(temperature=temperature)
    resp = client.models.generate_content(model="gemini-1.5-flash", contents=[prompt], config=cfg)
    return resp.text or ""

# -----------------------------------------------------------------------------
# 7) 動態輸入欄
# -----------------------------------------------------------------------------
LABELS = {
    "專案名稱": "🧩 專案名稱",
    "專案目標": "🎯 專案目標",
    "專案效益": "✨ 專案效益",
    "開發流程": "🛠️ 開發流程",
    "作業時程": "⏳ 作業時程",
    "專案分工": "👥 專案分工",
}
values: dict[str, str] = {}

for i in range(0, len(selected_blocks), 3):
    cols = st.columns(3)
    for j, block in enumerate(selected_blocks[i : i + 3]):
        with cols[j]:
            st.markdown(f"<div class='block-title'>{LABELS[block]}</div>", unsafe_allow_html=True)
            values[block] = st.text_area(f"請填寫 {block}：", key=f"ta_{block}", label_visibility="collapsed")

project_title       = values.get("專案名稱", "")
project_goal        = values.get("專案目標", "")
project_benefit     = values.get("專案效益", "")
dev_process         = values.get("開發流程", "")
timeline_schedule   = values.get("作業時程", "")
project_assignment  = values.get("專案分工", "")

st.write("")

# -----------------------------------------------------------------------------
# 8) 標題解析 + 檔名決策
# -----------------------------------------------------------------------------
def sanitize_filename(name: str) -> str:
    """Windows/Unix 不允許的字元改為底線；去頭尾空白與底線。"""
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    return name.strip("_ ").strip()

def extract_title_from_numbered(text: str) -> Optional[str]:
    """舊模板：『一、專案名稱』後一行的內容。"""
    pats = [
        r"一、專案名稱[^\n\r]*\n\s*[-＊*]\s*(.+)",
        r"一、專案名稱[^\n\r]*\n\s*(.+)",
    ]
    for pat in pats:
        m = re.search(pat, text)
        if m:
            return m.group(1).strip()
    return None

def extract_title_from_colon(text: str) -> Optional[str]:
    """『專案名稱：XXX』格式。"""
    m = re.search(r"專案名稱[:：]\s*(.+)", text)
    return m.group(1).strip() if m else None

def extract_title_from_md(text: str) -> Optional[str]:
    """Markdown H1：『# XXX 專案』。"""
    m = re.search(r"^\s*#\s*(.+)$", text, re.M)
    return m.group(1).strip() if m else None

def decide_filename_base(user_title: str, generated: str, doc_type: str) -> str:
    """
    1) 優先用側欄『專案名稱』
    2) 再從 AI 內容解析（多種格式）
    3) 最後用 doc_type + 時間戳
    """
    if user_title.strip():
        return sanitize_filename(user_title)
    for fn in (extract_title_from_numbered, extract_title_from_colon, extract_title_from_md):
        val = fn(generated)
        if val:
            return sanitize_filename(val)
    return f"{doc_type}_{datetime.now():%Y%m%d_%H%M%S}"

# -----------------------------------------------------------------------------
# 9) 章節偵測：依標題切段（用於 PPT/Word 分頁）
# -----------------------------------------------------------------------------
HEADING_LINE = re.compile(
    r"""^(?:
        \s*#{1,6}\s*(?P<h_md>.+?)\s*$                                  # Markdown 標題
        |
        \s*\*\*\s*(?P<h_bold>.+?)\s*\*\*\s*$                           # 整行粗體 **章節**
        |
        \s*(?P<h_num>\d+|[一二三四五六七八九十]+)[\.\、]\s*(?P<h_txt>.+?)\s*$  # 1. / 一、 章節
    )""",
    re.X
)

def split_sections(text: str) -> List[Tuple[str, str]]:
    """
    依『章節標題行』將全文切成多段。
    支援：
      - Markdown 標題：# 標題
      - 整行粗體：**1. 標題**、**需求目標**
      - 數字/中文數字編號：1. 標題 / 一、標題
    回傳：[(title, body), ...]；若偵測不到章節，回傳空清單。
    """
    lines = text.splitlines()
    sections: List[Tuple[str, List[str]]] = []
    cur_title: Optional[str] = None
    cur_buf: List[str] = []

    def flush():
        if cur_title is not None or cur_buf:
            title = (cur_title or "").strip()
            body = "\n".join(cur_buf).strip()
            sections.append((title, body))

    for line in lines:
        m = HEADING_LINE.match(line)
        if m:
            # 碰到新章節 → 先收前一段
            flush()
            # 取出標題文字
            title = m.group("h_md") or m.group("h_bold") or m.group("h_txt") or ""
            cur_title = title.strip()
            cur_buf = []
        else:
            cur_buf.append(line)

    flush()  # 收尾

    # 移除完全空的章節
    result = [(t, b) for (t, b) in sections if (t or b)]
    # 若只有一段而且沒有標題，視為「未偵測到章節」
    if len(result) == 1 and not result[0][0]:
        return []
    return result

# -----------------------------------------------------------------------------
# 10) 生成按鈕
# -----------------------------------------------------------------------------
if "generated_text" not in st.session_state:
    st.session_state["generated_text"] = ""

if st.button(f"🪄 生成 {doc_type}", use_container_width=True):
    if not selected_blocks:
        st.warning("請至少選擇一個內容區塊")
    else:
        missing = [b for b in selected_blocks if not values.get(b, "").strip()]
        if missing:
            st.warning("⚠️ 尚未填寫：「" + "、".join(missing) + "」")
        else:
            with st.spinner("AI 撰寫中，請稍候…"):
                try:
                    text = generate_content(
                        title=project_title,
                        goal=project_goal,
                        benefit=project_benefit,
                        process=dev_process,
                        schedule=timeline_schedule,
                        assignment=project_assignment,
                    )
                    st.session_state["generated_text"] = text
                except Exception as e:
                    logger.exception("Generate error")
                    st.error(f"❌ 產生內容失敗：{e}")
                    st.session_state["generated_text"] = ""

# -----------------------------------------------------------------------------
# 11) 預覽 + 下載（Word / PPT）
# -----------------------------------------------------------------------------
output = st.session_state.get("generated_text", "")
if output:
    st.success(f"🎉 已完成 {doc_type} 內容！")
    st.markdown("### 📌 預覽")
    st.markdown(output)

    # 章節切割（重點）
    sections = split_sections(output)  # → [(title, body), ...]
    filename_base = decide_filename_base(project_title, output, doc_type)

    # ------------------ 下載：PPTX（依章節分頁） ------------------
    try:
        prs = Presentation()

        # 首頁
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = filename_base
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = ""

        # 首頁標題樣式
        title_tf = slide.shapes.title.text_frame
        p = title_tf.paragraphs[0]
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.name = "微軟正黑體"
        p.alignment = PP_ALIGN.CENTER

        def add_content_slide(title: str, body: str) -> None:
            s = prs.slides.add_slide(prs.slide_layouts[1])

            # 標題
            tf = s.shapes.title.text_frame
            tf.clear()
            tf.margin_top = Pt(5)
            tf.vertical_anchor = MSO_ANCHOR.TOP

            h = tf.paragraphs[0]
            h.text = title
            h.font.name = "微軟正黑體"
            h.font.size = Pt(32)
            h.font.color.rgb = RGBColor(0, 108, 184)
            h.alignment = PP_ALIGN.LEFT

            # 內文（逐行）
            body_tf = s.placeholders[1].text_frame
            body_tf.clear()
            body_tf.margin_top = Pt(5)
            body_tf.vertical_anchor = MSO_ANCHOR.TOP

            for line in (body or "").split("\n"):
                para = body_tf.add_paragraph()
                para.text = line
                para.font.name = "微軟正黑體"
                para.font.size = Pt(24)
                para.alignment = PP_ALIGN.LEFT

        if sections:
            for title, body in sections:
                add_content_slide(title if title else filename_base, body)
        else:
            # 偵測不到章節時，整段放一頁
            add_content_slide(filename_base, output)

        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)

        st.download_button(
            "📥 下載 PPT 檔",
            data=ppt_buf,
            file_name=f"{filename_base}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except Exception as e:
        logger.exception("PPT export error")
        st.error(f"❌ 匯出 PPTX 失敗：{e}（請確認 requirements.txt 已含 python-pptx）")

    # ------------------ 下載：DOCX（同樣依章節寫入） ------------------
    try:
        from docx import Document
        from docx.shared import Pt as DocPt

        doc = Document()
        doc.styles["Normal"].font.name = "微軟正黑體"
        doc.styles["Normal"].font.size = DocPt(12)

        if sections:
            for title, body in sections:
                if title:
                    doc.add_heading(title, level=2)
                for ln in (body or "").split("\n"):
                    if ln.strip():
                        p = doc.add_paragraph(ln)
                        p.style = doc.styles["Normal"]
        else:
            doc.add_paragraph(output)

        doc_buf = io.BytesIO()
        doc.save(doc_buf)
        doc_buf.seek(0)

        st.download_button(
            "📥 下載 Word 檔",
            data=doc_buf,
            file_name=f"{filename_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    except Exception as e:
        logger.exception("DOCX export error")
        st.error(f"❌ 匯出 Word 失敗：{e}（請確認 requirements.txt 已含 python-docx）")
