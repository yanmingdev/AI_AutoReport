# =============================================================================
# report_generator.py  — Cloud-ready 版本（Streamlit Community Cloud）
# =============================================================================
# 變更重點：
# - 使用相對路徑 BASE_DIR；不再使用 D:\... 硬路徑
# - GEMINI_API_KEY 先讀 st.secrets，沒有再讀 .env
# - 下載檔名邏輯：優先用側欄「專案名稱」，再嘗試從 AI 內容擷取，最後退回時間戳
# - 文字標題擷取更健壯；DOCX/PPT 產生流程穩定
# =============================================================================

import os
import io
import re
import logging
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
import streamlit as st
import streamlit.components.v1 as components

from google import genai
from google.genai import types

# PPTX / DOCX 相關
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# =============================================================================
# 0. 路徑與環境
# =============================================================================
BASE_DIR = Path(__file__).parent.resolve()

# =============================================================================
# 1. Logging 設定（Cloud 檔案系統為暫存，可寫但不保證持久）
# =============================================================================
log_dir = BASE_DIR / "logs"
log_dir.mkdir(parents=True, exist_ok=True)
today = datetime.now().strftime("%m%d")
log_file = log_dir / f"log_{today}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
logger.info("=== Application start ===")

# =============================================================================
# 2. 讀取 API Key（優先 st.secrets，其次 .env）
# =============================================================================
load_dotenv(BASE_DIR / ".env")  # 本機開發用；Cloud 主要讀 secrets
API_KEY = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
if not API_KEY:
    st.error("❌ 找不到 GEMINI_API_KEY，請在 Streamlit Secrets（或本機 .env）設定")
    st.stop()

# =============================================================================
# 3. 頁面設定
# =============================================================================
st.set_page_config(page_title="Gemini 文件產生器", page_icon="✨", layout="wide")

# =============================================================================
# 4. 全域 CSS（依 doc_type 變色；doc_type 尚未選取時使用預設）
# =============================================================================
st.markdown(
    f"""
<style>
:root {{
  --primary-color: {("#FF8C00" if st.session_state.get("doc_type", "結案報告") == "結案報告" else "#1E90FF")};
  --primary-light: {("#FF8C0333" if st.session_state.get("doc_type", "結案報告") == "結案報告" else "#1E90FF33")};
}}
.stButton>button {{
  background-color: var(--primary-color)!important;
  color: #fff!important;
  border: none!important;
}}
[data-testid="stSidebar"] [data-baseweb="tag"] {{ background-color: var(--primary-color)!important; }}
[data-testid="stSidebar"] [data-baseweb="tag"] span,
[data-testid="stSidebar"] [data-baseweb="tag"] svg,
[data-testid="stSidebar"] [data-baseweb="tag-close-button"] {{ color: #fff!important; }}
.rc-slider-rail {{ background-color: var(--primary-light)!important; }}
.rc-slider-track, .rc-slider-handle {{ background-color: var(--primary-color)!important; border-color: var(--primary-color)!important; }}
.header {{ margin-top:-2.4rem!important; margin-bottom:0!important; padding:0!important; }}
.big-title {{ font-size:28px!important; font-weight:800!important; margin:0; }}
.subtitle {{ font-size:16px!important; color:#ddd!important; margin:0; }}
section[data-testid="stSidebar"] {{ width:200px!important; }}
.block-title {{ font-size:20px!important; margin-top:1rem!important; margin-bottom:0.3rem!important; }}
.stTextArea textarea {{ height:200px!important; }}
[data-testid="stMarkdownContainer"] h2 {{ color:inherit!important; }}
</style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# 5. Sidebar（報告類型、欄位選擇、溫度）
# =============================================================================
st.sidebar.markdown("<p>生成報告格式：</p>", unsafe_allow_html=True)
doc_type = st.sidebar.selectbox("", ["結案報告", "需求文件"], index=0, label_visibility="collapsed")
st.session_state["doc_type"] = doc_type

st.sidebar.markdown("<p>選擇要生成的內容區塊：</p>", unsafe_allow_html=True)
available_blocks = ["專案名稱", "專案目標", "專案效益", "開發流程", "作業時程", "專案分工"]
selected_blocks = st.sidebar.multiselect("區塊", available_blocks, default=[], label_visibility="collapsed")

st.sidebar.markdown("<p>創意溫度<br>(0.0＝保守 ↔ 1.0＝創意)</p>", unsafe_allow_html=True)
creativity_temp = st.sidebar.slider("", 0.0, 1.0, 0.5, 0.1)

# =============================================================================
# 6. 頁面標題
# =============================================================================
st.markdown(
    f"""
<div class="header">
  <div class="big-title">🚀 Gemini {doc_type} 產生器</div>
  <div class="subtitle">只要簡單輸入口語化內容，AI 幫你生成專業 {doc_type}！</div>
</div>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# 7. 載入模板（改為相對路徑）
# =============================================================================
def load_template(doc_type: str) -> str:
    template_file = "prompt_template.txt" if doc_type == "結案報告" else "requirement_template.txt"
    tpl_path = BASE_DIR / "templates" / template_file
    if not tpl_path.exists():
        st.error(f"❌ 找不到範本：{tpl_path}")
        st.stop()
    return tpl_path.read_text(encoding="utf-8")

# =============================================================================
# 8. 呼叫 Gemini 生成內容
# =============================================================================
def generate_content(project_title: str,
                     project_objective: str,
                     project_benefit: str,
                     development_process: str,
                     timeline_schedule: str,
                     project_assignment: str) -> str:
    prompt = load_template(doc_type).format(
        title=project_title,
        goal=project_objective,
        benefit=project_benefit,
        process=development_process,
        schedule=timeline_schedule,
        assignment=project_assignment,
    )
    client = genai.Client(api_key=API_KEY)
    cfg = types.GenerateContentConfig(temperature=creativity_temp)
    resp = client.models.generate_content(model="gemini-1.5-flash", contents=[prompt], config=cfg)
    return resp.text

# =============================================================================
# 9. 動態輸入區
# =============================================================================
field_labels = {
    "專案名稱": "🧩 專案名稱",
    "專案目標": "🎯 專案目標",
    "專案效益": "✨ 專案效益",
    "開發流程": "🛠️ 開發流程",
    "作業時程": "⏳ 作業時程",
    "專案分工": "👥 專案分工",
}
field_values = {}
for i in range(0, len(selected_blocks), 3):
    cols = st.columns(3)
    for j, block in enumerate(selected_blocks[i : i + 3]):
        with cols[j]:
            st.markdown(f"<div class='block-title'>{field_labels[block]}</div>", unsafe_allow_html=True)
            field_values[block] = st.text_area(f"請填寫 {block}：", height=200, label_visibility="collapsed")

project_title       = field_values.get("專案名稱", "")
project_objective   = field_values.get("專案目標", "")
project_benefit     = field_values.get("專案效益", "")
development_process = field_values.get("開發流程", "")
timeline_schedule   = field_values.get("作業時程", "")
project_assignment  = field_values.get("專案分工", "")

st.write("---")

# =============================================================================
# 10. 標題擷取與檔名決策（更健壯）
# =============================================================================
def extract_project_title_from_body(text: str) -> str | None:
    # 舊模板風格
    pats = [
        r"一、專案名稱[^\n\r]*\n\s*[-＊*]\s*(.+)",
        r"一、專案名稱[^\n\r]*\n\s*(.+)",
    ]
    for pat in pats:
        m = re.search(pat, text)
        if m:
            t = re.sub(r"\s*\(.*?\)$", "", m.group(1).strip())
            return t
    return None

def extract_colon_title(text: str) -> str | None:
    m = re.search(r"專案名稱[:：]\s*(.+)", text)
    return m.group(1).strip() if m else None

def extract_md_h1(text: str) -> str | None:
    m = re.search(r"^\s*#\s*(.+)$", text, re.M)  # 例如 "# XXX 專案"
    return m.group(1).strip() if m else None

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip("_ ").strip()

def decide_filename_base(user_title: str, generated_text: str, doc_type: str) -> str:
    # 1) 優先用側邊欄輸入
    if user_title and user_title.strip():
        return sanitize_filename(user_title)
    # 2) 從 AI 內容嘗試多種擷取法
    for fn in (extract_project_title_from_body, extract_colon_title, extract_md_h1):
        t = fn(generated_text)
        if t:
            return sanitize_filename(t)
    # 3) 退路：doc_type + 時間戳
    return f"{doc_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

# =============================================================================
# 11. session_state：保留 AI 內容
# =============================================================================
if "generated_text" not in st.session_state:
    st.session_state["generated_text"] = ""

# =============================================================================
# 12. 生成按鈕
# =============================================================================
if st.button(f"🪄 生成 {doc_type}", use_container_width=True):
    if not selected_blocks:
        st.warning("請至少選擇一個區塊")
    else:
        missing = [b for b in selected_blocks if not field_values.get(b, "").strip()]
        if missing:
            st.warning("⚠️ 尚未填寫：" + "、".join(missing))
        else:
            with st.spinner("AI 撰寫中..."):
                try:
                    out = generate_content(
                        project_title,
                        project_objective,
                        project_benefit,
                        development_process,
                        timeline_schedule,
                        project_assignment,
                    )
                    st.session_state["generated_text"] = out or ""
                except Exception as e:
                    st.error(f"❌ 發生錯誤：{e}")
                    st.session_state["generated_text"] = ""

# =============================================================================
# 13. 預覽 + 下載
# =============================================================================
if st.session_state.get("generated_text"):
    generated_text = st.session_state["generated_text"]

    st.success(f"🎉 {doc_type} 生成完成！")
    st.markdown(f"### 📌 {doc_type} 預覽")
    st.markdown(generated_text)
    st.code(generated_text, language="markdown")

    # Copy-to-clipboard 提示
    components.html(
        """
<script>
;(function(){
  const bind = ()=>{
    const btn = window.parent.document.querySelector('button[title="Copy to clipboard"]');
    if(!btn||btn.dataset.bound) return;
    btn.dataset.bound = '1';
    const lbl = document.createElement('span');
    lbl.id = 'copy-label';
    lbl.innerText = '點擊複製';
    lbl.style = 'margin-left:8px; color:var(--primary-color); font-size:16px; vertical-align:middle; top:8px; position:relative;';
    btn.parentElement.appendChild(lbl);
    btn.addEventListener('click', ()=>{ lbl.innerText = '已複製'; });
  };
  setInterval(bind,500);
})();
</script>
""",
        height=0,
    )

    # ---- 檔名決策（不再阻擋下載）----
    filename_base = decide_filename_base(project_title, generated_text, doc_type)

    # 預先解析 Markdown 標題（PPT & DOCX 都會用）
    headers = list(re.finditer(r"^(#+)\s*(.+)", generated_text, re.M))

    # -------------------------
    # 下載：PPTX
    # -------------------------
    try:
        ppt = Presentation()

        # 首頁大標題
        title_slide_layout = ppt.slide_layouts[0]
        slide = ppt.slides.add_slide(title_slide_layout)
        slide.shapes.title.text = filename_base
        if slide.placeholders and len(slide.placeholders) > 1:
            slide.placeholders[1].text = ""

        title_shape = slide.shapes.title
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.font.size = Pt(48)
        p.font.name = "微軟正黑體"
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

        def add_slide(title: str, content: str):
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            # 標題
            tf = slide.shapes.title.text_frame
            tf.clear()
            tf.margin_top = Pt(5)
            tf.vertical_anchor = MSO_ANCHOR.TOP
            p = tf.paragraphs[0]
            p.text = title
            p.font.name = "微軟正黑體"
            p.font.size = Pt(32)
            p.font.color.rgb = RGBColor(0, 108, 184)
            p.alignment = PP_ALIGN.LEFT
            # 內文
            body = slide.placeholders[1].text_frame
            body.clear()
            body.margin_top = Pt(5)
            body.vertical_anchor = MSO_ANCHOR.TOP
            for line in content.split("\n"):
                para = body.add_paragraph()
                para.text = line
                para.font.name = "微軟正黑體"
                para.font.size = Pt(24)
                para.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment = PP_ALIGN.LEFT
            try:
                body.fit_text(max_size=24)
            except Exception:
                pass

        if headers:
            for idx, h in enumerate(headers):
                start = h.end()
                end = headers[idx + 1].start() if idx + 1 < len(headers) else len(generated_text)
                add_slide(h.group(2).strip(), generated_text[start:end].strip())
        else:
            add_slide(filename_base, generated_text)

        buf = io.BytesIO()
        ppt.save(buf)
        buf.seek(0)
        st.download_button(
            label="📥 下載 PPT 檔",
            data=buf,
            file_name=f"{filename_base}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except ImportError:
        st.error("❌ 無法匯出 PPTX，請在 requirements.txt 加入 python-pptx")

    # -------------------------
    # 下載：DOCX
    # -------------------------
    try:
        from docx import Document
        from docx.shared import Pt as DocPt

        doc = Document()
        doc.styles["Normal"].font.name = "微軟正黑體"
        doc.styles["Normal"].font.size = DocPt(12)

        if headers:
            for idx, h in enumerate(headers):
                start = h.end()
                end = headers[idx + 1].start() if idx + 1 < len(headers) else len(generated_text)
                heading = h.group(2).strip()
                lines = generated_text[start:end].strip().split("\n")
                doc.add_heading(heading, level=2)
                for line in lines:
                    if line.strip():
                        p = doc.add_paragraph(line)
                        p.style = doc.styles["Normal"]
        else:
            doc.add_paragraph(generated_text)

        doc_buf = io.BytesIO()
        doc.save(doc_buf)
        doc_buf.seek(0)
        st.download_button(
            label="📥 下載 Word 檔",
            data=doc_buf,
            file_name=f"{filename_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    except ImportError:
        st.error("❌ 無法匯出 Word，請在 requirements.txt 加入 python-docx")
