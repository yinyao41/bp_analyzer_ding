import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os
import io
from docx import Document   # 新增：支持读取 BP提示词内容.docx

# ────────────────────────────────────────────────
# 1. 全局常量 & 缓存读取（程序启动时执行一次）
# ────────────────────────────────────────────────
@st.cache_data(show_spinner="正在加载内置知识库...")
def load_builtin_knowledge():
    try:
        # 原有 Excel 文件
        survey_df = pd.read_excel("调研要点.xlsx", sheet_name=0)
        survey_text = survey_df.to_string(index=False)

        tbs = pd.read_excel("TBS-V2.xlsx", sheet_name="Sheet2")
        tbs_grouped = tbs.groupby("模块")
        tbs_by_category = {}
        for cat in ["禁止类", "限制类", "关注类", "部分有", "是", "否"]:
            subset = tbs[tbs["禁止类 / 限制类 / 关注类"].str.contains(cat, na=False)]
            tbs_by_category[cat] = subset.to_dict(orient="records")

        # 新增：加载 BP提示词内容.docx（核心输出格式模板）
        bp_doc = Document("BP提示词内容.docx")
        bp_template = "\n".join(p.text.strip() for p in bp_doc.paragraphs if p.text.strip())

        return {
            "survey_text": survey_text,
            "tbs_raw": tbs,
            "tbs_grouped": tbs_grouped,
            "tbs_by_category": tbs_by_category,
            "tbs_text": tbs.to_string(index=False),
            "bp_template": bp_template   # ← 核心：BP提示词完整模板
        }
    except Exception as e:
        st.error(f"加载内置知识库失败: {e}")
        st.stop()


BUILTIN_KNOWLEDGE = load_builtin_knowledge()

# ────────────────────────────────────────────────
# 2. 文件读取函数（用户上传的商业计划书）
# ────────────────────────────────────────────────
def extract_text_from_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return ""
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    bytes_data = uploaded_file.read()
    uploaded_file.seek(0)
    if file_ext == ".pdf":
        reader = PdfReader(io.BytesIO(bytes_data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    elif file_ext == ".pptx":
        prs = Presentation(io.BytesIO(bytes_data))
        texts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texts.append(shape.text)
        return "\n".join(texts)
    else:
        return ""


# ────────────────────────────────────────────────
# 3. 主界面（标题已优化为 BP 初筛场景）
# ────────────────────────────────────────────────
st.set_page_config(page_title="科技项目初筛分析（BP提示词格式）", layout="wide")
st.title("🔍 科技项目初筛分析")
st.caption("严格按照 BP提示词内容.docx 格式输出 · 结合 TBS-V2 + 调研要点 · 仅基于上传文档")

uploaded_file = st.file_uploader("上传项目资料（PDF / PPTX）", type=["pdf", "pptx"])

if uploaded_file:
    with st.spinner("正在提取文档内容..."):
        doc_text = extract_text_from_uploaded_file(uploaded_file)

    if not doc_text.strip():
        st.error("无法从文件中提取有效文本，请检查文件内容。")
    else:
        st.success("文档内容提取完成")
        with st.expander("文档提取预览（前800字）"):
            st.text(doc_text[:800] + "...")

    if st.button("🚀 开始初筛分析（严格按BP提示词格式输出）", type="primary"):
        with st.spinner("正在生成结构化初筛报告..."):
            # 核心修改：使用 BP提示词内容.docx 作为完整输出模板
            prompt = f"""{BUILTIN_KNOWLEDGE['bp_template']}

用户上传的BP文档内容（请严格基于此内容进行分析，不得编造任何未提及信息）：
{doc_text[:12000]}

内置 TBS-V2 规则摘要（用于合规与风险评价）：
{BUILTIN_KNOWLEDGE['tbs_text'][:8000]}

内置调研要点摘要（用于评估维度补充）：
{BUILTIN_KNOWLEDGE['survey_text'][:3000]}
"""

            try:
                response = dashscope.Generation.call(
                    model="qwen-max",
                    prompt=prompt,
                    temperature=0.1,
