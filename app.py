import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import dashscope
from dashscope import Generation
import os
import io
from docx import Document
from http import HTTPStatus

# ────────────────────────────────────────────────
# 1. 全局常量 & 缓存读取
# ────────────────────────────────────────────────
@st.cache_data(show_spinner="正在加载内置知识库...")
def load_builtin_knowledge():
    try:
        survey_df = pd.read_excel("调研要点.xlsx", sheet_name=0)
        survey_text = survey_df.to_string(index=False)

        tbs = pd.read_excel("TBS-V2.xlsx", sheet_name="Sheet2")
        tbs_grouped = tbs.groupby("模块")
        tbs_by_category = {}
        for cat in ["禁止类", "限制类", "关注类", "部分有", "是", "否"]:
            subset = tbs[tbs["禁止类 / 限制类 / 关注类"].str.contains(cat, na=False)]
            tbs_by_category[cat] = subset.to_dict(orient="records")

        bp_doc = Document("BP提示词内容.docx")
        bp_template = "\n".join(p.text.strip() for p in bp_doc.paragraphs if p.text.strip())

        return {
            "survey_text": survey_text,
            "tbs_raw": tbs,
            "tbs_grouped": tbs_grouped,
            "tbs_by_category": tbs_by_category,
            "tbs_text": tbs.to_string(index=False),
            "bp_template": bp_template
        }
    except Exception as e:
        st.error(f"加载内置知识库失败: {e}")
        st.stop()


BUILTIN_KNOWLEDGE = load_builtin_knowledge()

# ────────────────────────────────────────────────
# 2. 文件读取函数
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
# 3. 流式调用函数
# ────────────────────────────────────────────────
def call_llm_streaming(prompt: str):
    """使用流式输出调用 qwen-max，边生成边显示"""
    responses = Generation.call(
        model="qwen-max",
        prompt=prompt,
        temperature=0.1,
        max_tokens=4000,          # ← 从8000降至4000，大幅提速
        result_format="message",
        stream=True,              # ← 开启流式
        incremental_output=True   # ← 增量输出
    )
    return responses


# ────────────────────────────────────────────────
# 4. 主界面
# ────────────────────────────────────────────────
st.set_page_config(page_title="科技项目初筛分析", layout="wide")
st.title("🔍 科技项目初筛分析")

uploaded_file = st.file_uploader("上传项目资料（PDF / PPTX）", type=["pdf", "pptx"])

if uploaded_file:
    with st.spinner("正在提取文档内容..."):
        doc_text = extract_text_from_uploaded_file(uploaded_file)

    if not doc_text.strip():
        st.error("无法从文件中提取有效文本，请检查文件内容。")
    else:
        st.success("文档内容提取完成")

    if st.button("🚀 开始初筛分析（严格按BP提示词格式输出）", type="primary"):

        # ── 缩减 prompt，控制总 token 量 ──
        prompt = f"""{BUILTIN_KNOWLEDGE['bp_template']}

用户上传的BP文档内容（严格基于此内容分析，不得编造）：
{doc_text[:6000]}

TBS-V2 规则摘要（合规与风险评价）：
{BUILTIN_KNOWLEDGE['tbs_text'][:3000]}

调研要点摘要（评估维度补充）：
{BUILTIN_KNOWLEDGE['survey_text'][:1500]}
"""

        st.info(f"📊 Prompt 长度：约 {len(prompt)} 字符，正在生成报告...")

        # ── 流式输出到页面 ──
        result_placeholder = st.empty()
        full_result = ""

        try:
            with st.spinner("⏳ 正在生成初筛报告（流式输出，约30-50秒）..."):
                responses = call_llm_streaming(prompt)
                for response in responses:
                    if response.status_code == HTTPStatus.OK:
                        chunk = response.output.choices[0].message.content
                        if chunk:
                            full_result += chunk
                            result_placeholder.markdown(full_result + "▌")  # 光标效果
                    else:
                        st.error(f"API错误：{response.status_code} - {response.message}")
                        break

            # 输出完成，去掉光标
            result_placeholder.markdown(full_result)
            st.success("✅ 报告生成完成")

            # 下载按钮
            if full_result:
                st.download_button(
                    label="📥 下载完整初筛报告（Markdown）",
                    data=full_result,
                    file_name="项目初筛报告.md",
                    mime="text/markdown"
                )

        except Exception as e:
            st.error(f"调用大模型失败：{e}")
            st.info("💡 建议：检查 API Key 是否有效，或网络是否正常")
