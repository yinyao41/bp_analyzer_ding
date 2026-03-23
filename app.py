import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os
import io
from docx import Document

dashscope.api_key = os.environ.get("DASHSCOPE_API_KEY")

st.set_page_config(page_title="科技项目初筛分析", layout="wide")
st.title("🔍 科技项目初筛分析")

# ========================== 知识库加载 ==========================
@st.cache_data
def load_builtin_knowledge():
    try:
        bp_doc = Document("BP提示词内容.docx")
        bp_template = "\n".join(p.text.strip() for p in bp_doc.paragraphs if p.text.strip())
        
        tbs = pd.read_excel("TBS-V2.xlsx", sheet_name="Sheet2")
        survey_df = pd.read_excel("调研要点.xlsx", sheet_name=0)
        
        return {
            "bp_template": bp_template,
            "tbs_text": tbs.to_string(index=False),
            "survey_text": survey_df.to_string(index=False)
        }
    except Exception as e:
        st.error(f"知识库加载失败: {e}")
        st.stop()

BUILTIN_KNOWLEDGE = load_builtin_knowledge()

# ========================== 文件读取 ==========================
def extract_text_from_file(uploaded_file):
    if not uploaded_file:
        return ""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    bytes_data = uploaded_file.read()
    uploaded_file.seek(0)
    
    if ext == ".pdf":
        reader = PdfReader(io.BytesIO(bytes_data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    elif ext == ".pptx":
        prs = Presentation(io.BytesIO(bytes_data))
        return "\n".join(shape.text for slide in prs.slides 
                        for shape in slide.shapes if hasattr(shape, "text"))
    return ""

# ========================== 上传区 ==========================
uploaded_file = st.file_uploader(
    "上传项目资料（PDF / PPTX）",
    type=["pdf", "pptx"],
    help="建议文件不超过20MB"
)

if uploaded_file:
    with st.spinner("正在提取文档内容..."):
        doc_text = extract_text_from_file(uploaded_file)
    
    st.success("✅ 文档内容提取完成")
    
    if st.button("🚀 开始初筛分析", type="primary"):
        # === 关键优化：大幅缩短长度，避免超限 ===
        prompt = f"""{BUILTIN_KNOWLEDGE['bp_template']}

用户上传的BP文档内容（请严格基于此内容进行分析）：
{doc_text[:6000]}

TBS-V2 规则摘要：
{BUILTIN_KNOWLEDGE['tbs_text'][:3000]}

调研要点摘要：
{BUILTIN_KNOWLEDGE['survey_text'][:3000]}
"""

        with st.spinner("AI 正在生成报告..."):
            try:
                response = dashscope.Generation.call(
                    model="qwen-max",
                    messages=[{"role": "user", "content": prompt}],
                    result_format="message",
                    temperature=0.1,
                    max_tokens=12000
                )
                
                # 安全提取（兼容新旧返回结构）
                if (hasattr(response, "output") and 
                    hasattr(response.output, "choices") and 
                    response.output.choices):
                    result = response.output.choices[0].message.content
                else:
                    result = ""
                    
            except Exception as e:
                st.error(f"API 调用失败: {str(e)}")
                result = ""

        # ====================== 显示结果 ======================
        st.subheader("📋 项目初筛分析报告")
        
        if result and result.strip():
            st.markdown(result)
            st.download_button(
                label="📥 下载完整报告（Markdown）",
                data=result,
                file_name="项目初筛分析报告.md",
                mime="text/markdown"
            )
        else:
            st.warning("⚠️ AI 返回内容为空（可能是提示词过长或文档内容复杂）。建议尝试更短的 PDF 或稍后重试。")
        
        # ====================== 调试面板（上线后可删） ======================
        with st.expander("🔧 调试信息（点开看真实返回）", expanded=False):
            st.write("**Prompt 长度**:", len(prompt))
            st.write("**文档提取长度**:", len(doc_text))
            if 'response' in locals():
                st.write("**Raw Response 类型**:", type(response))
                st.write("**Output 是否存在**:", hasattr(response, "output"))
                if hasattr(response, "output"):
                    st.json(str(response.output)[:2000])  # 只显示前2000字符

else:
    st.info("请上传项目资料文件（PDF 或 PPTX）")
