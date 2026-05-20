"""证书批量生成器 — 网页版
本地运行: streamlit run web_app.py
线上部署: 上传到 Streamlit Cloud / Zeabur 等平台
"""

import os
import io
import zipfile
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation
from pypinyin import lazy_pinyin, load_phrases_dict


# ====================== 工具函数 ======================

@st.cache_resource
def init_surname_dict():
    """加载姓氏多音字字典"""
    csv_path = Path(__file__).parent / "surname_pinyin.csv"
    if csv_path.exists():
        df = pd.read_csv(csv_path, dtype=str)
        phrases = {}
        for _, row in df.iterrows():
            char = row["character"].strip()
            pinyin_str = row["surname_pinyin"].strip()
            phrases[char] = [[s] for s in pinyin_str.split()]
        load_phrases_dict(phrases)


def is_chinese(text):
    if not text:
        return False
    return any('一' <= ch <= '鿿' for ch in str(text))


def chinese_to_pinyin(name):
    if not name or str(name).strip() == "":
        return ""
    pinyin_list = lazy_pinyin(str(name), strict=False)
    return " ".join(w.capitalize() for w in pinyin_list)


def clean_cell(val):
    s = str(val).strip()
    return s[:-2] if s.endswith(".0") else s


def fill_template(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        for key, value in replacements.items():
                            run.text = run.text.replace(key, value)
    return prs


def generate_pptxs(template_bytes, df):
    """批量生成 PPTX，返回 {文件名: bytes}"""
    results = {}
    has_en_column = "en_name" in df.columns

    for _, row in df.iterrows():
        uid = clean_cell(row["id"])
        cn_name = str(row["cn_name"]).strip()
        num = clean_cell(row["number"])

        en_name = ""
        if is_chinese(cn_name):
            if has_en_column:
                val = row["en_name"]
                if pd.notna(val) and str(val).strip() != "":
                    en_name = str(val).strip()
            if en_name == "":
                en_name = chinese_to_pinyin(cn_name)

        prs = Presentation(io.BytesIO(template_bytes))
        fill_template(prs, {
            "{{ID}}": uid,
            "{{CN_NAME}}": cn_name,
            "{{EN_NAME}}": en_name
        })

        fname = f"{num}_{uid}_{cn_name}.pptx"
        buf = io.BytesIO()
        prs.save(buf)
        results[fname] = buf.getvalue()

    return results


# ====================== 页面 ======================

st.set_page_config(page_title="证书批量生成器", page_icon="📜", layout="centered")
st.title("📜 证书批量生成器")
st.caption("上传 CSV 数据文件和 PPTX 模板，一键批量生成证书")

init_surname_dict()

# ---- 步骤说明 ----
with st.expander("📖 使用说明", expanded=False):
    st.markdown("""
    **CSV 文件格式**（必须有表头）：
    | number | id | cn_name | en_name |
    |--------|----|---------|---------|
    | 001 | 2026001 | 张三 | |
    | 002 | 2026002 | Lisa | |
    | 003 | 2026003 | 卜苗 | |

    **PPTX 模板占位符**：`{{ID}}` `{{CN_NAME}}` `{{EN_NAME}}`

    **EN_NAME 规则**：
    - cn_name 是英文 → 留空
    - en_name 列有值 → 直接用
    - 否则 → 拼音自动生成（含多音字纠正）
    """)

# ---- 文件上传 ----
col1, col2 = st.columns(2)
with col1:
    csv_file = st.file_uploader("📄 上传 CSV 数据文件", type=["csv"])
with col2:
    pptx_file = st.file_uploader("📄 上传 PPTX 模板文件", type=["pptx"])

# ---- 预览 ----
if csv_file is not None:
    try:
        preview_df = pd.read_csv(csv_file, dtype={"id": str, "number": str})
        st.subheader("📋 数据预览")
        st.dataframe(preview_df, use_container_width=True)
        st.caption(f"共 {len(preview_df)} 条记录")
    except Exception as e:
        st.error(f"CSV 解析失败: {e}")

# ---- 生成按钮 ----
if csv_file and pptx_file:
    if st.button("🚀 开始批量生成", type="primary", use_container_width=True):
        with st.spinner("正在生成中..."):
            try:
                csv_file.seek(0)
                df = pd.read_csv(csv_file, dtype={"id": str, "number": str})
                pptx_bytes = pptx_file.read()

                results = generate_pptxs(pptx_bytes, df)

                if len(results) == 1:
                    # 单个文件直接下载
                    name, data = list(results.items())[0]
                    st.success(f"生成完成！")
                    st.download_button(
                        "⬇ 下载 {name}", data=data, file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                else:
                    # 多个文件打包成 zip
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, data in results.items():
                            zf.writestr(name, data)
                    zip_buf.seek(0)

                    st.success(f"生成完成！共 {len(results)} 个 PPTX 文件")
                    st.download_button(
                        label=f"⬇ 下载全部 PPTX（{len(results)} 个，ZIP 压缩包）",
                        data=zip_buf.getvalue(),
                        file_name="certificates.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                    # 同时展示文件列表
                    with st.expander("📂 文件列表"):
                        for name in sorted(results.keys()):
                            st.write(f"- {name}")
                            st.download_button(
                                f"⬇ {name}", data=results[name],
                                file_name=name,
                                key=name,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )

            except Exception as e:
                st.error(f"生成失败: {e}")

# ---- 页脚 ----
st.divider()
st.caption("💡 姓氏多音字自动纠正（卜→Bǔ、单→Shàn、仇→Qiú 等）| 基于 pypinyin")
