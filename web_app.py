"""证书批量生成器 — 网页版
本地运行: streamlit run web_app.py
线上部署: 上传到 Streamlit Cloud / Zeabur 等平台
"""

import io
import os
import shutil
import subprocess
import tempfile
import zipfile
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


@st.cache_resource
def check_png_deps():
    """检测 PNG 转换依赖是否可用"""
    return shutil.which("soffice") is not None and shutil.which("pdftoppm") is not None


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


def pptx_to_png_first_page(pptx_bytes):
    """将 PPTX 第一页转为 PNG 图片（返回 bytes）"""
    with tempfile.TemporaryDirectory() as tmpdir:
        base = Path(tmpdir)
        pptx_path = base / "input.pptx"
        pptx_path.write_bytes(pptx_bytes)

        # 1. PPTX → PDF
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(base), str(pptx_path)],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            return None

        pdf_files = list(base.glob("*.pdf"))
        if not pdf_files:
            return None

        # 2. PDF 第一页 → PNG
        prefix = base / "page"
        result = subprocess.run(
            ["pdftoppm", "-png", "-r", "200", "-f", "1", "-l", "1",
             str(pdf_files[0]), str(prefix)],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            return None

        generated = list(base.glob("page-*.png"))
        if not generated:
            return None

        return generated[0].read_bytes()


def generate_all(template_bytes, df, make_png):
    """批量生成 PPTX + 可选 PNG，返回 {文件名: bytes}"""
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

        base = f"{num}_{uid}_{cn_name}"
        pptx_buf = io.BytesIO()
        prs.save(pptx_buf)
        pptx_bytes = pptx_buf.getvalue()
        results[f"pptx/{base}.pptx"] = pptx_bytes

        if make_png:
            png_bytes = pptx_to_png_first_page(pptx_bytes)
            if png_bytes:
                results[f"png/{base}.png"] = png_bytes

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

# ---- PNG 选项 ----
png_available = check_png_deps()
make_png = st.checkbox("同时生成 PNG 图片（每份证书的第一页）",
                        value=png_available,
                        disabled=not png_available)
if not png_available:
    st.caption("⚠️ 云端未安装 LibreOffice，暂不支持 PNG（本地部署可启用）")

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

                results = generate_all(pptx_bytes, df, make_png)

                pptx_count = sum(1 for k in results if k.endswith(".pptx"))
                png_count = sum(1 for k in results if k.endswith(".png"))

                if len(results) == 1:
                    name, data = list(results.items())[0]
                    st.success("生成完成！")
                    st.download_button(
                        f"⬇ 下载 {name}", data=data, file_name=name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, data in results.items():
                            zf.writestr(name, data)
                    zip_buf.seek(0)

                    summary = f"PPTX {pptx_count} 个"
                    if png_count > 0:
                        summary += f"，PNG {png_count} 个"
                    st.success(f"生成完成！共 {summary}")

                    zip_name = "certificates.zip"
                    st.download_button(
                        label=f"⬇ 下载全部（{summary}，ZIP 压缩包）",
                        data=zip_buf.getvalue(),
                        file_name=zip_name,
                        mime="application/zip",
                        use_container_width=True
                    )

                    with st.expander("📂 文件列表"):
                        for name in sorted(results.keys()):
                            icon = "🖼" if name.endswith(".png") else "📄"
                            st.write(f"{icon} {name}")
                            ext = name.split(".")[-1]
                            mime_map = {
                                "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                "png": "image/png"
                            }
                            st.download_button(
                                f"⬇ {name}", data=results[name],
                                file_name=name, key=name,
                                mime=mime_map.get(ext, "application/octet-stream")
                            )

            except Exception as e:
                st.error(f"生成失败: {e}")

# ---- 页脚 ----
st.divider()
st.caption("💡 姓氏多音字自动纠正（卜→Bǔ、单→Shàn、仇→Qiú 等）| 基于 pypinyin")
