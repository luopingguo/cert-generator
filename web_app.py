"""证书批量生成器 — 网页版
本地运行: streamlit run web_app.py
线上部署: 上传到 Streamlit Cloud / Zeabur 等平台
"""

import io
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


def pptx_bytes_to_png(pptx_bytes):
    """将 PPTX 第一页转为 PNG（返回 bytes，失败返回 None）"""
    with tempfile.TemporaryDirectory() as tmpdir:
        base = Path(tmpdir)
        pptx_path = base / "input.pptx"
        pptx_path.write_bytes(pptx_bytes)

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

        prefix = base / "page"
        result = subprocess.run(
            ["pdftoppm", "-png", "-r", "200", "-f", "1", "-l", "1",
             str(pdf_files[0]), str(prefix)],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            return None

        generated = list(base.glob("page-*.png"))
        return generated[0].read_bytes() if generated else None


def generate_pngs(template_bytes, df):
    """批量生成 PNG 图片，返回 {文件名: bytes}"""
    results = {}
    has_en_column = "en_name" in df.columns
    total = len(df)

    for idx, (_, row) in enumerate(df.iterrows()):
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

        # 生成 PPTX → 转 PNG → PPTX 丢弃
        prs = Presentation(io.BytesIO(template_bytes))
        fill_template(prs, {
            "{{ID}}": uid,
            "{{CN_NAME}}": cn_name,
            "{{EN_NAME}}": en_name
        })
        pptx_buf = io.BytesIO()
        prs.save(pptx_buf)

        png_bytes = pptx_bytes_to_png(pptx_buf.getvalue())
        if png_bytes:
            results[f"{num}_{uid}_{cn_name}.png"] = png_bytes

    return results


# ====================== 页面 ======================

st.set_page_config(page_title="证书批量生成器", page_icon="📜", layout="centered")
st.title("📜 证书批量生成器")
st.caption("上传 CSV 数据文件 + PPTX 模板，批量导出证书 PNG 图片")

init_surname_dict()

# ---- 依赖检测 ----
if not check_png_deps():
    st.error("❌ 服务端未安装 LibreOffice，暂时无法使用。请联系管理员。")
    st.stop()

# ---- 使用说明 ----
with st.expander("📖 使用说明", expanded=True):
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
    if st.button("🚀 开始批量生成 PNG", type="primary", use_container_width=True):
        with st.spinner("正在生成证书图片..."):
            try:
                csv_file.seek(0)
                df = pd.read_csv(csv_file, dtype={"id": str, "number": str})
                pptx_bytes = pptx_file.read()

                results = generate_pngs(pptx_bytes, df)

                if not results:
                    st.error("生成失败，请检查模板和 CSV 数据是否正确。")
                elif len(results) == 1:
                    name, data = list(results.items())[0]
                    st.success("生成完成！")
                    st.image(data, caption=name, use_container_width=True)
                    st.download_button(
                        f"⬇ 下载 {name}", data=data, file_name=name,
                        mime="image/png"
                    )
                else:
                    st.success(f"生成完成！共 {len(results)} 张 PNG 图片")

                    # 打包 ZIP
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, data in results.items():
                            zf.writestr(name, data)
                    zip_buf.seek(0)

                    st.download_button(
                        label=f"⬇ 下载全部 PNG（{len(results)} 张，ZIP）",
                        data=zip_buf.getvalue(),
                        file_name="certificates.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                    # 默认展示第一张预览，其余可折叠查看
                    first_name, first_data = list(results.items())[0]
                    st.image(first_data, caption=first_name, use_container_width=True)

                    if len(results) > 1:
                        with st.expander(f"🖼 其余 {len(results) - 1} 张预览"):
                            for name, data in list(results.items())[1:]:
                                st.image(data, caption=name, use_container_width=True)

                    # 单张下载
                    with st.expander("📂 全部文件列表"):
                        for name in sorted(results.keys()):
                            st.download_button(
                                f"⬇ {name}", data=results[name],
                                file_name=name, key=name,
                                mime="image/png"
                            )

            except Exception as e:
                st.error(f"生成失败: {e}")

# ---- 页脚 ----
st.divider()
st.caption("💡 姓氏多音字自动纠正（卜→Bǔ、单→Shàn、仇→Qiú 等）| 基于 pypinyin")
