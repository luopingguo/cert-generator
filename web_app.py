"""证书批量生成器 — 网页版
本地运行: streamlit run web_app.py
线上部署: 上传到 Streamlit Cloud / Zeabur 等平台
"""

import io
import os
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation
from pypinyin import lazy_pinyin, load_phrases_dict


# ====================== 初始化 ======================

def _init_surname_dict():
    """加载姓氏多音字字典"""
    csv_path = Path(__file__).parent / "surname_pinyin.csv"
    if not csv_path.exists():
        return
    try:
        df = pd.read_csv(str(csv_path), dtype=str)
        phrases = {}
        for _, row in df.iterrows():
            char = str(row.get("character", "")).strip()
            pinyin_str = str(row.get("surname_pinyin", "")).strip()
            if char and pinyin_str:
                phrases[char] = [[s] for s in pinyin_str.split()]
        if phrases:
            load_phrases_dict(phrases)
    except Exception:
        pass  # 多音字字典加载失败不影响核心功能


def _check_deps():
    """检查 LibreOffice 和 poppler 是否可用"""
    return os.path.exists("/usr/bin/soffice") and os.path.exists("/usr/bin/pdftoppm")


_init_surname_dict()
DEPS_OK = _check_deps()


# ====================== 工具函数 ======================

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
    with tempfile.TemporaryDirectory() as tmpdir:
        base = Path(tmpdir)
        pptx_path = base / "input.pptx"
        pptx_path.write_bytes(pptx_bytes)

        env = os.environ.copy()
        env.setdefault("HOME", "/tmp")
        env.setdefault("SAL_USE_VCLPLUGIN", "gen")

        result = subprocess.run(
            ["soffice", "--headless", "--norestore", "--convert-to", "pdf",
             "--outdir", str(base), str(pptx_path)],
            capture_output=True, text=True, timeout=120,
            env=env,
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


# ====================== 页面 ======================

st.set_page_config(page_title="证书批量生成器", page_icon="📜", layout="wide")

BATCH_SIZE = 5  # 每批处理数量，避免容器内存超限

# 初始化 session state
for key, default in [("trigger", False), ("pptx_bytes", None),
                     ("work_dir", None), ("batch_idx", 0), ("total_rows", []),
                     ("png_names", []), ("first_png_name", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

if not DEPS_OK:
    st.error("❌ 服务端未安装 LibreOffice / poppler，无法使用。")
    st.stop()

# ===== 标题 =====
st.title("📜 证书批量生成器")
st.caption("上传 CSV + PPTX 模板，批量导出证书 PNG 图片")
st.divider()

# ===== 左右分栏 =====
left, right = st.columns([1, 1])

# ===== 左栏 =====
with left:
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

    ul1, ul2 = st.columns(2)
    with ul1:
        csv_file = st.file_uploader("📄 上传 CSV", type=["csv"], key="csv_up")
    with ul2:
        pptx_file = st.file_uploader("📄 上传 PPTX", type=["pptx"], key="pptx_up")

    if csv_file is not None:
        try:
            csv_bytes = csv_file.getvalue()
            preview_df = pd.read_csv(io.BytesIO(csv_bytes), dtype={"id": str, "number": str})
            st.caption(f"📋 数据预览（共 {len(preview_df)} 条，显示前 3 条）")
            st.dataframe(preview_df.head(3), use_container_width=True)
        except Exception as e:
            st.error(f"CSV 解析失败: {e}")

    if csv_file and pptx_file:
        if st.button("🚀 开始批量生成 PNG", type="primary", use_container_width=True):
            try:
                # 读取原始字节
                csv_bytes = csv_file.getvalue()
                pptx_bytes = pptx_file.getvalue()
                st.session_state.pptx_bytes = pptx_bytes

                # 清理旧临时目录
                if st.session_state.work_dir:
                    shutil.rmtree(st.session_state.work_dir, ignore_errors=True)

                # 准备批次数据
                df = pd.read_csv(io.BytesIO(csv_bytes), dtype={"id": str, "number": str})
                st.session_state.total_rows = df.to_dict("records")
                st.session_state.batch_idx = 0
                st.session_state.png_names = []
                st.session_state.first_png_name = None
                st.session_state.work_dir = str(Path(tempfile.mkdtemp(prefix="cert_")))
                st.session_state.trigger = True
                st.rerun()
            except Exception as e:
                st.error(f"启动生成失败: {e}")

# ===== 右栏 =====
with right:
    if st.session_state.trigger:
        try:
            total_rows = st.session_state.total_rows
            total = len(total_rows)
            batch_idx = st.session_state.batch_idx
            pptx_bytes = st.session_state.pptx_bytes
            work_dir = Path(st.session_state.work_dir)

            has_en_column = False
            if total_rows:
                row0 = total_rows[0]
                has_en_column = "en_name" in row0

            batch_end = min(batch_idx + BATCH_SIZE, total)
            is_last_batch = (batch_end >= total)
            batch_rows = total_rows[batch_idx:batch_end]

            progress_bar = st.progress(0, text=f"正在处理第 {batch_idx + 1}-{batch_end} 张...")
            status_text = st.empty()
            preview_spot = st.empty()

            for j, row in enumerate(batch_rows):
                uid = clean_cell(row.get("id", ""))
                cn_name = str(row.get("cn_name", "")).strip()
                num = clean_cell(row.get("number", ""))

                en_name = ""
                if is_chinese(cn_name):
                    if has_en_column:
                        val = row.get("en_name", "")
                        if pd.notna(val) and str(val).strip() != "":
                            en_name = str(val).strip()
                    if en_name == "":
                        en_name = chinese_to_pinyin(cn_name)

                prs = Presentation(io.BytesIO(pptx_bytes))
                fill_template(prs, {
                    "{{ID}}": uid,
                    "{{CN_NAME}}": cn_name,
                    "{{EN_NAME}}": en_name
                })
                buf = io.BytesIO()
                prs.save(buf)

                png_bytes = pptx_bytes_to_png(buf.getvalue())
                if png_bytes:
                    fname = f"{num}_{uid}_{cn_name}.png"
                    (work_dir / fname).write_bytes(png_bytes)
                    st.session_state.png_names.append(fname)
                    if batch_idx == 0 and j == 0:
                        st.session_state.first_png_name = fname
                        preview_spot.image(png_bytes, caption=fname, use_container_width=True)

                global_i = batch_idx + j + 1
                progress_bar.progress(global_i / total,
                                      text=f"正在处理 {global_i}/{total}")
                status_text.caption(f"当前：{cn_name}")

            # 清理僵尸 soffice 进程
            subprocess.run(["pkill", "-f", "soffice"], capture_output=True)

            if is_last_batch:
                progress_bar.empty()
                status_text.empty()
                st.session_state.trigger = False
            else:
                st.session_state.batch_idx = batch_end
            st.rerun()

        except Exception as e:
            st.session_state.trigger = False
            st.error(f"生成失败: {e}")

    elif st.session_state.work_dir and st.session_state.png_names:
        work_dir = Path(st.session_state.work_dir)
        png_names = st.session_state.png_names
        count = len(png_names)

        # 生成 ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in png_names:
                zf.write(work_dir / name, arcname=name)
        zip_buf.seek(0)

        st.download_button(
            label=f"⬇ 下载全部（{count} 张，ZIP）",
            data=zip_buf.getvalue(),
            file_name="certificates.zip",
            mime="application/zip",
            use_container_width=True
        )

        st.success(f"✅ 共生成 {count} 张证书图片")

        first_name = st.session_state.first_png_name
        if first_name:
            first_data = (work_dir / first_name).read_bytes()
            st.image(first_data, caption=first_name, use_container_width=True)

    else:
        st.markdown("""
        <div style="height:300px;display:flex;align-items:center;justify-content:center;
        color:#aaa;font-size:16px;text-align:center;border:2px dashed #ddd;border-radius:12px">
        👈 上传 CSV 和 PPTX 模板<br>点击生成后，证书预览将显示在这里
        </div>
        """, unsafe_allow_html=True)

st.divider()
st.caption("💡 姓氏多音字自动纠正（卜→Bǔ、单→Shàn、仇→Qiú 等）| 基于 pypinyin")
