from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from jit_processor import (
    analyze_pdf,
    create_annotated_pdf,
    load_mapping_excel,
    load_mapping_file,
    review_csv_bytes,
)


APP_DIR = Path(__file__).resolve().parent
MAPPING_XLSX = APP_DIR / "data" / "mapping.xlsx"
MAPPING_CSV = APP_DIR / "data" / "mapping.csv"


st.set_page_config(page_title="JIT 面单商品名称注释", page_icon="📦", layout="wide")

st.title("JIT 面单商品名称注释")
st.caption("上传速卖通 JIT 面单 PDF，系统会先复核商品编码，再生成已插入商品名称注释页的新 PDF。")


@st.cache_data(show_spinner=False)
def get_default_mapping(path: str) -> dict[str, str]:
    return load_mapping_file(Path(path))


def get_mapping_source() -> Path:
    if MAPPING_XLSX.exists():
        return MAPPING_XLSX
    return MAPPING_CSV


def save_uploaded_pdf(uploaded_file) -> Path:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        return Path(tmp.name)


try:
    mapping_source = get_mapping_source()
    mapping = get_default_mapping(str(mapping_source))
except Exception as exc:
    mapping_source = get_mapping_source()
    mapping = {}
    st.error(f"内置对应表读取失败：{exc}")

with st.sidebar:
    st.header("对应表")
    if mapping:
        st.success(f"当前对应表：{mapping_source.name}，{len(mapping)} 个编码")
    else:
        st.error("当前没有可用对应表")

    with st.expander("更新对应表文件"):
        st.write("长期更新：把新的 Excel 对应表上传到 GitHub 的 data/mapping.xlsx。Streamlit 重新部署后会自动读取。")
        st.write("临时测试：这里上传 Excel，只在本次页面会话中使用。")
        excel_file = st.file_uploader("Excel 对应表", type=["xlsx"], key="mapping_excel")
        if excel_file:
            with st.spinner("正在提取商品编码和商品名称"):
                try:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        tmp.write(excel_file.getvalue())
                        temp_excel_path = Path(tmp.name)
                    mapping = load_mapping_excel(temp_excel_path)
                    temp_excel_path.unlink(missing_ok=True)
                    st.success(f"本次临时对应表已生效：{len(mapping)} 个编码。")
                except Exception as exc:
                    st.error(f"提取失败：{exc}")


uploaded_pdf = st.file_uploader("面单 PDF", type=["pdf"])

if uploaded_pdf:
    if not mapping:
        st.stop()

    pdf_path = save_uploaded_pdf(uploaded_pdf)
    try:
        with st.spinner("正在复核商品编码和对应表"):
            review = analyze_pdf(pdf_path, mapping)
    except Exception as exc:
        st.error(f"PDF 读取失败：{exc}")
        st.stop()

    left, right = st.columns(2)
    left.metric("源 PDF 页数", review["page_count"])
    right.metric("识别订单数", review["order_count"])

    if review["warnings"]:
        for warning in review["warnings"]:
            st.warning(warning)
    if review["errors"]:
        for error in review["errors"]:
            st.error(error)
    else:
        st.success(f"复核通过，预计输出 {review['expected_output_pages']} 页。")

    st.subheader("复核清单")
    st.dataframe(review["rows"], use_container_width=True, hide_index=True)

    csv_bytes = review_csv_bytes(review)
    st.download_button(
        "下载复核清单 CSV",
        csv_bytes,
        file_name="JIT面单复核清单.csv",
        mime="text/csv",
        use_container_width=False,
    )

    if review["passed"]:
        with st.spinner("正在生成 PDF"):
            try:
                output_pdf = create_annotated_pdf(pdf_path, review)
            except Exception as exc:
                st.error(f"生成失败：{exc}")
                st.stop()
        st.download_button(
            "下载已加商品名称注释的 PDF",
            output_pdf,
            file_name="JIT面单_已加商品名称注释.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=False,
        )
else:
    st.info("请先上传面单 PDF。")
