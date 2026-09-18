from __future__ import annotations

import csv
import io
import re
import tempfile
from pathlib import Path
from typing import Any

import openpyxl
import pdfplumber
import streamlit as st
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas


APP_DIR = Path(__file__).resolve().parent
MAPPING_XLSX = APP_DIR / "data" / "mapping.xlsx"
SKU_RE = re.compile(r"\b(?:[A-Z]\d{7,12}|\d{10,14})(?:-\d+)?\b")
DEFAULT_FONT = "STSong-Light"


def safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def pick_mapping_sheet(workbook: openpyxl.Workbook):
    if "Sheet2" in workbook.sheetnames:
        return workbook["Sheet2"]
    if not workbook.sheetnames:
        raise ValueError("Excel 中没有可读取的工作表。")
    return workbook[workbook.sheetnames[0]]


def mapping_start_row(ws) -> int:
    first_code = safe_text(ws.cell(row=1, column=1).value)
    first_name = safe_text(ws.cell(row=1, column=2).value)
    if "商品编码" in first_code or "商品名称" in first_name:
        return 2
    return 1


@st.cache_data(show_spinner=False)
def load_mapping_excel(path: str) -> dict[str, str]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = pick_mapping_sheet(wb)
    mapping: dict[str, str] = {}
    for code_raw, name_raw in ws.iter_rows(min_row=mapping_start_row(ws), min_col=1, max_col=2, values_only=True):
        code = safe_text(code_raw)
        name = safe_text(name_raw)
        if code and name:
            mapping[code] = name
    if not mapping:
        raise ValueError("Excel 中没有读取到有效的商品编码和商品名称。")
    return mapping


def save_uploaded_pdf(uploaded_file) -> Path:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        return Path(tmp.name)


def annotated_download_name(uploaded_name: str) -> str:
    original = Path(uploaded_name or "JIT面单.pdf").stem
    return f"{original}注释.pdf"


def extract_code_from_page(page: pdfplumber.page.Page) -> str:
    text = page.extract_text() or ""
    matches = SKU_RE.findall(text)
    letter_codes = [match for match in matches if re.match(r"^[A-Z]", match)]
    if letter_codes:
        return letter_codes[0]
    numeric_codes = [
        match
        for match in matches
        if re.match(r"^\d{10,14}(?:-\d+)?$", match) and not match.startswith("120000")
    ]
    return numeric_codes[0] if numeric_codes else ""


def analyze_pdf(pdf_path: Path, mapping: dict[str, str]) -> dict[str, Any]:
    reader = PdfReader(str(pdf_path))
    page_count = len(reader.pages)
    errors: list[str] = []
    warnings: list[str] = []

    if page_count == 0:
        errors.append("PDF 没有页面。")
    if page_count % 2 != 0:
        errors.append(f"PDF 页数为 {page_count}，不是偶数；需要按“揽收面单 + 货品标签”两页一组处理。")

    sizes = [
        (round(float(page.mediabox.width), 2), round(float(page.mediabox.height), 2))
        for page in reader.pages
    ]
    unique_sizes = sorted(set(sizes))
    if len(unique_sizes) > 1:
        warnings.append("PDF 内页面尺寸不完全一致；系统会按每组揽收面单尺寸插入注释页。")

    rows: list[dict[str, Any]] = []
    order_count = page_count // 2
    seen_codes: dict[str, int] = {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for order_idx in range(order_count):
            label_page_index = order_idx * 2 + 1
            raw_code = extract_code_from_page(pdf.pages[label_page_index])
            lookup_code = raw_code.split("-", 1)[0] if raw_code else ""
            name = mapping.get(raw_code) or mapping.get(lookup_code) or ""
            status = "通过"
            notes: list[str] = []

            if not raw_code:
                status = "需处理"
                notes.append("货品标签页未识别到商品编码")
            elif raw_code != lookup_code and lookup_code in mapping:
                notes.append(f"标签编码为 {raw_code}，已按主编码 {lookup_code} 匹配")
            if raw_code and not name:
                status = "需处理"
                notes.append("对应表未找到商品名称")
            if lookup_code:
                seen_codes[lookup_code] = seen_codes.get(lookup_code, 0) + 1

            rows.append(
                {
                    "订单": order_idx + 1,
                    "揽收页": order_idx * 2 + 1,
                    "标签页": label_page_index + 1,
                    "标签编码": raw_code,
                    "匹配编码": lookup_code,
                    "商品名称": name,
                    "状态": status,
                    "备注": "；".join(notes),
                    "输出注释页": order_idx * 3 + 1,
                }
            )

    repeated = sorted(code for code, count in seen_codes.items() if count > 1)
    if repeated:
        warnings.append("同一个商品编码在本 PDF 中出现多次：" + "、".join(repeated))
    if any(row["状态"] != "通过" for row in rows):
        errors.append("有订单未通过复核，请先更新对应表或检查 PDF。")

    return {
        "page_count": page_count,
        "order_count": order_count,
        "expected_output_pages": order_count * 3,
        "errors": errors,
        "warnings": warnings,
        "rows": rows,
        "passed": not errors,
    }


def register_font() -> str:
    try:
        pdfmetrics.getFont(DEFAULT_FONT)
    except KeyError:
        pdfmetrics.registerFont(UnicodeCIDFont(DEFAULT_FONT))
    return DEFAULT_FONT


def wrap_text(text: str, font_name: str, font_size: float, max_width: float) -> list[str]:
    lines: list[str] = []
    current = ""
    for char in text:
        trial = current + char
        if pdfmetrics.stringWidth(trial, font_name, font_size) <= max_width:
            current = trial
            continue
        if current:
            lines.append(current)
        current = char
    if current:
        lines.append(current)
    return lines


def make_note_page(width: float, height: float, code: str, name: str) -> PdfReader:
    font_name = register_font()
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    c.setFillColor(HexColor("#FFFFFF"))
    c.rect(0, 0, width, height, fill=1, stroke=0)

    margin = 12
    max_width = width - margin * 2
    title_size = 16
    body_size = 13
    label_size = 10
    name_lines = wrap_text(name, font_name, body_size, max_width)
    while len(name_lines) > 7 and body_size > 7:
        body_size -= 0.4
        name_lines = wrap_text(name, font_name, body_size, max_width)

    y = height - 20
    c.setFillColor(HexColor("#111111"))
    c.setFont(font_name, title_size)
    c.drawString(margin, y, "商品名称注释")

    y -= 21
    c.setFont(font_name, label_size)
    c.setFillColor(HexColor("#555555"))
    c.drawString(margin, y, "商品编码")
    y -= 14
    c.setFont(font_name, 13)
    c.setFillColor(HexColor("#111111"))
    c.drawString(margin, y, code or "未识别")

    y -= 22
    c.setFont(font_name, label_size)
    c.setFillColor(HexColor("#555555"))
    c.drawString(margin, y, "商品名称")
    y -= 15
    c.setFont(font_name, body_size)
    c.setFillColor(HexColor("#111111"))
    for line in name_lines:
        if y < margin:
            break
        c.drawString(margin, y, line)
        y -= body_size + 2.4

    c.showPage()
    c.save()
    buffer.seek(0)
    return PdfReader(buffer)


def create_annotated_pdf(pdf_path: Path, review: dict[str, Any]) -> bytes:
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for row in review["rows"]:
        pickup_index = (row["订单"] - 1) * 2
        pickup_page = reader.pages[pickup_index]
        note_pdf = make_note_page(
            float(pickup_page.mediabox.width),
            float(pickup_page.mediabox.height),
            row["匹配编码"],
            row["商品名称"],
        )
        writer.add_page(note_pdf.pages[0])
        writer.add_page(pickup_page)
        writer.add_page(reader.pages[pickup_index + 1])
    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def review_csv_bytes(review: dict[str, Any]) -> bytes:
    buffer = io.StringIO()
    fieldnames = ["订单", "揽收页", "标签页", "标签编码", "匹配编码", "商品名称", "状态", "备注", "输出注释页"]
    writer = csv.DictWriter(buffer, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(review["rows"])
    return buffer.getvalue().encode("utf-8-sig")


st.set_page_config(page_title="JIT 面单商品名称注释", page_icon="📦", layout="wide")
st.title("JIT 面单商品名称注释")
st.caption("上传速卖通 JIT 面单 PDF，系统会先复核商品编码，再生成已插入商品名称注释页的新 PDF。")

try:
    mapping = load_mapping_excel(str(MAPPING_XLSX))
except Exception as exc:
    mapping = {}
    st.error(f"对应表读取失败：{exc}")

with st.sidebar:
    st.header("对应表")
    if mapping:
        st.success(f"当前对应表：mapping.xlsx，{len(mapping)} 个编码")
    else:
        st.error("当前没有可用对应表")
    st.write("长期更新：替换 GitHub 里的 data/mapping.xlsx。A 列是商品编码，B 列是商品名称。")

uploaded_pdf = st.file_uploader("面单 PDF", type=["pdf"])

if not uploaded_pdf:
    st.info("请先上传面单 PDF。")
    st.stop()

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

for warning in review["warnings"]:
    st.warning(warning)
for error in review["errors"]:
    st.error(error)
if review["passed"]:
    st.success(f"复核通过，预计输出 {review['expected_output_pages']} 页。")

st.subheader("复核清单")
st.dataframe(review["rows"], use_container_width=True, hide_index=True)

st.download_button(
    "下载复核清单 CSV",
    review_csv_bytes(review),
    file_name="JIT面单复核清单.csv",
    mime="text/csv",
)

if review["passed"]:
    with st.spinner("正在生成 PDF"):
        output_pdf = create_annotated_pdf(pdf_path, review)
    st.download_button(
        "下载已加商品名称注释的 PDF",
        output_pdf,
        file_name=annotated_download_name(uploaded_pdf.name),
        mime="application/pdf",
        type="primary",
    )
