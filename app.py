import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st


# 1. LOGIKA KATEGORI & PARSER
def get_category(description):
    desc_lower = description.lower()
    if "sugar & spice" in desc_lower:
        return "Sugar & Spice"
    elif "lobby lounge" in desc_lower:
        return "Lobby Lounge"
    elif "room service" in desc_lower:
        return "Room Service"
    elif "minibar" in desc_lower:
        return "Minibar"
    elif "banquet" in desc_lower:
        return "Banquet"
    elif "spa" in desc_lower:
        return "Spa"
    elif "laundry" in desc_lower:
        return "Guest Laundry"
    elif any(
        kw in desc_lower
        for kw in ["accomodation", "check-out", "checkout", "upsell"]
    ):
        return "Room Revenue"
    else:
        return "Other Revenue"


def parse_and_categorize(text):
    lines = [
        line.strip()
        for line in text.strip().split("\n")
        if line.strip() and line.strip().lower() != "revenue"
    ]
    descriptions, amounts = [], []
    amount_pattern = re.compile(r"^-?\s*[\d,]+$")

    for line in lines:
        if amount_pattern.match(line):
            clean_num = line.replace(" ", "").replace(",", "")
            amounts.append(int(clean_num))
        else:
            descriptions.append(line)

    categorized_data = []
    for desc, amount in zip(descriptions, amounts):
        cat = get_category(desc)
        categorized_data.append((cat, desc, amount))

    return categorized_data


def generate_excel(revenue_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Revenue Report"
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A5"

    NAVY, ICE_BLUE, BORDER_GRAY, ZEBRA_FILL = (
        "1B365D",
        "F2F5F9",
        "D9D9D9",
        "F9FAFC",
    )
    font_title = Font(name="Calibri", size=16, bold=True, color=NAVY)
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="555555")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)
    font_total = Font(name="Calibri", size=12, bold=True, color=NAVY)

    fill_header = PatternFill(
        start_color=NAVY, end_color=NAVY, fill_type="solid"
    )
    fill_zebra = PatternFill(
        start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid"
    )
    fill_total = PatternFill(
        start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )
    double_bottom_border = Border(
        top=Side(style="thin", color=NAVY),
        bottom=Side(style="double", color=NAVY),
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
    )

    ws.merge_cells("A1:C1")
    ws["A1"] = "InterContinental Jakarta Pondok Indah"
    ws["A1"].font = font_title

    ws.merge_cells("A2:C2")
    ws["A2"] = "Trial Balance - Categorized Revenue Section"
    ws["A2"].font = font_subtitle

    headers = ["Category", "Account Description", "Amount (IDR)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(
            horizontal=(
                "center"
                if col_idx == 1
                else ("left" if col_idx == 2 else "right")
            ),
            vertical="center",
        )
        cell.border = thin_border

    start_row = 5
    num_format = '#,##0;(#,##0);"-";@'

    for i, (category, desc, amount) in enumerate(revenue_data):
        row_num = start_row + i
        cell_cat = ws.cell(row=row_num, column=1, value=category)
        cell_desc = ws.cell(row=row_num, column=2, value=desc)
        cell_amount = ws.cell(row=row_num, column=3, value=amount)

        cell_cat.alignment = Alignment(horizontal="center")
        cell_desc.alignment = Alignment(horizontal="left")
        cell_amount.alignment = Alignment(horizontal="right")

        cell_cat.font = font_regular
        cell_desc.font = font_regular
        cell_amount.font = font_regular
        cell_amount.number_format = num_format

        if i % 2 == 1:
            cell_cat.fill = fill_zebra
            cell_desc.fill = fill_zebra
            cell_amount.fill = fill_zebra

        cell_cat.border = thin_border
        cell_desc.border = thin_border
        cell_amount.border = thin_border

    total_row = start_row + len(revenue_data)
    ws.cell(row=total_row, column=1, value="").border = double_bottom_border

    cell_label = ws.cell(row=total_row, column=2, value="Revenue Total")
    cell_label.font = font_total
    cell_label.fill = fill_total
    cell_label.alignment = Alignment(horizontal="right")
    cell_label.border = double_bottom_border

    cell_value = ws.cell(
        row=total_row, column=3, value=f"=SUM(C{start_row}:C{total_row-1})"
    )
    cell_value.font = font_total
    cell_value.fill = fill_total
    cell_value.alignment = Alignment(horizontal="right")
    cell_value.number_format = num_format
    cell_value.border = double_bottom_border

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 22

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# 2. TAMPILAN WEBSITE (STREAMLIT UI)
st.title("📊 Trial Balance Revenue Converter")
st.write("Paste data mentah di bawah ini untuk menghasilkan file Excel otomatis.")

raw_input = st.text_area("Paste Raw Data:", height=250)

if st.button("Generate Excel"):
    if raw_input.strip():
        data = parse_and_categorize(raw_input)
        excel_file = generate_excel(data)

        st.success(f"Berhasil memproses {len(data)} baris data!")
        st.download_button(
            label="📥 Download File Excel",
            data=excel_file,
            file_name="Revenue_Report_Categorized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Silakan paste data mentah terlebih dahulu.")
