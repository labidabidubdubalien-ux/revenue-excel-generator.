import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# Config Halaman Website
st.set_page_config(
    page_title="Revenue Converter & Categorizer", page_icon="📊", layout="wide"
)

# ==========================================
# 1. INITIAL STATE & DEFAULT CONFIGURATION
# ==========================================
DEFAULT_RULES = {
    "Room Revenue": {
        "account_codes": ["1000", "1060", "1070"],
        "keywords": ["accomodation", "check-out", "checkout", "upsell", "early check in"],
    },
    "Sugar & Spice": {
        "account_codes": [],
        "keywords": ["sugar & spice", "sugar and spice"],
    },
    "Lobby Lounge": {
        "account_codes": [],
        "keywords": ["lobby lounge"],
    },
    "Room Service": {
        "account_codes": [],
        "keywords": ["room service"],
    },
    "Minibar": {
        "account_codes": [],
        "keywords": ["minibar"],
    },
    "Banquet": {
        "account_codes": [],
        "keywords": ["banquet"],
    },
    "Spa": {
        "account_codes": [],
        "keywords": ["spa"],
    },
    "Guest Laundry": {
        "account_codes": [],
        "keywords": ["laundry", "guest laundry"],
    },
}

if "classification_rules" not in st.session_state:
    st.session_state.classification_rules = DEFAULT_RULES.copy()


# ==========================================
# 2. PARSER & DYNAMIC CLASSIFIER ENGINE
# ==========================================
def parse_raw_text(text):
    """
    Ekstraksi data mentah dari text clipboard menjadi list transaksi:
    (account_code, description, amount)
    """
    lines = [
        line.strip()
        for line in text.strip().split("\n")
        if line.strip() and line.strip().lower() != "revenue"
    ]

    descriptions_and_codes = []
    amounts = []
    amount_pattern = re.compile(r"^-?\s*[\d,]+$")

    for line in lines:
        if amount_pattern.match(line):
            clean_num = line.replace(" ", "").replace(",", "")
            amounts.append(int(clean_num))
        else:
            descriptions_and_codes.append(line)

    transactions = []
    for raw_desc, amount in zip(descriptions_and_codes, amounts):
        # Ekstraksi Account Code jika diawali oleh 3-6 digit angka
        code_match = re.match(r"^(\d{3,6})\s*-\s*(.*)", raw_desc)
        if code_match:
            acc_code = code_match.group(1).strip()
            desc = code_match.group(2).strip()
        else:
            acc_code = ""
            desc = raw_desc.strip()

        transactions.append({
            "account_code": acc_code,
            "description": desc,
            "amount": amount
        })

    return transactions


def classify_transaction(transaction, rules):
    """
    Menentukan kategori transaksi berdasarkan Hirarki Prioritas:
    1. Account Code Match
    2. Description Keyword Match
    3. Deteksi Conflict jika cocok lebih dari satu kategori
    4. Unclassified jika tidak ada cocok
    """
    acc_code = transaction["account_code"]
    desc_lower = transaction["description"].lower()

    # Prioritas 1: Matching via Account Code
    if acc_code:
        for category, rule in rules.items():
            if acc_code in rule.get("account_codes", []):
                return category, "MATCH_ACCOUNT_CODE"

    # Prioritas 2: Matching via Keywords
    matched_categories = []
    for category, rule in rules.items():
        for kw in rule.get("keywords", []):
            if kw.strip() and kw.strip().lower() in desc_lower:
                matched_categories.append(category)
                break

    # Prioritas 3 & 4: Evaluasi Hasil Pencocokan Keyword
    if len(matched_categories) == 1:
        return matched_categories[0], "MATCH_KEYWORD"
    elif len(matched_categories) > 1:
        return f"Conflict ({', '.join(matched_categories)})", "CONFLICT"
    else:
        return "Unclassified", "UNCLASSIFIED"


def process_all_transactions(transactions, rules):
    """
    Proses klasifikasi, agregasi subtotal, dan validasi rekonsilasi
    """
    categorized_items = []
    category_totals = {}
    unclassified_count = 0
    conflict_count = 0
    total_parsed_amount = 0

    for tx in transactions:
        cat, match_type = classify_transaction(tx, rules)
        
        if match_type == "UNCLASSIFIED":
            unclassified_count += 1
        elif match_type == "CONFLICT":
            conflict_count += 1

        categorized_items.append({
            "category": cat,
            "account_code": tx["account_code"],
            "description": tx["description"],
            "amount": tx["amount"],
            "status": match_type
        })

        category_totals[cat] = category_totals.get(cat, 0) + tx["amount"]
        total_parsed_amount += tx["amount"]

    return categorized_items, category_totals, total_parsed_amount, unclassified_count, conflict_count


# ==========================================
# 3. GENERATE EXCEL (SUMMARY & DETAIL SHEET)
# ==========================================
def generate_excel(revenue_data, category_totals):
    wb = openpyxl.Workbook()

    NAVY, ICE_BLUE, BORDER_GRAY, ZEBRA_FILL, ALERT_RED = (
        "1B365D",
        "F2F5F9",
        "D9D9D9",
        "F9FAFC",
        "FFF0F0"
    )
    font_title = Font(name="Calibri", size=16, bold=True, color=NAVY)
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="555555")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)
    font_total = Font(name="Calibri", size=12, bold=True, color=NAVY)
    font_unclassified = Font(name="Calibri", size=11, italic=True, color="9C0006")

    fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    fill_total = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type="solid")
    fill_alert = PatternFill(start_color=ALERT_RED, end_color=ALERT_RED, fill_type="solid")

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

    num_format = '#,##0;(#,##0);"-";@'

    # SHEET 1: SUMMARY
    ws_summary = wb.active
    ws_summary.title = "Category Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:B1")
    ws_summary["A1"] = "InterContinental Jakarta Pondok Indah"
    ws_summary["A1"].font = font_title

    ws_summary.merge_cells("A2:B2")
    ws_summary["A2"] = "Revenue Summary - Dynamic Categorization"
    ws_summary["A2"].font = font_subtitle

    summary_headers = ["Category", "Total Amount (IDR)"]
    for col_idx, header in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="left" if col_idx == 1 else "right", vertical="center")
        cell.border = thin_border

    ws_summary.row_dimensions[4].height = 25

    start_row_sum = 5
    for i, (cat_name, total_val) in enumerate(category_totals.items()):
        row_num = start_row_sum + i
        c_cat = ws_summary.cell(row=row_num, column=1, value=cat_name)
        c_val = ws_summary.cell(row=row_num, column=2, value=total_val)

        c_cat.font = font_unclassified if "Unclassified" in cat_name or "Conflict" in cat_name else font_regular
        c_val.font = font_regular
        c_val.number_format = num_format
        c_cat.alignment = Alignment(horizontal="left")
        c_val.alignment = Alignment(horizontal="right")

        if "Unclassified" in cat_name or "Conflict" in cat_name:
            c_cat.fill, c_val.fill = fill_alert, fill_alert
        elif i % 2 == 1:
            c_cat.fill, c_val.fill = fill_zebra, fill_zebra

        c_cat.border, c_val.border = thin_border, thin_border

    total_row_sum = start_row_sum + len(category_totals)
    ws_summary.cell(row=total_row_sum, column=1, value="Grand Total").font = font_total
    ws_summary.cell(row=total_row_sum, column=1).fill = fill_total
    ws_summary.cell(row=total_row_sum, column=1).alignment = Alignment(horizontal="right")
    ws_summary.cell(row=total_row_sum, column=1).border = double_bottom_border

    c_tot_val = ws_summary.cell(
        row=total_row_sum,
        column=2,
        value=f"=SUM(B{start_row_sum}:B{total_row_sum-1})",
    )
    c_tot_val.font = font_total
    c_tot_val.fill = fill_total
    c_tot_val.alignment = Alignment(horizontal="right")
    c_tot_val.number_format = num_format
    c_tot_val.border = double_bottom_border

    ws_summary.column_dimensions["A"].width = 35
    ws_summary.column_dimensions["B"].width = 25

    # SHEET 2: DETAIL
    ws_detail = wb.create_sheet(title="Transaction Detail")
    ws_detail.views.sheetView[0].showGridLines = True
    ws_detail.freeze_panes = "A5"

    ws_detail.merge_cells("A1:D1")
    ws_detail["A1"] = "InterContinental Jakarta Pondok Indah"
    ws_detail["A1"].font = font_title

    ws_detail.merge_cells("A2:D2")
    ws_detail["A2"] = "Trial Balance - Categorized Detail Section"
    ws_detail["A2"].font = font_subtitle

    headers = ["Category", "Account Code", "Account Description", "Amount (IDR)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws_detail.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(
            horizontal="center" if col_idx in [1, 2] else ("left" if col_idx == 3 else "right"),
            vertical="center",
        )
        cell.border = thin_border

    ws_detail.row_dimensions[4].height = 25

    start_row_det = 5
    for i, item in enumerate(revenue_data):
        row_num = start_row_det + i
        cell_cat = ws_detail.cell(row=row_num, column=1, value=item["category"])
        cell_code = ws_detail.cell(row=row_num, column=2, value=item["account_code"])
        cell_desc = ws_detail.cell(row=row_num, column=3, value=item["description"])
        cell_amount = ws_detail.cell(row=row_num, column=4, value=item["amount"])

        cell_cat.alignment = Alignment(horizontal="center")
        cell_code.alignment = Alignment(horizontal="center")
        cell_desc.alignment = Alignment(horizontal="left")
        cell_amount.alignment = Alignment(horizontal="right")

        cell_cat.font = font_unclassified if item["status"] in ["UNCLASSIFIED", "CONFLICT"] else font_regular
        cell_code.font, cell_desc.font, cell_amount.font = font_regular, font_regular, font_regular
        cell_amount.number_format = num_format

        if item["status"] in ["UNCLASSIFIED", "CONFLICT"]:
            cell_cat.fill, cell_code.fill, cell_desc.fill, cell_amount.fill = fill_alert, fill_alert, fill_alert, fill_alert
        elif i % 2 == 1:
            cell_cat.fill, cell_code.fill, cell_desc.fill, cell_amount.fill = fill_zebra, fill_zebra, fill_zebra, fill_zebra

        cell_cat.border, cell_code.border, cell_desc.border, cell_amount.border = thin_border, thin_border, thin_border, thin_border

    total_row_det = start_row_det + len(revenue_data)
    ws_detail.cell(row=total_row_det, column=1, value="").border = double_bottom_border
    ws_detail.cell(row=total_row_det, column=2, value="").border = double_bottom_border

    cell_label = ws_detail.cell(row=total_row_det, column=3, value="Revenue Total")
    cell_label.font = font_total
    cell_label.fill = fill_total
    cell_label.alignment = Alignment(horizontal="right")
    cell_label.border = double_bottom_border

    cell_value = ws_detail.cell(
        row=total_row_det,
        column=4,
        value=f"=SUM(D{start_row_det}:D{total_row_det-1})",
    )
    cell_value.font = font_total
    cell_value.fill = fill_total
    cell_value.alignment = Alignment(horizontal="right")
    cell_value.number_format = num_format
    cell_value.border = double_bottom_border

    ws_detail.column_dimensions["A"].width = 25
    ws_detail.column_dimensions["B"].width = 15
    ws_detail.column_dimensions["C"].width = 45
    ws_detail.column_dimensions["D"].width = 22

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 4. STREAMLIT NAVIGATION & UI
# ==========================================
st.sidebar.title("📌 Menu Navigasi")
page = st.sidebar.radio("Pilih Halaman:", ["🚀 Converter & Dashboard", "⚙️ Manage Classification Rules"])

# ------------------------------------------
# PAGE 1: CONVERTER & DASHBOARD
# ------------------------------------------
if page == "🚀 Converter & Dashboard":
    st.title("📊 Trial Balance Revenue Converter & Summarizer")
    st.write("Paste data mentah di bawah untuk memproses klasifikasi otomatis berdasarkan aturan dinamis.")

    raw_input = st.text_area("Paste Raw Data di Sini:", height=200)

    if st.button("🚀 Proses & Rekonsilasi Data"):
        if raw_input.strip():
            transactions = parse_raw_text(raw_input)
            data_detail, data_summary, total_amount, unclassified_cnt, conflict_cnt = process_all_transactions(
                transactions, st.session_state.classification_rules
            )

            # METRICS SUMMARY
            st.markdown("### 📊 Metrics & Data Integrity")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Transaksi", len(data_detail))
            m2.metric("Total Nominal (IDR)", f"{total_amount:,.0f}")
            m3.metric("Unclassified", unclassified_cnt, delta_color="inverse")
            m4.metric("Conflicts", conflict_cnt, delta_color="inverse")

            # VALIDASI & WARNINGS
            if unclassified_cnt > 0 or conflict_cnt > 0:
                st.warning(
                    f"⚠️ Ditemukan **{unclassified_cnt} Unclassified** dan **{conflict_cnt} Conflict**. "
                    "Anda dapat menambahkan keyword baru di menu **Manage Classification Rules** agar transaksi ini dapat dikenali secara otomatis."
                )
            else:
                st.success("✅ Semua transaksi berhasil terklasifikasi 100% tanpa error!")

            excel_file = generate_excel(data_detail, data_summary)

            st.download_button(
                label="📥 Download File Excel (Summary + Detail)",
                data=excel_file,
                file_name="Trial_Balance_Revenue_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.markdown("---")

            # VISUALISASI DASHBOARD
            col1, col2 = st.columns([1, 1])

            with col1:
                st.subheader("📋 Ringkasan Total per Kategori")
                df_summary = pd.DataFrame(
                    list(data_summary.items()), columns=["Kategori", "Total (IDR)"]
                )
                df_summary_formatted = df_summary.copy()
                df_summary_formatted["Total (IDR)"] = df_summary_formatted["Total (IDR)"].apply(lambda x: f"{x:,.0f}")
                st.dataframe(df_summary_formatted, use_container_width=True, hide_index=True)

            with col2:
                st.subheader("📈 Grafik Kontribusi Kategori")
                st.bar_chart(df_summary.set_index("Kategori"))

            st.markdown("---")

            # TABEL DETAIL
            st.subheader("📑 Detail Transaksi Lengkap")
            df_detail = pd.DataFrame(data_detail)
            df_detail_formatted = df_detail[["category", "account_code", "description", "amount", "status"]].copy()
            df_detail_formatted.columns = ["Category", "Account Code", "Description", "Amount (IDR)", "Status"]
            df_detail_formatted["Amount (IDR)"] = df_detail_formatted["Amount (IDR)"].apply(lambda x: f"{x:,.0f}")

            st.dataframe(df_detail_formatted, use_container_width=True, hide_index=True)

        else:
            st.warning("Silakan paste data mentah terlebih dahulu.")

# ------------------------------------------
# PAGE 2: MANAGE CLASSIFICATION RULES
# ------------------------------------------
elif page == "⚙️️ Manage Classification Rules":
    st.title("⚙️ Pengaturan Aturan Klasifikasi Kategori")
    st.write("Kelola daftar kategori, account code, dan keyword tanpa perlu mengubah kode Python.")

    rules = st.session_state.classification_rules

    st.subheader("➕ Tambah Kategori Baru")
    with st.form("add_category_form"):
        new_cat_name = st.text_input("Nama Kategori Baru:")
        new_cat_codes = st.text_input("Account Codes (pisahkan dengan koma):", help="Contoh: 1000, 1060, 1070")
        new_cat_kws = st.text_input("Keywords (pisahkan dengan koma):", help="Contoh: accomodation, checkout, upsell")
        submit_add = st.form_submit_button("Tambah Kategori")

        if submit_add:
            if new_cat_name.strip():
                if new_cat_name in rules:
                    st.error("Kategori tersebut sudah ada!")
                else:
                    rules[new_cat_name] = {
                        "account_codes": [c.strip() for c in new_cat_codes.split(",") if c.strip()],
                        "keywords": [k.strip() for k in new_cat_kws.split(",") if k.strip()]
                    }
                    st.session_state.classification_rules = rules
                    st.success(f"Kategori '{new_cat_name}' berhasil ditambahkan!")
                    st.rerun()
            else:
                st.error("Nama kategori tidak boleh kosong.")

    st.markdown("---")
    st.subheader("📝 Kelola Aturan Kategori Saat Ini")

    categories_to_delete = []

    for cat, rule in list(rules.items()):
        with st.expander(f"📌 {cat}", expanded=False):
            c1, c2 = st.columns(2)
            
            with c1:
                codes_str = ", ".join(rule.get("account_codes", []))
                updated_codes = st.text_input(f"Account Codes ({cat})", value=codes_str, key=f"code_{cat}")
            
            with c2:
                kws_str = ", ".join(rule.get("keywords", []))
                updated_kws = st.text_input(f"Keywords ({cat})", value=kws_str, key=f"kw_{cat}")

            col_btn1, col_btn2 = st.columns([1, 4])
            with col_btn1:
                if st.button(f"💾 Simpan {cat}", key=f"save_{cat}"):
                    rules[cat]["account_codes"] = [c.strip() for c in updated_codes.split(",") if c.strip()]
                    rules[cat]["keywords"] = [k.strip() for k in updated_kws.split(",") if k.strip()]
                    st.session_state.classification_rules = rules
                    st.success(f"Aturan untuk '{cat}' berhasil diperbarui!")
                    st.rerun()

            with col_btn2:
                if st.button(f"🗑️ Hapus Kategori {cat}", key=f"del_{cat}"):
                    categories_to_delete.append(cat)

    if categories_to_delete:
        for cat in categories_to_delete:
            del st.session_state.classification_rules[cat]
        st.success("Kategori berhasil dihapus.")
        st.rerun()

    st.markdown("---")
    if st.button("🔄 Reset ke Aturan Default"):
        st.session_state.classification_rules = DEFAULT_RULES.copy()
        st.success("Aturan berhasil dikembalikan ke default!")
        st.rerun()
