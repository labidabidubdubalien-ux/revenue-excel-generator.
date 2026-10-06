import re
import pdfplumber


def extract_revenue_from_pdf(pdf_file):
    raw_lines = []

    # Buka PDF dan baca setiap halaman
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                raw_lines.extend(text.split("\n"))

    revenue_items = []
    is_revenue_section = False

    for line in raw_lines:
        line_clean = line.strip()
        if not line_clean:
            continue

        # Deteksi Awal Section Revenue
        if line_clean.lower() == "revenue":
            is_revenue_section = True
            continue

        # Deteksi Akhir Section Revenue (Stop Parsing)
        if (
            "revenue total" in line_clean.lower()
            or "non revenue" in line_clean.lower()
        ):
            is_revenue_section = False
            break

        # Proses Baris Dalam Section Revenue
        if is_revenue_section:
            # Hilangkan karakter pipe '|' jika ada hasil pembacaan PDF
            clean_str = line_clean.replace("|", "").strip()

            # Mencari angka nominal di paling akhir baris (termasuk format - 165,888)
            match_amount = re.search(r"(-?\s*[\d,]+(?:\.\d+)?)$", clean_str)

            if match_amount:
                amount_str = match_amount.group(1)
                # Sisa teks di sebelah kiri angka adalah Kode Akun & Deskripsi
                left_text = clean_str[: match_amount.start()].strip()

                # Pisahkan Kode Akun (jika ada angka di paling depan) dan Deskripsi
                match_code = re.match(r"^(\d+)\s+(.*)$", left_text)
                if match_code:
                    account_code = match_code.group(1)
                    desc = match_code.group(2).strip()
                else:
                    account_code = ""
                    desc = left_text.strip()

                # Bersihkan format angka ke integer
                try:
                    num_clean = (
                        amount_str.replace(" ", "").replace(",", "").strip()
                    )
                    amount = int(float(num_clean))

                    # Pastikan baris bukan header/label berulang
                    if desc and desc.lower() != "revenue":
                        revenue_items.append((account_code, desc, amount))
                except ValueError:
                    continue

    return revenue_items
