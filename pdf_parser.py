import re
import pdfplumber

ROW_PATTERN = re.compile(r"^(\d{4,5})\s+(.+?)\s+(-\s*)?([\d,]+)$")
TOTAL_PATTERN = re.compile(r"^Revenue Total\s+(-\s*)?([\d,]+)$", re.IGNORECASE)


def parse_pdf_revenue(file):
    """Ambil baris section 'Revenue' (sebelum 'Revenue Total') dari PDF Trial Balance.
    Return: (list[(desc, amount)], revenue_total_di_pdf atau None)"""
    items, pdf_total, in_revenue = [], None, False

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                line = line.strip()

                total_match = TOTAL_PATTERN.match(line)
                if total_match:
                    pdf_total = int(total_match.group(2).replace(",", ""))
                    return items, pdf_total

                if line.lower() == "revenue":
                    in_revenue = True
                    continue

                if in_revenue:
                    m = ROW_PATTERN.match(line)
                    if m:
                        amount = int(m.group(4).replace(",", ""))
                        if m.group(3):
                            amount = -amount
                        items.append((m.group(2).strip(), amount))

    return items, pdf_total
