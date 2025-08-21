import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# ========= CONFIG =========
BASE_DIR = os.path.dirname(__file__)
TEMPLATE_IMG = os.path.join(BASE_DIR, "static", "excel_demo.png")
bg = ImageReader(TEMPLATE_IMG)


# ---- Numbers → Words ----
try:
    from num2words import num2words
    def number_to_words(n: int) -> str:
        return num2words(int(n)).title()
except Exception:
    def number_to_words(n: int) -> str:
        n = int(n)
        under_20 = ['Zero','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
                    'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
                    'Seventeen','Eighteen','Nineteen']
        tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']
        if n < 20: return under_20[n]
        if n < 100: return tens[n//10] + ('' if n%10==0 else ' ' + under_20[n%10])
        if n < 1000:
            rem = n % 100
            return under_20[n//100] + ' Hundred' + ('' if rem==0 else ' ' + number_to_words(rem))
        return str(n)

# ========= HELPERS =========
def _norm(s: str) -> str:
    return ''.join(str(s).lower().replace('\n', ' ').split())

def pick(colnames, *candidates):
    table = {_norm(c): c for c in colnames}
    for want in candidates:
        k = _norm(want)
        if k in table:
            return table[k]
    raise KeyError(f"None of {candidates} found in columns: {list(colnames)}")

# ========= SUBJECT ORDER =========
SUBJECTS = [
    ("Kannada",     "01"),
    ("English",     "02"),
    ("Chemistry",   "31"),
    ("Mathematics", "35"),
    ("Physics",     "33"),
    ("Biology",     "32"),
]

# ========= MAIN FUNCTION =========
def generate_individual_pdfs(excel_file, output_folder):
    """
    Reads an Excel sheet and generates one PDF per student inside output_folder.
    """
    df = pd.read_excel(excel_file)
    os.makedirs(output_folder, exist_ok=True)

    PAGE_W, PAGE_H = A4
    X_LEFT, X_RIGHT = 145, 470
    Y_BASE, LINE_GAP = 580, 40

    X_FIGURES, X_WORDS = 350, 445
    Y_FIRST_ROW, ROW_GAP = 405, 22

    X_TOTAL_FIGS, Y_TOTAL_ROW = 350, Y_FIRST_ROW - 7*ROW_GAP
    X_MARKS_WORDS, Y_MARKS_WORDS = 185, 210
    X_PERCENTAGE, Y_PERCENTAGE = 500, Y_FIRST_ROW - 7.3*ROW_GAP

    X_CLASS, Y_CLASS = 500, Y_FIRST_ROW - 9*ROW_GAP

    bg = ImageReader(TEMPLATE_IMG)

    generated_files = []

    for _, row in df.iterrows():
        cols = df.columns
        name   = str(row[pick(cols, "CandidateName", "Name")])
        mother = str(row[pick(cols, "MotherName", "Mother")])
        father = str(row[pick(cols, "FatherName", "Father")])
        regno  = str(row[pick(cols, "RegisterNumber", "RegNo", "Register No")])
        try:
            sats = str(row[pick(cols, "SATSNumber", "SATS", "SATS No", "SATSNo")])
        except:
            sats = ""

        marks = []
        for subj, _code in SUBJECTS:
            val = row[pick(cols, subj)]
            m = int(val) if not (pd.isna(val) or val == "") else 0
            marks.append(m)

        total = sum(marks)
        percentage = round((total / (len(SUBJECTS)*100))*100, 1)

        out_path = os.path.join(output_folder, f"{name}_{regno}.pdf")
        c = canvas.Canvas(out_path, pagesize=A4)
        c.drawImage(bg, 0, 0, width=PAGE_W, height=PAGE_H)

        c.setFont("Times-Bold", 13)
        c.drawString(X_LEFT,  Y_BASE,              name)
        c.drawString(X_LEFT,  Y_BASE - LINE_GAP,   mother)
        c.drawString(X_LEFT,  Y_BASE - 2*LINE_GAP, father)
        c.drawString(X_RIGHT, Y_BASE,              regno)
        if sats:
            c.drawString(X_RIGHT, Y_BASE - LINE_GAP, sats)

        # Marks Part I (Kannada & English)
        for i, m in enumerate(marks[:2]):
            y = Y_FIRST_ROW - i*ROW_GAP
            c.drawString(X_FIGURES, y, str(int(m)))
            c.drawString(X_WORDS,   y, number_to_words(int(m)))

        # Marks Part II (Chemistry onwards, shifted one row down)
        start_y = Y_FIRST_ROW - 2.2*ROW_GAP - ROW_GAP
        for j, m in enumerate(marks[2:]):
            y = start_y - j*ROW_GAP
            c.drawString(X_FIGURES, y, str(int(m)))
            c.drawString(X_WORDS,   y, number_to_words(int(m)))

        # Totals
        c.drawString(X_TOTAL_FIGS, Y_TOTAL_ROW-10, str(int(total)))
        c.drawString(X_MARKS_WORDS, Y_MARKS_WORDS, number_to_words(int(total)))
        c.drawString(X_PERCENTAGE, Y_PERCENTAGE, f"{percentage}%")

        # Class obtained
        class_obtained = "PASS" if all(m >= 35 for m in marks) else "FAIL"
        c.drawString(X_CLASS, Y_CLASS, class_obtained)

        c.save()
        generated_files.append(out_path)

    return generated_files
