import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_IMG = os.path.join(BASE_DIR, "static", "template.png")
bg = ImageReader(TEMPLATE_IMG)

# Number to words
try:
    from num2words import num2words
    def number_to_words(n: int) -> str:
        return num2words(int(n)).title()
except:
    def number_to_words(n: int) -> str:
        return str(n)

def _norm(s: str) -> str:
    return ''.join(str(s).lower().replace('\n', ' ').split())

def pick(colnames, *candidates):
    table = {_norm(c): c for c in colnames}
    for want in candidates:
        if _norm(want) in table:
            return table[_norm(want)]
    raise KeyError(f"Columns missing. Expected one of: {candidates}")

# Subjects
SUBJECTS = [
    ("Kannada",     "01"),
    ("English",     "02"),
    ("Chemistry",   "31"),
    ("Mathematics", "35"),
    ("Physics",     "33"),
    ("Biology",     "32"),
]

def generate_individual_pdfs(excel_file, output_folder):
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

        # Marks
        marks = []
        for subj, _ in SUBJECTS:
            raw_val = str(row.get(pick(cols, subj), "")).strip()
            if raw_val.upper() in ["AB", "A", "ABSENT"]:
                m = 0
            else:
                try:
                    m = int(float(raw_val)) if raw_val else 0
                except:
                    m = 0
            marks.append(m)

        total = sum(marks)
        percentage = round(total / len(SUBJECTS), 1)

        # PDF
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

        # Marks Part I (Kannada, English)
        for i, m in enumerate(marks[:2]):
            y = Y_FIRST_ROW - i*ROW_GAP
            c.drawString(X_FIGURES, y, str(m))
            c.drawString(X_WORDS,   y, number_to_words(m))

        # Marks Part II (Chemistry onwards)
        GAP_AFTER_PART1 = 2  # leave two empty rows after English
        for j, m in enumerate(marks[2:]):
            y = Y_FIRST_ROW - (3.2 + j) * ROW_GAP
            c.drawString(X_FIGURES, y, str(m))
            c.drawString(X_WORDS,   y, number_to_words(m))        # Totals & Result


        c.drawString(X_TOTAL_FIGS, Y_TOTAL_ROW-10, str(total))
        c.drawString(X_MARKS_WORDS, Y_MARKS_WORDS, number_to_words(total))
        c.drawString(X_PERCENTAGE, Y_PERCENTAGE, f"{percentage}%")
        class_obtained = "PASS" if all(m >= 35 for m in marks) else "FAIL"
        c.drawString(X_CLASS, Y_CLASS, class_obtained)

        c.save()
        generated_files.append(out_path)

    return generated_files
