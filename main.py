import os
import fitz  # PyMuPDF
import easyocr
from docx import Document
import pandas as pd
from tqdm import tqdm
import re

# ---------- CONFIG ----------
INPUT_PDF = r"book.pdf"
OUTPUT_DIR = r"output"
PAGES_TO_CONVERT = 20
# -----------------------------

# Create necessary folders
os.makedirs(OUTPUT_DIR, exist_ok=True)
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
DOCS_DIR = os.path.join(OUTPUT_DIR, "word")
EXCEL_DIR=os.path.join(OUTPUT_DIR,"excel")
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(DOCS_DIR, exist_ok=True)

# 1) Convert PDF pages to images using PyMuPDF
print("Converting PDF pages to images...")
doc = fitz.open(INPUT_PDF)
for i in range(min(PAGES_TO_CONVERT, len(doc))):
    page = doc[i]
    pix = page.get_pixmap(dpi=300)
    img_path = os.path.join(IMAGES_DIR, f"page_{i+1:02d}.png")
    pix.save(img_path)

# 2) OCR with EasyOCR (English + Bengali)
print("Running OCR (English + Bengali)...")
reader = easyocr.Reader(['bn','en'])
page_texts = []

# Arabic Unicode regex to skip Arabic lines
arabic_re = re.compile(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]')

def clean_text(text_lines):
    """Merge short lines, remove Arabic, empty lines, and extra spaces."""
    clean_lines = []
    buffer = ""
    for line in text_lines:
        line = line.strip()
        if not line or arabic_re.search(line):
            continue  # skip empty or Arabic lines
        # merge short lines into buffer
        if len(line) < 50:
            buffer += line + " "
        else:
            if buffer:
                clean_lines.append(buffer.strip())
                buffer = ""
            clean_lines.append(line)
    if buffer:
        clean_lines.append(buffer.strip())
    return clean_lines

def split_columns(paragraphs):
    """
    Split paragraphs into columns if ':' or '-' is present.
    Returns list of dicts for DataFrame rows.
    """
    rows = []
    for para in paragraphs:
        if ':' in para:
            parts = para.split(':', 1)
            rows.append({'Title': parts[0].strip(), 'Description': parts[1].strip(), 'FullText': para})
        elif '-' in para:
            parts = para.split('-', 1)
            rows.append({'Title': parts[0].strip(), 'Description': parts[1].strip(), 'FullText': para})
        else:
            rows.append({'Title': para, 'Description': '', 'FullText': para})
    return rows

# OCR each page
for i in tqdm(range(1, PAGES_TO_CONVERT+1), desc="Processing pages"):
    img_path = os.path.join(IMAGES_DIR, f"page_{i:02d}.png")
    results = reader.readtext(img_path, detail=0)
    paragraphs = clean_text(results)
    page_texts.append(paragraphs)

# 3) Save to Word
print("Creating Word file...")
docx_file = Document()
for i, paragraphs in enumerate(page_texts, start=1):
    heading = docx_file.add_heading(f"Page {i}", level=2)
    # Make heading bold
    heading.runs[0].bold = True
    for para in paragraphs:
        docx_file.add_paragraph(para)
    docx_file.add_page_break()

word_out = os.path.join(DOCS_DIR, "Output.docx")
docx_file.save(word_out)
print(f"Word file saved: {word_out}")

# 4) Save to Excel
print("Saving text to Excel...")
all_rows = []
for i, paragraphs in enumerate(page_texts, start=1):
    rows = split_columns(paragraphs)
    for row in rows:
        row['PageNumber'] = i
        all_rows.append(row)

df = pd.DataFrame(all_rows, columns=['PageNumber', 'Title', 'Description', 'FullText'])
excel_out = os.path.join(EXCEL_DIR, "Output.xlsx")
df.to_excel(excel_out, index=False)
print(f"Excel file saved: {excel_out}")

print("OCR extraction completed successfully.")
