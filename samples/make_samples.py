"""Creates sample_form.pdf (blank non-editable form) and sample_master.xlsx next to this script.
Run:  python samples/make_samples.py
Then in the app: New from file -> sample_form.pdf, Import JSON -> sample_template.json,
Customer records -> Import Excel/CSV -> sample_master.xlsx (or sample_master.csv)."""
import csv
import os

import openpyxl
import pymupdf as fitz

HERE = os.path.dirname(os.path.abspath(__file__))
ROWS = [
    ["customer_name", "father_name", "cnic", "dob", "account_no", "account_type", "opening_amount",
     "address", "signature", "has_nominee", "nominee"],
    ["Muhammad Ali", "Ahmad Khan", "32203-1234567-1", "1995-03-14", "0012345678", "Savings", "25000",
     "House 12, Street 4, Chowk Azam, Layyah", "", "yes", "Sara Ali"],
    ["Ayesha Bibi", "Rashid Mehmood", "32203-7654321-2", "1988-11-02", "0012345679", "Current", "150000.5",
     "Mohallah Islam Pura, Layyah", "", "no", ""],
    ["Usman Tariq", "Tariq Hussain", "32203-1112223-3", "2001-07-30", "0012345680", "Term Deposit", "1000000",
     "Basti Malana, Tehsil Karor Lal Esan, District Layyah", "", "no", ""],
]


def form():
    doc = fitz.open()
    p = doc.new_page(width=595, height=842)
    p.insert_text((50, 60), "ACCOUNT OPENING FORM (SAMPLE)", fontsize=16)
    for label, x, y in [("Customer Name:", 50, 120), ("Father Name:", 50, 150), ("CNIC:", 50, 180),
                        ("Date of Birth:", 320, 180), ("Account No:", 50, 210), ("Account Type:", 50, 250),
                        ("Opening Amount (PKR):", 50, 290), ("Address:", 50, 330), ("Signature:", 50, 400)]:
        p.insert_text((x, y), label, fontsize=11)
    for x0, x1, y in [(150, 400, 122), (150, 400, 152), (150, 300, 182), (410, 540, 182), (150, 300, 212), (190, 330, 292)]:
        p.draw_line((x0, y), (x1, y), width=0.6)
    for x, lab in [(150, "Savings"), (240, "Current"), (330, "Term Deposit")]:
        p.draw_rect(fitz.Rect(x, 240, x + 12, 252), width=0.8)
        p.insert_text((x + 16, 250), lab, fontsize=10)
    p.draw_rect(fitz.Rect(150, 310, 540, 360), width=0.8)
    p.draw_rect(fitz.Rect(150, 380, 300, 420), width=0.8)
    q = doc.new_page(width=595, height=842)
    q.insert_text((50, 60), "FOR BRANCH USE ONLY", fontsize=16)
    for label, y in [("Branch:", 100), ("Customer:", 130), ("Nominee:", 160)]:
        q.insert_text((50, y), label, fontsize=11)
        q.draw_line((150, y + 2), (400, y + 2), width=0.6)
    doc.save(os.path.join(HERE, "sample_form.pdf"))


def master():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "customers"
    for r in ROWS:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=5):  # keep leading zeros: store as text
        for c in row:
            c.number_format = "@"
    wb.save(os.path.join(HERE, "sample_master.xlsx"))
    with open(os.path.join(HERE, "sample_master.csv"), "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(ROWS)


if __name__ == "__main__":
    form()
    master()
    print("Created sample_form.pdf, sample_master.xlsx, sample_master.csv in", HERE)
