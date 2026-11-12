#!/usr/bin/env python3
"""
تصحیح‌یار | اسکریپت تصحیح دسته‌ای
================================
خواندن پاسخ‌های چند دانش‌آموز از فایل Excel و تولید گزارش Word + خلاصه نمرات
"""

import json
import sys
import os
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from engine import TashihyarEngine, StudentReport

# Paths
BASE = Path(__file__).parent
ANSWER_KEY = BASE / "answer_key.json"
OUTPUT_DIR = BASE / "خروجی_گزارش‌ها"
OUTPUT_DIR.mkdir(exist_ok=True)


def parse_matching(text: str) -> dict:
    """Parse matching answer like '۱=ت ۲=الف ۳=پ ۴=ب' or '1=ت,2=الف'"""
    if not text:
        return {}
    result = {}
    # Normalize
    text = text.replace("۱", "1").replace("۲", "2").replace("۳", "3").replace("۴", "4")
    text = text.replace("=", "→").replace(":", "→")
    parts = text.replace(",", " ").replace("،", " ").split()
    for p in parts:
        if "→" in p:
            k, v = p.split("→", 1)
            result[k.strip()] = v.strip()
        elif "=" in p:
            k, v = p.split("=", 1)
            result[k.strip()] = v.strip()
    return result


def parse_table_or_structured(text: str) -> dict:
    """Parse table answers separated by | or /"""
    if not text:
        return {}
    parts = [p.strip() for p in text.replace("/", "|").split("|")]
    keys = ["natural", "human", "social", "example"]
    # For question 11
    keys11 = ["subject", "method", "goal"]
    
    if len(parts) >= 4:
        return dict(zip(keys, parts[:4]))
    elif len(parts) >= 3:
        return dict(zip(keys11, parts[:3]))
    elif len(parts) == 1:
        return {"raw": parts[0]}
    return {f"p{i}": p for i, p in enumerate(parts)}


def read_student_from_sheet(ws) -> dict:
    """Read one student's answers from a worksheet."""
    answers = {}
    for row in ws.iter_rows(min_row=2, max_col=6, values_only=True):
        if not row or not row[0]:
            continue
        qid = str(row[0]).strip()
        student_ans = row[4]  # column E
        
        if student_ans is None or str(student_ans).strip() == "":
            answers[qid] = None
            continue
        
        ans_str = str(student_ans).strip()
        
        # Type-specific parsing
        if qid.startswith("1"):  # true/false
            if ans_str in ["ص", "درست", "true", "True", "1", "✓"]:
                answers[qid] = True
            elif ans_str in ["غ", "نادرست", "false", "False", "0", "✗"]:
                answers[qid] = False
            else:
                answers[qid] = ans_str
        elif qid.startswith("3"):  # multiple choice
            # extract number
            ans_str = ans_str.replace("گزینه", "").strip()
            ans_str = ans_str.replace("۱", "1").replace("۲", "2").replace("۳", "3").replace("۴", "4")
            try:
                answers[qid] = int(ans_str[0]) if ans_str else None
            except:
                answers[qid] = ans_str
        elif qid == "4":  # matching
            answers[qid] = parse_matching(ans_str)
        elif qid in ["9a", "11"]:  # table / structured
            answers[qid] = parse_table_or_structured(ans_str)
        else:
            answers[qid] = ans_str

    # Ensure table-type answers are always dicts
    for qid in ["9a", "11"]:
        if qid in answers and not isinstance(answers[qid], dict):
            answers[qid] = parse_table_or_structured(str(answers[qid]) if answers[qid] else "")
    
    return answers


def process_excel(excel_path: str):
    """Main batch processing."""
    print(f"در حال خواندن فایل: {excel_path}")
    wb = load_workbook(excel_path, data_only=True)
    
    # Read student list
    if "لیست دانش‌آموزان" not in wb.sheetnames:
        print("خطا: برگه «لیست دانش‌آموزان» یافت نشد.")
        return
    
    ws_list = wb["لیست دانش‌آموزان"]
    students_meta = []
    for row in ws_list.iter_rows(min_row=2, max_col=4, values_only=True):
        if row[0] and row[1]:
            students_meta.append({
                "id": str(row[0]).strip(),
                "name": str(row[1]).strip()
            })
    
    if not students_meta:
        print("هیچ دانش‌آموزی در لیست یافت نشد. از برگه‌های موجود استفاده می‌شود.")
    
    engine = TashihyarEngine(str(ANSWER_KEY))
    all_reports = []
    
    # Process each student sheet
    for sheet_name in wb.sheetnames:
        if sheet_name in ["راهنما", "لیست دانش‌آموزان"]:
            continue
        
        print(f"  پردازش برگه: {sheet_name}")
        ws = wb[sheet_name]
        answers = read_student_from_sheet(ws)
        
        # Find name from list or use sheet name
        name = sheet_name
        sid = sheet_name
        for m in students_meta:
            if m["id"] == sheet_name or m["name"] == sheet_name:
                name = m["name"]
                sid = m["id"]
                break
        
        student_data = {
            "id": sid,
            "name": name,
            "answers": answers
        }
        
        report = engine.score_student(student_data)
        all_reports.append(report)
        
        # Generate individual Word report
        out_docx = OUTPUT_DIR / f"گزارش_{sid}_{name}.docx"
        engine.generate_word_report(report, str(out_docx))
        print(f"    → نمره: {report.total_score:.2f} | ذخیره: {out_docx.name}")
    
    # Generate summary Excel
    summary_path = OUTPUT_DIR / "خلاصه_نمرات_دسته.xlsx"
    create_summary_excel(all_reports, str(summary_path))
    print(f"\nجدول خلاصه نمرات: {summary_path}")
    print(f"تعداد دانش‌آموزان پردازش‌شده: {len(all_reports)}")
    print("پایان تصحیح دسته‌ای.")
    return all_reports


def create_summary_excel(reports: list, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "خلاصه نمرات"
    
    headers = ["ردیف", "شناسه", "نام و نام خانوادگی", "نمره تستی", "نمره تشریحی", "نمره کل", "تعداد Red Flags", "وضعیت بازبینی"]
    header_fill = PatternFill('solid', fgColor='1F4E79')
    header_font = Font(name='Tahoma', bold=True, color='FFFFFF', size=10)
    
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
    
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for i, r in enumerate(reports, 1):
        ws.cell(i+1, 1, i)
        ws.cell(i+1, 2, r.student_id)
        ws.cell(i+1, 3, r.name)
        ws.cell(i+1, 4, round(r.test_score, 2))
        ws.cell(i+1, 5, round(r.essay_score, 2))
        ws.cell(i+1, 6, round(r.total_score, 2))
        ws.cell(i+1, 7, len(r.red_flags))
        ws.cell(i+1, 8, "نیازمند بازبینی" if r.red_flags else "بدون مورد خاص")
        
        for col in range(1, 9):
            ws.cell(i+1, col).font = Font(name='Tahoma', size=10)
            ws.cell(i+1, col).border = thin
            ws.cell(i+1, col).alignment = Alignment(horizontal='center')
        
        # Highlight total score
        ws.cell(i+1, 6).fill = PatternFill('solid', fgColor='E2EFDA')
    
    # Column widths
    widths = [8, 12, 22, 12, 14, 12, 14, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64+i)].width = w
    
    # Note
    note_row = len(reports) + 3
    ws.cell(note_row, 1, "نمرات پیشنهادی هستند. تصمیم نهایی با دبیر است. موارد «نیازمند بازبینی» را با برگه فیزیکی کنترل کنید.")
    ws.cell(note_row, 1).font = Font(name='Tahoma', size=9, italic=True, color='C00000')
    
    wb.save(path)


def main():
    if len(sys.argv) < 2:
        # Default: look for the template or any xlsx in the folder
        candidates = list(BASE.glob("*.xlsx"))
        if not candidates:
            print("استفاده: python batch_score.py <مسیر_فایل_اکسل>")
            print("یا فایل Excel را در همین پوشه قرار دهید.")
            sys.exit(1)
        excel_path = str(candidates[0])
        print(f"فایل پیش‌فرض انتخاب شد: {excel_path}")
    else:
        excel_path = sys.argv[1]
    
    if not os.path.exists(excel_path):
        print(f"فایل یافت نشد: {excel_path}")
        sys.exit(1)
    
    process_excel(excel_path)


if __name__ == "__main__":
    main()
