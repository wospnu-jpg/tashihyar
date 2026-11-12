# -*- coding: utf-8 -*-
"""
تصحیح‌یار – نسخه ۲
ورود نتایج JSON (خروجی جمینای/Claude) ← بازبینی دبیر ← جمع نمره با کد ← خروجی Excel و Word
"""
import hashlib
import io
import json
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

try:
    import pymupdf as fitz  # PyMuPDF
except Exception:  # pragma: no cover
    try:
        import fitz
    except Exception:
        fitz = None

# ───────────────────────── توابع کمکی داده ─────────────────────────

_DIGITS = {ord(c): str(i) for i, c in enumerate("۰۱۲۳۴۵۶۷۸۹")}
_DIGITS.update({ord(c): str(i) for i, c in enumerate("٠١٢٣٤٥٦٧٨٩")})
_DIGITS.update({ord("٫"): ".", ord("/"): ".", ord("،"): ".", ord(","): ".", ord("٬"): ""})


def to_float(x):
    """تبدیل نمره به عدد؛ None برای «در انتظار دبیر» یا مقدار خالی."""
    if x is None:
        return None
    if isinstance(x, bool):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().translate(_DIGITS)
    if not s or s.lower() in ("null", "none", "nan") or "انتظار" in s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_bool(x):
    if isinstance(x, bool):
        return x
    return str(x).strip().lower() in ("true", "1", "yes", "بله", "درست")


def extract_json_objects(text):
    """همه‌ی آبجکت‌های JSON دارای کلید answers را از یک متن (حتی با متن اضافه) بیرون می‌کشد."""
    dec = json.JSONDecoder()
    out, i = [], 0
    while True:
        i = text.find("{", i)
        if i == -1:
            break
        try:
            obj, end = dec.raw_decode(text[i:])
        except json.JSONDecodeError:
            i += 1
            continue
        if isinstance(obj, dict) and isinstance(obj.get("answers"), list):
            out.append(obj)
            i += end
        else:
            i += 1
    return out


def normalize_student(obj, page_start):
    st_info = obj.get("student", {})
    if isinstance(st_info, str):
        st_info = {"name": st_info}
    answers = []
    for a in obj.get("answers", []):
        if not isinstance(a, dict):
            continue
        score = to_float(a.get("score"))
        mx = to_float(a.get("max")) or 0.0
        needs = to_bool(a.get("needs_review", False)) or score is None
        conf = to_float(a.get("confidence"))
        answers.append({
            "q": str(a.get("q", "")),
            "page": int(to_float(a.get("page")) or 1),
            "verbatim": str(a.get("verbatim", a.get("extracted", "")) or ""),
            "interpretation": str(a.get("interpretation", "") or ""),
            "confidence": conf,
            "legibility": str(a.get("legibility", "") or ""),
            "suggested": to_float(a.get("suggested_score")),
            "score": score,
            "max": mx,
            "reason": str(a.get("reason", "") or ""),
            "needs_review": needs,
            "final": score,
            "note": "",
        })
    student = {
        "id": hashlib.md5((json.dumps(obj, ensure_ascii=False, sort_keys=True)
                           + str(datetime.now().timestamp())).encode()).hexdigest()[:10],
        "name": str(st_info.get("name", "") or "بدون نام"),
        "father": str(st_info.get("father", "") or ""),
        "cls": str(st_info.get("class", "") or ""),
        "page_start": page_start,
        "answers": answers,
        "ver": 0,
    }
    student["base"] = [(a["final"], a["note"]) for a in answers]
    return student


def totals(s):
    fin = [a["final"] for a in s["answers"]]
    total = round(sum(f for f in fin if f is not None), 2)
    pending = sum(1 for f in fin if f is None)
    maxsum = round(sum(a["max"] for a in s["answers"]), 2)
    flagged = sum(1 for a in s["answers"] if a["needs_review"])
    return total, pending, maxsum, flagged


def clamp(f, mx):
    if f is None:
        return None
    f = max(0.0, float(f))
    if mx and f > mx:
        f = mx
    return round(f, 2)


# ───────────────────────── تصویر صفحات PDF ─────────────────────────

@st.cache_data(show_spinner=False, max_entries=64)
def render_page(digest, page_number, zoom, _pdf_bytes):
    doc = fitz.open(stream=_pdf_bytes, filetype="pdf")
    try:
        if page_number < 1 or page_number > doc.page_count:
            return None
        page = doc.load_page(page_number - 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        return pix.tobytes("png")
    finally:
        doc.close()


def pdf_page_count(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = doc.page_count
    doc.close()
    return n


# ───────────────────────── خروجی Excel ─────────────────────────

def _safe_sheet_name(name, used):
    bad = '[]:*?/\\'
    n = "".join("_" if c in bad else c for c in name)[:28] or "student"
    base, k = n, 2
    while n in used:
        n = f"{base[:25]}_{k}"
        k += 1
    used.add(n)
    return n


def build_excel(students, exam_title, exam_total):
    wb = Workbook()
    ws = wb.active
    ws.title = "خلاصه"
    ws.sheet_view.rightToLeft = True
    head_fill = PatternFill("solid", fgColor="DCE6F1")
    ws.append([exam_title])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])
    ws.append(["ردیف", "نام", "نام پدر", "کلاس", "نمره نهایی", "بارم کل آزمون",
               "موارد در انتظار", "موارد نیازمند بازبینی"])
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = head_fill
    for i, s in enumerate(students, 1):
        total, pending, _mx, flagged = totals(s)
        ws.append([i, s["name"], s["father"], s["cls"], total, exam_total, pending, flagged])
    for col, w in zip("ABCDEFGH", (7, 28, 18, 14, 12, 14, 16, 20)):
        ws.column_dimensions[col].width = w

    used = {"خلاصه"}
    for s in students:
        sh = wb.create_sheet(_safe_sheet_name(s["name"], used))
        sh.sheet_view.rightToLeft = True
        sh.append(["سؤال", "متن عینی دانش‌آموز", "تفسیر مدل", "اطمینان", "پیشنهاد مدل",
                   "نمره نهایی", "بارم", "دلیل مدل", "یادداشت دبیر"])
        for c in sh[1]:
            c.font = Font(bold=True)
            c.fill = head_fill
        for a in s["answers"]:
            sh.append([a["q"], a["verbatim"], a["interpretation"], a["confidence"],
                       a["suggested"], a["final"], a["max"], a["reason"], a["note"]])
        total, pending, _mx, _fl = totals(s)
        sh.append([])
        sh.append(["جمع نمره‌ی نهایی", "", "", "", "", total, "", f"در انتظار: {pending}"])
        sh.cell(row=sh.max_row, column=1).font = Font(bold=True)
        sh.cell(row=sh.max_row, column=6).font = Font(bold=True)
        for col, w in zip("ABCDEFGHI", (10, 50, 30, 10, 12, 12, 8, 40, 30)):
            sh.column_dimensions[col].width = w
        for row in sh.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ───────────────────────── خروجی Word ─────────────────────────

def _rtl_par(par):
    pPr = par._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:bidi"))
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _add_run(par, text, bold=False, size=11):
    run = par.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Tahoma"
    rPr = run._r.get_or_add_rPr()
    rPr.get_or_add_rFonts().set(qn("w:cs"), "Tahoma")
    rPr.append(OxmlElement("w:rtl"))
    return run


def _par(doc, text="", bold=False, size=11):
    p = doc.add_paragraph()
    _rtl_par(p)
    if text:
        _add_run(p, text, bold, size)
    return p


def _cell(cell, text, bold=False, size=10):
    p = cell.paragraphs[0]
    _rtl_par(p)
    _add_run(p, text, bold, size)


def add_student_to_doc(doc, s, exam_title, exam_total):
    total, pending, _mx, _fl = totals(s)
    _par(doc, f"گزارش تصحیح: {s['name']}", bold=True, size=15)
    meta = f"{exam_title}"
    if s["father"]:
        meta += f" | نام پدر: {s['father']}"
    if s["cls"]:
        meta += f" | کلاس: {s['cls']}"
    _par(doc, meta, size=10)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    tblPr = table._tbl.tblPr
    tblPr.append(OxmlElement("w:bidiVisual"))
    for cell, h in zip(table.rows[0].cells, ["سؤال", "پاسخ دانش‌آموز", "نمره", "بارم", "یادداشت دبیر"]):
        _cell(cell, h, bold=True)
    for a in s["answers"]:
        row = table.add_row().cells
        fin = "در انتظار" if a["final"] is None else f"{a['final']:g}"
        _cell(row[0], a["q"])
        _cell(row[1], a["verbatim"] or "—")
        _cell(row[2], fin)
        _cell(row[3], f"{a['max']:g}")
        _cell(row[4], a["note"])
    for row in table.rows:
        for cell, w in zip(row.cells, (Cm(1.6), Cm(8.2), Cm(1.8), Cm(1.6), Cm(4.0))):
            cell.width = w

    _par(doc)
    _par(doc, f"مجموع نمره‌ی نهایی: {total:g} از {exam_total:g}", bold=True, size=12)
    if pending:
        _par(doc, f"توجه: {pending} مورد هنوز نمره‌ی نهایی ندارد و در مجموع لحاظ نشده است.", size=10)
    flags = [a for a in s["answers"] if a["needs_review"]]
    if flags:
        _par(doc, "موارد بازبینی‌شده/نیازمند بازبینی:", bold=True, size=11)
        for a in flags:
            _par(doc, f"• سؤال {a['q']}: {a['reason']}", size=10)


def build_word(students, exam_title, exam_total):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    for i, s in enumerate(students):
        if i:
            doc.add_page_break()
        add_student_to_doc(doc, s, exam_title, exam_total)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ───────────────────────── ذخیره/بارگذاری پروژه ─────────────────────────

def project_json():
    data = {
        "exam_title": st.session_state.exam_title,
        "exam_total": st.session_state.exam_total,
        "pages_per_student": st.session_state.pps,
        "students": st.session_state.students,
    }
    return json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")


def load_project(raw):
    data = json.loads(raw.decode("utf-8-sig"))
    st.session_state.exam_title = data.get("exam_title", "")
    st.session_state.exam_total = float(data.get("exam_total", 20.0))
    st.session_state.pps = int(data.get("pages_per_student", 2))
    studs = data.get("students", [])
    for s in studs:
        s["base"] = [tuple(b) for b in s.get("base", [])] or [(a["final"], a["note"]) for a in s["answers"]]
        s["ver"] = s.get("ver", 0) + 1
    st.session_state.students = studs


# ───────────────────────── رابط کاربری ─────────────────────────

def init_state():
    st.session_state.setdefault("students", [])
    st.session_state.setdefault("exam_title", "آزمون")
    st.session_state.setdefault("exam_total", 20.0)
    st.session_state.setdefault("pps", 2)
    st.session_state.setdefault("pdf_bytes", None)
    st.session_state.setdefault("pdf_digest", None)


CSS = """
<style>
.stApp, .stMarkdown, label, .stAlert, .stTabs, [data-testid="stSidebar"] { direction: rtl; text-align: right; }
.stTabs [data-baseweb="tab-list"] { direction: rtl; }
textarea, input { direction: rtl; text-align: right; }
</style>
"""


def sidebar():
    with st.sidebar:
        st.header("تنظیمات آزمون")
        st.text_input("عنوان آزمون", key="exam_title")
        st.number_input("بارم کل آزمون", min_value=1.0, step=0.5, key="exam_total")
        st.number_input("تعداد صفحه‌ی هر دانش‌آموز در PDF", min_value=1, max_value=20, step=1, key="pps")

        st.divider()
        st.subheader("اسکن برگه‌ها (اختیاری)")
        if fitz is None:
            st.caption("نمایش تصویر برگه فعال نیست (PyMuPDF نصب نشده).")
        else:
            up = st.file_uploader("فایل PDF اسکن همه‌ی برگه‌ها", type=["pdf"], key="pdf_up")
            if up is not None:
                raw = up.getvalue()
                st.session_state.pdf_bytes = raw
                st.session_state.pdf_digest = hashlib.md5(raw).hexdigest()
                st.caption(f"{pdf_page_count(raw)} صفحه بارگذاری شد.")
            st.caption("فایل فقط در حافظه‌ی همین نشست می‌ماند و جایی ذخیره نمی‌شود.")

        st.divider()
        st.subheader("ذخیره و بازیابی کار")
        st.caption("با بسته‌شدن صفحه یا ری‌استارت سرور، داده‌ها پاک می‌شوند. فایل ذخیره را نگه دارید.")
        if st.session_state.students:
            st.download_button("⬇️ ذخیره‌ی پروژه (JSON)", project_json(),
                               file_name="tashihyar_project.json", mime="application/json")
        proj = st.file_uploader("بازیابی پروژه‌ی ذخیره‌شده", type=["json"], key="proj_up")
        if proj is not None and st.button("بازیابی"):
            try:
                load_project(proj.getvalue())
                st.success("پروژه بازیابی شد.")
                st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"فایل پروژه معتبر نیست: {e}")


def tab_import():
    st.subheader("ورود نتایج تصحیح (JSON)")
    st.write("خروجی هر برگه را از جمینای یا Claude کپی کنید، همه‌ی متن را بچسبانید "
             "(حتی اگر قبل و بعد از JSON متن اضافه باشد)، یا فایل بارگذاری کنید. "
             "می‌توانید چند دانش‌آموز را یک‌جا بچسبانید.")
    txt = st.text_area("متن خروجی", height=220, key="paste_txt")
    files = st.file_uploader("یا فایل(های) JSON/TXT", type=["json", "txt"],
                             accept_multiple_files=True, key="json_files")
    if st.button("➕ افزودن دانش‌آموز(ان)", type="primary"):
        objs = []
        if txt.strip():
            objs += extract_json_objects(txt)
        for f in files or []:
            objs += extract_json_objects(f.getvalue().decode("utf-8-sig", errors="ignore"))
        if not objs:
            st.error("هیچ JSON معتبری (دارای کلید answers) پیدا نشد.")
        else:
            pps = st.session_state.pps
            for o in objs:
                start = len(st.session_state.students) * pps + 1
                st.session_state.students.append(normalize_student(o, start))
            st.success(f"{len(objs)} دانش‌آموز اضافه شد. به تب «بازبینی» بروید.")

    if st.session_state.students:
        st.divider()
        st.write(f"دانش‌آموزان ثبت‌شده: {len(st.session_state.students)}")


def tab_review():
    studs = st.session_state.students
    if not studs:
        st.info("ابتدا در تب «ورود نتایج» دانش‌آموزی اضافه کنید.")
        return
    names = [f"{i + 1}. {s['name']}" for i, s in enumerate(studs)]
    idx = st.selectbox("دانش‌آموز", range(len(studs)), format_func=lambda i: names[i])
    s = studs[idx]

    c1, c2, c3 = st.columns(3)
    s["name"] = c1.text_input("نام و نام خانوادگی", s["name"], key=f"nm_{s['id']}")
    s["father"] = c2.text_input("نام پدر", s["father"], key=f"fa_{s['id']}")
    s["cls"] = c3.text_input("کلاس", s["cls"], key=f"cl_{s['id']}")

    left, right = st.columns([3, 2])

    with left:
        only_flag = st.checkbox("فقط موارد نیازمند بازبینی / در انتظار", value=False, key=f"of_{s['id']}")
        rows = []
        for i, (a, b) in enumerate(zip(s["answers"], s["base"])):
            fin0, note0 = b
            if only_flag and not (a["needs_review"] or fin0 is None):
                continue
            rows.append({
                "idx": i,
                "سؤال": a["q"],
                "وضعیت": "⏳ در انتظار" if fin0 is None else ("⚠️ بازبینی" if a["needs_review"] else "✓"),
                "متن عینی دانش‌آموز": a["verbatim"] or "(سفید)",
                "تفسیر مدل": a["interpretation"],
                "اطمینان": a["confidence"],
                "پیشنهاد مدل": a["suggested"],
                "نمره نهایی": fin0,
                "بارم": a["max"],
                "دلیل مدل": a["reason"],
                "یادداشت دبیر": note0,
            })
        df = pd.DataFrame(rows)
        if df.empty:
            st.success("موردی برای نمایش نیست.")
        else:
            df = df.set_index("idx")
            edited = st.data_editor(
                df,
                key=f"ed_{s['id']}_{s['ver']}_{int(only_flag)}",
                disabled=[c for c in df.columns if c not in ("نمره نهایی", "یادداشت دبیر")],
                column_config={
                    "نمره نهایی": st.column_config.NumberColumn("نمره نهایی", min_value=0.0, step=0.05, format="%.2f"),
                    "اطمینان": st.column_config.NumberColumn("اطمینان", format="%.2f"),
                    "پیشنهاد مدل": st.column_config.NumberColumn("پیشنهاد مدل", format="%.2f"),
                    "بارم": st.column_config.NumberColumn("بارم", format="%.2f"),
                    "متن عینی دانش‌آموز": st.column_config.TextColumn("متن عینی دانش‌آموز", width="large"),
                    "تفسیر مدل": st.column_config.TextColumn("تفسیر مدل", width="medium"),
                    "دلیل مدل": st.column_config.TextColumn("دلیل مدل", width="medium"),
                },
            )
            # نوشتن ویرایش‌ها به حالت اصلی
            for i, row in edited.iterrows():
                a = s["answers"][int(i)]
                f = row["نمره نهایی"]
                f = None if pd.isna(f) else clamp(f, a["max"])
                a["final"] = f
                a["note"] = "" if pd.isna(row["یادداشت دبیر"]) else str(row["یادداشت دبیر"])

        total, pending, maxsum, flagged = totals(s)
        m1, m2, m3 = st.columns(3)
        m1.metric("جمع نمره‌ی نهایی (محاسبه با کد)", f"{total:g} از {st.session_state.exam_total:g}")
        m2.metric("در انتظار نمره‌ی دبیر", pending)
        m3.metric("علامت‌خورده برای بازبینی", flagged)
        if abs(maxsum - st.session_state.exam_total) > 0.01:
            st.warning(f"جمع بارم سؤال‌های این برگه {maxsum:g} است و با بارم کل آزمون "
                       f"({st.session_state.exam_total:g}) نمی‌خواند. ممکن است سؤالی جا افتاده باشد.")

        b1, b2 = st.columns(2)
        if b1.button("پذیرش پیشنهاد مدل برای موارد «در انتظار»", key=f"acc_{s['id']}"):
            for a in s["answers"]:
                if a["final"] is None and a["suggested"] is not None:
                    a["final"] = clamp(a["suggested"], a["max"])
            s["base"] = [(a["final"], a["note"]) for a in s["answers"]]
            s["ver"] += 1
            st.rerun()
        if b2.button("🗑 حذف این دانش‌آموز", key=f"del_{s['id']}"):
            studs.pop(idx)
            st.rerun()

    with right:
        st.markdown("**تصویر برگه**")
        if st.session_state.pdf_bytes is None or fitz is None:
            st.caption("برای دیدن تصویر برگه، PDF اسکن را در نوار کناری بارگذاری کنید.")
        else:
            s["page_start"] = st.number_input(
                "شماره‌ی صفحه‌ی اول این دانش‌آموز در PDF", min_value=1, step=1,
                value=int(s["page_start"]), key=f"ps_{s['id']}")
            pps = st.session_state.pps
            local = st.radio("صفحه", list(range(1, pps + 1)), horizontal=True,
                             format_func=lambda n: f"صفحه {n}", key=f"pg_{s['id']}")
            pdf_page = int(s["page_start"]) + int(local) - 1
            img = render_page(st.session_state.pdf_digest, pdf_page, 1.6, st.session_state.pdf_bytes)
            if img is None:
                st.warning("این صفحه در PDF وجود ندارد؛ شماره‌ی صفحه‌ی اول را اصلاح کنید.")
            else:
                cap = f"صفحه‌ی {pdf_page} از PDF"
                try:
                    st.image(img, caption=cap, width="stretch")
                except Exception:  # نسخه‌های قدیمی‌تر Streamlit
                    st.image(img, caption=cap, use_container_width=True)


def tab_export():
    studs = st.session_state.students
    if not studs:
        st.info("هنوز دانش‌آموزی ثبت نشده است.")
        return
    rows = []
    for i, s in enumerate(studs, 1):
        total, pending, _mx, flagged = totals(s)
        rows.append({"ردیف": i, "نام": s["name"], "نام پدر": s["father"], "کلاس": s["cls"],
                     "نمره نهایی": total, "در انتظار": pending, "بازبینی": flagged})
    st.dataframe(pd.DataFrame(rows), hide_index=True)
    if any(r["در انتظار"] for r in rows):
        st.warning("برخی دانش‌آموزان هنوز مورد «در انتظار» دارند؛ نمره‌ی آن‌ها کامل نیست.")

    title, total_max = st.session_state.exam_title, st.session_state.exam_total
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ جدول کلاسی و جزئیات (Excel)",
                       build_excel(studs, title, total_max),
                       file_name="tashihyar_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇️ گزارش همه‌ی دانش‌آموزان (Word)",
                       build_word(studs, title, total_max),
                       file_name="tashihyar_reports.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


def main():
    st.set_page_config(page_title="تصحیح‌یار", page_icon="📝", layout="wide")
    init_state()
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("📝 تصحیح‌یار")
    st.caption("نمره‌ها پیشنهادی‌اند؛ تصمیم نهایی با دبیر است. جمع نمره‌ها با کد محاسبه می‌شود، نه با مدل.")
    sidebar()
    t1, t2, t3 = st.tabs(["۱. ورود نتایج", "۲. بازبینی دبیر", "۳. خلاصه و خروجی"])
    with t1:
        tab_import()
    with t2:
        tab_review()
    with t3:
        tab_export()


if __name__ == "__main__":
    main()
