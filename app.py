#!/usr/bin/env python3
"""
تصحیح‌یار | نسخه وب
==================
رابط کاربری تحت وب برای تصحیح دسته‌ای اوراق امتحانی
"""

import streamlit as st
import json
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import sys

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from engine import TashihyarEngine
from batch_score import read_student_from_sheet, parse_matching, parse_table_or_structured, create_summary_excel
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Page config
st.set_page_config(
    page_title="تصحیح‌یار | سامانه تصحیح اوراق",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for RTL
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
        direction: rtl;
        text-align: right;
    }
    
    .stApp {
        direction: rtl;
    }
    
    h1, h2, h3 {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
        direction: rtl;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    
    .score-box {
        background: #E8F5E9;
        border-right: 5px solid #2E7D32;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
    
    .red-flag {
        background: #FFEBEE;
        border-right: 5px solid #C62828;
        padding: 0.8rem;
        border-radius: 8px;
        margin: 0.3rem 0;
    }
    
    div[data-testid="stMetricValue"] {
        direction: ltr;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_resource
def load_engine():
    key_path = Path(__file__).parent / "answer_key.json"
    return TashihyarEngine(str(key_path))


def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 style="color:white;margin:0;">📝 تصحیح‌یار</h1>
        <p style="color:#B0D4F1;margin:0.5rem 0 0 0;">سامانه تخصصی تصحیح اوراق امتحانی جامعه‌شناسی</p>
    </div>
    """, unsafe_allow_html=True)
    
    engine = load_engine()
    
    # Sidebar
    with st.sidebar:
        st.header("منوی اصلی")
        mode = st.radio(
            "حالت کار:",
            ["تصحیح دسته‌ای (Excel)", "تصحیح تک‌دانش‌آموز", "راهنما"],
            index=0
        )
        st.divider()
        st.caption("نسخه ۱.۱ | الگوریتم بهینه‌شده")
        st.caption("دبیرستان فرقدانی | ۱۴۰۵")
    
    if mode == "راهنما":
        show_help()
        return
    
    if mode == "تصحیح دسته‌ای (Excel)":
        batch_mode(engine)
    else:
        single_mode(engine)


def show_help():
    st.subheader("راهنمای استفاده")
    st.markdown("""
    ### تصحیح دسته‌ای (پیشنهادی)
    1. فایل قالب Excel را از بخش مربوطه دانلود کنید.
    2. برگه «لیست دانش‌آموزان» را پر کنید.
    3. برای هر دانش‌آموز برگه «نمونه_پاسخ» را کپی کرده و نام آن را به شناسه دانش‌آموز تغییر دهید.
    4. ستون زرد «پاسخ دانش‌آموز» را پر کنید.
    5. فایل را اینجا آپلود کنید و دکمه «شروع تصحیح» را بزنید.
    6. گزارش‌های Word و خلاصه نمرات را دانلود کنید.
    
    ### نکات مهم
    - برای سؤالات صحیح/غلط بنویسید: `ص` یا `غ`
    - برای تستی فقط شماره گزینه را بنویسید: `۲`
    - برای وصل‌کردنی: `۱=ت ۲=الف ۳=پ ۴=ب`
    - برای تشریحی عین نوشته دانش‌آموز را وارد کنید.
    - نمرات پیشنهادی هستند و تصمیم نهایی با دبیر است.
    """)
    
    # Download template
    template_path = Path(__file__).parent / "قالب_ورود_دسته‌ای_پاسخ_دانش‌آموزان.xlsx"
    if template_path.exists():
        with open(template_path, "rb") as f:
            st.download_button(
                "📥 دانلود قالب Excel",
                f,
                file_name="قالب_ورود_دسته‌ای_پاسخ_دانش‌آموزان.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


def batch_mode(engine):
    st.subheader("تصحیح دسته‌ای از فایل Excel")
    
    uploaded = st.file_uploader(
        "فایل Excel پاسخ دانش‌آموزان را آپلود کنید",
        type=["xlsx"],
        help="از قالب استاندارد تصحیح‌یار استفاده کنید"
    )
    
    if uploaded is None:
        st.info("ابتدا فایل Excel را آپلود کنید یا از بخش راهنما قالب را دانلود نمایید.")
        return
    
    if st.button("🚀 شروع تصحیح دسته‌ای", type="primary", use_container_width=True):
        with st.spinner("در حال تصحیح..."):
            try:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded.getvalue())
                    tmp_path = tmp.name
                
                wb = load_workbook(tmp_path, data_only=True)
                
                # Get student list
                students_meta = {}
                if "لیست دانش‌آموزان" in wb.sheetnames:
                    ws_list = wb["لیست دانش‌آموزان"]
                    for row in ws_list.iter_rows(min_row=2, max_col=4, values_only=True):
                        if row[0] and row[1]:
                            students_meta[str(row[0]).strip()] = str(row[1]).strip()
                
                reports = []
                progress = st.progress(0)
                sheets = [s for s in wb.sheetnames if s not in ["راهنما", "لیست دانش‌آموزان"]]
                
                for i, sheet_name in enumerate(sheets):
                    ws = wb[sheet_name]
                    answers = read_student_from_sheet(ws)
                    
                    sid = sheet_name
                    name = students_meta.get(sheet_name, sheet_name)
                    
                    student_data = {"id": sid, "name": name, "answers": answers}
                    report = engine.score_student(student_data)
                    reports.append(report)
                    progress.progress((i + 1) / len(sheets))
                
                st.success(f"✅ تصحیح {len(reports)} دانش‌آموز با موفقیت انجام شد.")
                
                # Summary table
                st.subheader("خلاصه نمرات")
                import pandas as pd
                data = []
                for r in reports:
                    data.append({
                        "شناسه": r.student_id,
                        "نام": r.name,
                        "تستی": round(r.test_score, 2),
                        "تشریحی": round(r.essay_score, 2),
                        "نمره کل": round(r.total_score, 2),
                        "Red Flags": len(r.red_flags)
                    })
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True, hide_index=True)
                
                # Metrics
                col1, col2, col3 = st.columns(3)
                col1.metric("میانگین کلاس", f"{df['نمره کل'].mean():.2f}")
                col2.metric("بالاترین نمره", f"{df['نمره کل'].max():.2f}")
                col3.metric("کمترین نمره", f"{df['نمره کل'].min():.2f}")
                
                # Generate downloadable files
                st.subheader("دانلود نتایج")
                
                # Create ZIP with all reports
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for r in reports:
                        # Generate Word report in memory
                        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
                            engine.generate_word_report(r, tmp_docx.name)
                            with open(tmp_docx.name, "rb") as f:
                                zf.writestr(f"گزارش_{r.student_id}_{r.name}.docx", f.read())
                    
                    # Summary Excel
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
                        create_summary_excel(reports, tmp_xlsx.name)
                        with open(tmp_xlsx.name, "rb") as f:
                            zf.writestr("خلاصه_نمرات.xlsx", f.read())
                
                zip_buffer.seek(0)
                st.download_button(
                    "📦 دانلود همه گزارش‌ها (ZIP)",
                    zip_buffer,
                    file_name="گزارش‌های_تصحیح_یار.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                # Individual Red Flags
                with st.expander("مشاهده موارد نیازمند بازبینی"):
                    for r in reports:
                        if r.red_flags:
                            st.markdown(f"**{r.name}** ({r.student_id})")
                            for flag in r.red_flags:
                                st.markdown(f'<div class="red-flag">⚠️ {flag}</div>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"خطا در پردازش: {str(e)}")
                st.exception(e)


def single_mode(engine):
    st.subheader("تصحیح تک‌دانش‌آموز")
    
    col1, col2 = st.columns(2)
    with col1:
        student_id = st.text_input("شناسه دانش‌آموز", value="S001")
    with col2:
        student_name = st.text_input("نام و نام خانوادگی", value="")
    
    st.markdown("---")
    st.markdown("#### پاسخ‌ها را وارد کنید:")
    
    answers = {}
    
    # True/False
    st.markdown("**سؤال ۱ – صحیح/غلط**")
    c1, c2, c3, c4 = st.columns(4)
    answers["1a"] = c1.selectbox("۱-الف", ["", "ص", "غ"], key="1a")
    answers["1b"] = c2.selectbox("۱-ب", ["", "ص", "غ"], key="1b")
    answers["1c"] = c3.selectbox("۱-پ", ["", "ص", "غ"], key="1c")
    answers["1d"] = c4.selectbox("۱-ت", ["", "ص", "غ"], key="1d")
    
    # Fill blanks
    st.markdown("**سؤال ۲ – جاخالی**")
    answers["2a"] = st.text_input("۲-الف (دانش ...)", key="2a")
    answers["2b"] = st.text_input("۲-ب (علوم ...)", key="2b")
    answers["2c"] = st.text_input("۲-پ", key="2c")
    answers["2d"] = st.text_input("۲-ت", key="2d")
    answers["2e"] = st.text_input("۲-ث", key="2e")
    
    # MCQ
    st.markdown("**سؤال ۳ – تستی**")
    c1, c2, c3, c4 = st.columns(4)
    answers["3a"] = c1.selectbox("۳-الف", ["", "۱", "۲", "۳", "۴"], key="3a")
    answers["3b"] = c2.selectbox("۳-ب", ["", "۱", "۲", "۳", "۴"], key="3b")
    answers["3c"] = c3.selectbox("۳-پ", ["", "۱", "۲", "۳", "۴"], key="3c")
    answers["3d"] = c4.selectbox("۳-ت", ["", "۱", "۲", "۳", "۴"], key="3d")
    
    # Matching
    st.markdown("**سؤال ۴ – وصل‌کردنی** (مثال: ۱=ت ۲=الف ۳=پ ۴=ب)")
    answers["4"] = st.text_input("پاسخ وصل", key="4", placeholder="۱=ت ۲=الف ۳=پ ۴=ب")
    
    # Essays
    st.markdown("**سؤالات تشریحی**")
    answers["5a"] = st.text_area("۵-الف", key="5a", height=80)
    answers["5b"] = st.text_area("۵-ب", key="5b", height=80)
    answers["5c"] = st.text_area("۵-پ", key="5c", height=80)
    answers["5d"] = st.text_input("۵-ت (فیزیک اجتماعی)", key="5d")
    
    answers["6a"] = st.text_input("۶-الف", key="6a")
    answers["6b"] = st.text_input("۶-ب", key="6b")
    answers["6c"] = st.text_input("۶-پ", key="6c")
    answers["6d"] = st.text_input("۶-ت", key="6d")
    
    answers["7a"] = st.text_input("۷-الف (تفسیری/تبیینی/...)", key="7a")
    answers["7b"] = st.text_area("۷-ب", key="7b", height=80)
    answers["8"] = st.text_area("۸ – دلایل علوم بومی", key="8", height=100)
    answers["9a"] = st.text_input("۹-الف (جدول: طبیعت|کنش انسان|کنش اجتماعی|جمعیت‌شناسی)", key="9a")
    answers["9b"] = st.text_area("۹-ب", key="9b", height=80)
    answers["10"] = st.text_area("۱۰ – نقش علوم اجتماعی در فناوری", key="10", height=100)
    answers["11"] = st.text_input("۱۱ (موضوع|روش|هدف)", key="11")
    answers["12"] = st.text_area("۱۲ – پیامدهای تصور جامعه همچون طبیعت", key="12", height=100)
    
    if st.button("🎯 تصحیح این دانش‌آموز", type="primary", use_container_width=True):
        # Convert answers to proper types
        processed = {}
        for k, v in answers.items():
            if v is None or str(v).strip() == "":
                processed[k] = None
                continue
            v = str(v).strip()
            if k.startswith("1"):
                processed[k] = True if v == "ص" else (False if v == "غ" else v)
            elif k.startswith("3"):
                try:
                    processed[k] = int(v.replace("۱","1").replace("۲","2").replace("۳","3").replace("۴","4"))
                except:
                    processed[k] = v
            elif k == "4":
                processed[k] = parse_matching(v)
            elif k in ["9a", "11"]:
                processed[k] = parse_table_or_structured(v)
            else:
                processed[k] = v
        
        student_data = {
            "id": student_id or "S000",
            "name": student_name or "بدون نام",
            "answers": processed
        }
        
        report = engine.score_student(student_data)
        
        # Display results
        st.markdown("---")
        st.markdown(f"### نتیجه تصحیح: {report.name}")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("نمره تستی", f"{report.test_score:.2f}")
        c2.metric("نمره تشریحی", f"{report.essay_score:.2f}")
        c3.metric("نمره کل", f"{report.total_score:.2f} / 20")
        
        if report.red_flags:
            st.markdown("#### موارد نیازمند بازبینی")
            for flag in report.red_flags:
                st.markdown(f'<div class="red-flag">⚠️ {flag}</div>', unsafe_allow_html=True)
        
        # Detailed table
        with st.expander("جزئیات نمره‌دهی هر سؤال"):
            for res in report.results:
                color = "🟢" if res.score == res.max_score else ("🟡" if res.score > 0 else "🔴")
                st.write(f"{color} **{res.question_id}**: {res.score:.2f}/{res.max_score} — {res.reason}")
        
        # Download report
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            engine.generate_word_report(report, tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button(
                    "📥 دانلود گزارش Word",
                    f,
                    file_name=f"گزارش_{report.student_id}_{report.name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )


if __name__ == "__main__":
    main()
