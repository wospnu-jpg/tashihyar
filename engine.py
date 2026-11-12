#!/usr/bin/env python3
"""
تصحیح‌یار - موتور تصحیح خودکار (نسخه ۱.۱ – الگوریتم بهینه‌شده)
=============================================================
ورودی: پاسخنامه ساخت‌یافته (JSON) + پاسخ دانش‌آموز (JSON)
خروجی: گزارش Word استاندارد + نمرات محاسبه‌شده

بهینه‌سازی‌های نسخه ۱.۱:
- تطبیق فازی مفاهیم کلیدی (token overlap)
- وزن‌دهی مفاهیم مهم‌تر
- نمره‌دهی پلکانی منصفانه‌تر برای پاسخ‌های ناقص
- دلیل شفاف‌تر برای بازبینی دبیر
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
import re


@dataclass
class ScoreResult:
    question_id: str
    student_answer: str
    score: float
    max_score: float
    reason: str
    is_red_flag: bool = False
    red_flag_note: str = ""


@dataclass
class StudentReport:
    student_id: str
    name: str
    results: List[ScoreResult] = field(default_factory=list)
    test_score: float = 0.0
    essay_score: float = 0.0
    total_score: float = 0.0
    red_flags: List[str] = field(default_factory=list)
    feedback: Dict[str, str] = field(default_factory=dict)


class TashihyarEngine:
    def __init__(self, answer_key_path: str):
        with open(answer_key_path, "r", encoding="utf-8") as f:
            self.key = json.load(f)
        self.questions = {q["id"]: q for q in self.key["questions"]}
        self.exam_info = self.key["exam_info"]

    def _normalize(self, text: str) -> str:
        """Normalize Persian text for comparison."""
        if not text:
            return ""
        text = str(text).strip()
        # Standardize Persian characters
        text = text.replace("ي", "ی").replace("ك", "ک").replace("ة", "ه")
        text = text.replace("‌", " ")  # ZWNJ to space
        # Remove extra punctuation for matching
        text = re.sub(r"[،.,;؛:!?()\[\]«»\"']", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip().lower()

    def _token_set(self, text: str) -> set:
        return set(self._normalize(text).split())

    def _concept_match_score(self, text: str, concept: str) -> float:
        """
        Return 1.0 for strong match, 0.5–0.8 for partial/fuzzy, 0.0 for none.
        Handles common Persian morphological variations (هدفدار ↔ هدف، معنادار ↔ معنا).
        """
        t = self._normalize(text)
        c = self._normalize(concept)
        if not c:
            return 0.0
        if c in t:
            return 1.0
        
        # Stem-like suffixes common in Persian adjectives/nouns
        stems = {c}
        for suf in ["دار", "انه", "ی", "گر", "مند", "انه"]:
            if c.endswith(suf) and len(c) > len(suf) + 1:
                stems.add(c[:-len(suf)])
        
        for stem in stems:
            if len(stem) >= 3 and stem in t:
                return 0.85  # strong partial
        
        # Token overlap
        t_tokens = self._token_set(text)
        c_tokens = self._token_set(concept)
        if not c_tokens:
            return 0.0
        overlap = len(t_tokens & c_tokens) / len(c_tokens)
        if overlap >= 0.8:
            return 1.0
        if overlap >= 0.5:
            return 0.6
        # Single important token present
        for tok in c_tokens:
            if len(tok) >= 3 and tok in t_tokens:
                return 0.45
        return 0.0

    def _score_concepts(self, text: str, concepts: List[str], weights: Optional[List[float]] = None) -> Tuple[float, int, List[str]]:
        """
        Score based on weighted key concepts.
        Returns: (normalized_score 0-1, matched_count, list of matched concepts)
        """
        if not concepts:
            return 0.0, 0, []
        if weights is None or len(weights) != len(concepts):
            weights = [1.0] * len(concepts)
        
        total_weight = sum(weights)
        earned = 0.0
        matched = []
        for concept, w in zip(concepts, weights):
            m = self._concept_match_score(text, concept)
            if m > 0:
                earned += m * w
                matched.append(concept)
        return (earned / total_weight if total_weight else 0.0), len(matched), matched

    def score_true_false(self, q: Dict, answer: Any) -> ScoreResult:
        max_s = q["max_score"]
        if answer is None or str(answer).strip() == "":
            return ScoreResult(q["id"], "[سفید]", 0.0, max_s, "سفید", True, "پاسخ سفید")
        
        # Accept True/False, ص/غ, ✓/✗, 1/0, درست/نادرست
        val = str(answer).strip().lower()
        is_true = val in ["true", "1", "ص", "درست", "✓", "√", "true", "تیک"]
        is_false = val in ["false", "0", "غ", "نادرست", "✗", "×", "ضربدر"]
        
        if not (is_true or is_false):
            return ScoreResult(q["id"], str(answer), 0.0, max_s, "علامت نامشخص", True, "علامت قابل تشخیص نیست")
        
        student_bool = is_true
        correct = q["correct"]
        if student_bool == correct:
            return ScoreResult(q["id"], "درست" if student_bool else "نادرست", max_s, max_s, "صحیح")
        else:
            return ScoreResult(q["id"], "درست" if student_bool else "نادرست", 0.0, max_s, "نادرست")

    def score_fill_or_short(self, q: Dict, answer: Any) -> ScoreResult:
        max_s = q["max_score"]
        if answer is None or str(answer).strip() in ["", "[سفید]", "سفید"]:
            return ScoreResult(q["id"], "[سفید]", 0.0, max_s, "سفید", True, "پاسخ سفید")
        
        ans_raw = str(answer).strip()
        acceptable = q.get("acceptable", q.get("correct_answers", []))
        
        best = 0.0
        for acc in acceptable:
            m = self._concept_match_score(ans_raw, acc)
            if m > best:
                best = m
        
        if best >= 0.9:
            return ScoreResult(q["id"], ans_raw, max_s, max_s, "صحیح")
        elif best >= 0.5:
            score = round(max_s * 0.5, 2)
            return ScoreResult(q["id"], ans_raw, score, max_s, "نزدیک به پاسخ صحیح (نیم‌نمره)", True, "نیاز به تأیید دبیر")
        else:
            return ScoreResult(q["id"], ans_raw, 0.0, max_s, "نادرست")

    def score_multiple_choice(self, q: Dict, answer: Any) -> ScoreResult:
        max_s = q["max_score"]
        if answer is None or str(answer).strip() == "":
            return ScoreResult(q["id"], "[سفید]", 0.0, max_s, "سفید", True, "پاسخ سفید")
        
        try:
            choice = int(str(answer).strip())
        except:
            return ScoreResult(q["id"], str(answer), 0.0, max_s, "گزینه نامعتبر", True, "گزینه قابل تشخیص نیست")
        
        if choice == q["correct"]:
            return ScoreResult(q["id"], f"گزینه {choice}", max_s, max_s, "صحیح")
        else:
            return ScoreResult(q["id"], f"گزینه {choice}", 0.0, max_s, "نادرست")

    def score_matching(self, q: Dict, answer: Dict) -> ScoreResult:
        max_s = q["max_score"]
        item_s = q["item_score"]
        if not answer or not isinstance(answer, dict):
            return ScoreResult(q["id"], "[سفید یا ناقص]", 0.0, max_s, "پاسخ ناقص", True, "وصل ناقص")
        
        correct = q["correct_mapping"]
        score = 0.0
        details = []
        for k, v in correct.items():
            student_v = str(answer.get(k, "")).strip()
            if student_v == v:
                score += item_s
                details.append(f"{k}→{v}✓")
            else:
                details.append(f"{k}→{student_v or 'سفید'}✗")
        
        reason = " | ".join(details)
        is_flag = score < max_s
        return ScoreResult(q["id"], reason, score, max_s, f"{score}/{max_s}", is_flag, "برخی موارد نادرست" if is_flag else "")

    def score_essay(self, q: Dict, answer: Any) -> ScoreResult:
        """
        Optimized essay scoring:
        - Weighted key concepts
        - Partial credit based on coverage ratio
        - Full score when coverage >= threshold
        - Transparent reason for teacher review
        """
        max_s = q["max_score"]
        if answer is None or str(answer).strip() in ["", "[سفید]", "سفید"]:
            return ScoreResult(q["id"], "[سفید]", 0.0, max_s, "سفید", True, "پاسخ سفید")
        
        ans = str(answer).strip()
        concepts = q.get("key_concepts", [])
        weights = q.get("concept_weights")  # optional list of weights
        min_full = q.get("min_concepts_for_full", max(2, len(concepts) // 2 + 1))
        full_threshold = q.get("full_score_threshold", 0.75)  # 75% coverage → full mark
        
        coverage, matched_count, matched_list = self._score_concepts(ans, concepts, weights)
        
        # Scoring bands (optimized for fairness)
        if coverage >= full_threshold and matched_count >= min_full:
            score = max_s
            reason = f"کامل – پوشش مفاهیم {coverage:.0%} ({matched_count}/{len(concepts)})"
            is_flag = False
            note = ""
        elif coverage >= 0.50:
            # Strong partial – scale from 60% to 95% of max
            ratio = 0.60 + (coverage - 0.50) * 0.7
            score = round(max_s * min(ratio, 0.95), 2)
            reason = f"خوب اما ناقص – پوشش {coverage:.0%} | مفاهیم: {', '.join(matched_list[:4])}"
            is_flag = True
            note = "پیشنهاد بازبینی دبیر برای ارفاق احتمالی"
        elif coverage >= 0.28:
            score = round(max_s * (0.35 + coverage * 0.5), 2)
            reason = f"ناقص – پوشش {coverage:.0%} | مفاهیم یافت‌شده: {', '.join(matched_list) or '—'}"
            is_flag = True
            note = "نیاز به بازبینی انسانی"
        elif coverage > 0:
            score = round(max_s * 0.20, 2)
            reason = f"اشاره محدود – پوشش {coverage:.0%}"
            is_flag = True
            note = "نیاز به بازبینی انسانی"
        else:
            score = 0.0
            reason = "عدم اشاره به مفاهیم کلیدی پاسخنامه"
            is_flag = True
            note = "نیاز به بازبینی انسانی"
        
        display_ans = ans[:130] + ("..." if len(ans) > 130 else "")
        return ScoreResult(q["id"], display_ans, score, max_s, reason, is_flag, note)

    def score_table(self, q: Dict, answer) -> ScoreResult:
        max_s = q["max_score"]
        if not answer:
            return ScoreResult(q["id"], "[سفید]", 0.0, max_s, "سفید", True, "جدول خالی")
        
        # Accept both dict and raw string
        if isinstance(answer, str):
            parts = [p.strip() for p in answer.replace("/", "|").split("|")]
            keys = list(q["parts"].keys())
            answer = dict(zip(keys, parts)) if parts else {}
        
        if not isinstance(answer, dict):
            return ScoreResult(q["id"], str(answer), 0.0, max_s, "فرمت نامعتبر", True, "نیاز به بازبینی")
        
        score = 0.0
        details = []
        for part_name, part_info in q["parts"].items():
            student_val = self._normalize(str(answer.get(part_name, "")))
            correct_list = [self._normalize(c) for c in part_info["correct"]]
            part_score = part_info["score"]
            matched = any(c in student_val or student_val in c for c in correct_list if student_val)
            if matched:
                score += part_score
                details.append(f"{part_name}✓")
            else:
                details.append(f"{part_name}✗")
        
        return ScoreResult(q["id"], " | ".join(details), score, max_s, f"{score}/{max_s}")

    def score_structured_short(self, q: Dict, answer) -> ScoreResult:
        return self.score_table(q, answer)  # same logic

    def score_student(self, student_data: Dict) -> StudentReport:
        report = StudentReport(
            student_id=student_data.get("id", "unknown"),
            name=student_data.get("name", "نامشخص")
        )
        
        answers = student_data.get("answers", {})
        
        for qid, q in self.questions.items():
            ans = answers.get(qid)
            qtype = q["type"]
            
            if qtype == "true_false":
                res = self.score_true_false(q, ans)
            elif qtype in ["fill", "short_answer"]:
                res = self.score_fill_or_short(q, ans)
            elif qtype == "multiple_choice":
                res = self.score_multiple_choice(q, ans)
            elif qtype == "matching":
                res = self.score_matching(q, ans or {})
            elif qtype in ["short_essay", "essay"]:
                res = self.score_essay(q, ans)
            elif qtype in ["table", "structured_short"]:
                res = self.score_table(q, ans or {})
            else:
                res = ScoreResult(qid, str(ans), 0.0, q["max_score"], "نوع پشتیبانی نشده")
            
            report.results.append(res)
            if res.is_red_flag:
                report.red_flags.append(f"{qid}: {res.red_flag_note or res.reason}")
        
        # Calculate totals
        test_ids = ["3a", "3b", "3c", "3d"]
        report.test_score = sum(r.score for r in report.results if r.question_id in test_ids)
        report.total_score = sum(r.score for r in report.results)
        report.essay_score = report.total_score - report.test_score
        
        # Simple feedback
        if report.total_score >= 16:
            strength = "عملکرد بسیار خوب در اکثر بخش‌ها"
        elif report.total_score >= 12:
            strength = "عملکرد قابل قبول با نقاط قوت در تستی"
        else:
            strength = "نیاز به تقویت در پاسخ‌های تشریحی"
        
        report.feedback = {
            "strength": strength,
            "weakness": "؛ ".join(report.red_flags[:3]) if report.red_flags else "مورد خاصی ثبت نشد",
            "suggestion": "تمرین پاسخ‌گویی کامل به سؤالات تشریحی و توجه به مفاهیم کلیدی پاسخنامه"
        }
        
        return report

    def generate_word_report(self, report: StudentReport, output_path: str):
        doc = Document()
        
        # RTL setup
        section = doc.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.right_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        
        def add_rtl_para(text, bold=False, size=12, color=None, center=False):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(text)
            run.font.size = Pt(size)
            run.font.name = "Tahoma"
            run.bold = bold
            run.font.rtl = True
            if color:
                run.font.color.rgb = RGBColor(*color)
            # Set RTL
            pPr = p._p.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
            return p
        
        # Header
        add_rtl_para("تصحیح‌یار | گزارش تصحیح خودکار", bold=True, size=16, center=True)
        add_rtl_para(f"{self.exam_info['title']} – {self.exam_info['date']}", size=11, center=True)
        add_rtl_para(f"دانش‌آموز: {report.name}  |  شناسه: {report.student_id}", bold=True, size=12)
        add_rtl_para(f"نمره تستی: {report.test_score:.2f}   |   نمره تشریحی: {report.essay_score:.2f}   |   نمره کل پیشنهادی: {report.total_score:.2f} از {self.exam_info['total_score']}", 
                     bold=True, size=11, color=(0, 100, 0))
        
        doc.add_paragraph()
        
        # Results table
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["شماره", "پاسخ دانش‌آموز", "نمره", "بارم", "دلیل"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            for p in hdr[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(9)
                    r.font.name = "Tahoma"
        
        for res in report.results:
            row = table.add_row().cells
            row[0].text = res.question_id
            row[1].text = res.student_answer[:80]
            row[2].text = f"{res.score:.2f}"
            row[3].text = f"{res.max_score:.2f}"
            row[4].text = res.reason
            for c in row:
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    for r in p.runs:
                        r.font.size = Pt(8)
                        r.font.name = "Tahoma"
        
        doc.add_paragraph()
        add_rtl_para("موارد نیازمند بازبینی (Red Flags):", bold=True, size=11, color=(180, 0, 0))
        if report.red_flags:
            for rf in report.red_flags:
                add_rtl_para(f"• {rf}", size=10)
        else:
            add_rtl_para("مورد خاصی ثبت نشد.", size=10)
        
        doc.add_paragraph()
        add_rtl_para("بازخورد آموزشی:", bold=True, size=11)
        add_rtl_para(f"نقاط قوت: {report.feedback.get('strength', '')}", size=10)
        add_rtl_para(f"مهم‌ترین اشکال: {report.feedback.get('weakness', '')}", size=10)
        add_rtl_para(f"پیشنهاد: {report.feedback.get('suggestion', '')}", size=10)
        
        add_rtl_para("─" * 40, size=9, center=True)
        add_rtl_para("نمره ارائه‌شده صرفاً پیشنهادی است. تصمیم نهایی با دبیر است.", size=9, color=(100, 100, 100), center=True)
        
        doc.save(output_path)
        return output_path


def main_demo():
    """دمو با یک دانش‌آموز نمونه"""
    engine = TashihyarEngine("/home/workdir/artifacts/tashihyar_engine/answer_key.json")
    
    # نمونه پاسخ دانش‌آموز (نگارگهی - نمره بالا)
    sample_student = {
        "id": "S003",
        "name": "نگارگهی",
        "answers": {
            "1a": True, "1b": False, "1c": True, "1d": False,
            "2a": "دانش عمومی", "2b": "ابزاری", "2c": "کلان", "2d": "موجود زنده", "2e": "نظم اجتماعی",
            "3a": 2, "3b": 2, "3c": 3, "3d": 1,
            "4": {"1": "ت", "2": "الف", "3": "پ", "4": "ب"},
            "5a": "انسان در کنش خود هدف و معنایی دارد و کنش بدون فهم هدف و معنا ممکن نیست",
            "5b": "چون فردی که به تنهایی و بیرون از جامعه زندگی می‌کند زندگی‌اش پرزحمت می‌شود چون باید خودش به تنهایی راه و روش زندگی را یاد بگیرد",
            "5c": "چون جامعه‌شناسان شباهت جامعه با ماشین و بدن انسان را برجسته می‌کنند تا بتوانند پیش‌بینی و کنترل کنند",
            "5d": "فیزیک اجتماعی",
            "6a": "دانش علمی", "6b": "علوم تفسیری", "6c": "", "6d": "آشنایی‌زدایی",
            "7a": "تفسیری",
            "7b": "دانش عمومی پایه و گسترده‌تر است و دانش علمی جزئی از آن است و برتری دانش تجربی پذیرفته نمی‌شود",
            "8": "علوم غربی با فرهنگ جامعه ما تعارض دارد و این علوم به تنهایی قادر به حل مسائل جامعه ما نیستند و نیاز به علوم بومی مطابق فرهنگ اسلامی و ایرانی است",
            "9a": {"natural": "طبیعت", "human": "کنش انسان", "social": "کنش اجتماعی", "example": "جمعیت‌شناسی"},
            "9b": "چون کنش‌های انسان آگاهانه و ارادی است و انسان می‌تواند رفتار خود را تغییر دهد برخلاف پدیده‌های طبیعی که قوانین ثابت دارند",
            "10": "علوم اجتماعی می‌توانند با شناخت کنش‌ها پیامدهای فناوری را بررسی و پیش‌بینی کنند و با نقد و داوری جهت درست استفاده را نشان دهند تا انسان بر فناوری مسلط باشد",
            "11": {"subject": "پدیده‌های اجتماعی مانند پدیده‌های طبیعی", "method": "حس و تجربه نگاه از بیرون", "goal": "پیش‌بینی و کنترل"},
            "12": "فقط از دانش ابزاری استفاده می‌کنیم و نمی‌توانیم تغییر بنیادی ایجاد کنیم. نظم اجتماعی مستقل از ماست. انسان‌ها به پیچ و مهره تبدیل می‌شوند"
        }
    }
    
    report = engine.score_student(sample_student)
    print(f"دانش‌آموز: {report.name}")
    print(f"نمره تستی: {report.test_score:.2f}")
    print(f"نمره تشریحی: {report.essay_score:.2f}")
    print(f"نمره کل: {report.total_score:.2f}")
    print(f"Red Flags: {len(report.red_flags)}")
    
    out_path = "/home/workdir/artifacts/tashihyar_engine/sample_report_negar.docx"
    engine.generate_word_report(report, out_path)
    print(f"گزارش ذخیره شد: {out_path}")
    return report


if __name__ == "__main__":
    main_demo()
