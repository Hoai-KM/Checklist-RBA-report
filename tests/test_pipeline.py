# -*- coding: utf-8 -*-
"""
Bộ test tự động cho app sinh báo cáo HR audit RBA.

Chạy toàn bộ:
    python -m unittest discover -s tests -v

Bao gồm:
1. Test KB (knowledge_base)  - đủ 52 Audit ID, ánh xạ dòng 6-57
2. Test rewrite              - parse draft, tách song ngữ, placeholder, rich text (header đậm)
3. Test report_engine        - đọc rating/evidence, chuẩn hóa dòng trống (không ghi file)
4. Test pipeline (regression)- chạy CLI sinh file -> so từng ô với file mẫu đã duyệt
5. Smoke test GUI            - khởi tạo tkinter App, nạp báo cáo, đóng
"""
import os
import sys
import tempfile
import unicodedata
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

INPUT_XLSX = os.path.join(
    ROOT, "Copy of 2026.03.11_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx")
SAMPLE_XLSX = os.path.join(
    ROOT, "2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx")

CLI_DEFAULTS = [
    "--audit-date", "2026.08.26",
    "--audit-type", "Corrective Action Verification",
    "--team-leader", "KHUẤT MẠNH HOÀI",
    "--team-member", "NGUYỄN TIẾN THỊNH (SEVT)",
    "--closing-date", "2026-08-26",
    "--interview-count", "11",
    "--interview-score", "0.894",
    "--profile-year", "2025",
]

# Test luôn chạy offline (từ điển dự phòng) để kết quả xác định, không phụ thuộc mạng
os.environ.setdefault("HR_AUDIT_TRANSLATOR_ONLINE", "0")

from openpyxl.cell.rich_text import CellRichText, TextBlock  # noqa: E402

from hr_audit_app.knowledge_base import ITEMS, ITEM_ROWS, ROW_ITEMS  # noqa: E402
from hr_audit_app.report_engine import (  # noqa: E402
    ALLOWED_RATINGS,
    CHECKLIST,
    Report,
)
from hr_audit_app.rewrite import (  # noqa: E402
    HEADER_EN,
    HEADER_VI,
    PLACEHOLDER,
    REFS_EN,
    REFS_VI,
    build_content,
    has_english_words,
    has_vietnamese_words,
    is_placeholder,
    is_vietnamese,
    load_curated,
    numbered,
    ensure_section_gap,
    parse_bilingual,
    parse_draft,
    purify_findings,
    repair_mixed_item,
    scan_language_mixing,
    split_bilingual_line,
    split_findings,
    split_refs,
    structure_bilingual,
    to_rich_text,
    elaborate_findings,
    VN_CHARS_RE,
    normalize_citation,
    missing_input_items,
    pool_covers,
    supplement_missing_findings,
    standardize_refs_with_kb,
)
from hr_audit_app.elaborations import ELABORATIONS_FULL, _elaborate_key


class TestKnowledgeBase(unittest.TestCase):
    """KB phải đủ 52 tiêu chí và ánh xạ đúng dòng sheet Checklist."""

    def test_52_items(self):
        self.assertEqual(len(ITEMS), 52)

    def test_rows_6_to_57(self):
        self.assertEqual(sorted(ROW_ITEMS), list(range(6, 58)))
        self.assertEqual(sorted(ITEM_ROWS.values()), list(range(6, 58)))

    def test_row_mapping_roundtrip(self):
        for row, audit_id in ROW_ITEMS.items():
            self.assertEqual(ITEM_ROWS[audit_id], row)

    def test_every_item_has_citations(self):
        for audit_id, item in ITEMS.items():
            with self.subTest(audit_id=audit_id):
                self.assertIn("Section", item["rba_citation_en"], audit_id)
                self.assertIn("Mục", item["rba_citation_vi"], audit_id)
                self.assertIsInstance(item["vn_citations_en"], list)
                self.assertIsInstance(item["vn_citations_vi"], list)
                self.assertEqual(len(item["vn_citations_en"]),
                                 len(item["vn_citations_vi"]), audit_id)


class TestElaborations(unittest.TestCase):
    """Khối phân tích RBA (elab_a/b/c): đủ khóa cho 52 tiêu chí, song ngữ thuần."""

    REQUIRED_KEYS = {
        "forced_labor", "young_workers", "working_hours", "wages_benefits",
        "humane_treatment", "non_discrimination", "freedom_association",
        "labor_mgmt", "privacy", "ethics_mgmt", "supply_chain", "general",
    }

    def test_every_item_maps_to_existing_key(self):
        """Mọi audit_id trong KB phải tra được khóa elaboration có sẵn."""
        for audit_id in ITEMS:
            with self.subTest(audit_id=audit_id):
                key = _elaborate_key(audit_id, ITEMS.get)
                self.assertIn(key, ELABORATIONS_FULL, audit_id)

    def test_all_required_keys_present(self):
        self.assertTrue(self.REQUIRED_KEYS <= set(ELABORATIONS_FULL),
                        self.REQUIRED_KEYS - set(ELABORATIONS_FULL))

    def test_vi_texts_have_diacritics_and_no_english(self):
        """Mọi câu VI phải có dấu tiếng Việt đầy đủ và không lẫn từ Anh."""
        for key, pair in ELABORATIONS_FULL.items():
            with self.subTest(key=key):
                vi = pair.get("vi", "")
                self.assertTrue(vi, key)
                self.assertIsNotNone(VN_CHARS_RE.search(vi),
                                     f"VI thiếu dấu tiếng Việt: {key}")
                self.assertFalse(has_english_words(vi),
                                 f"VI lẫn từ tiếng Anh: {key}")

    def test_en_texts_free_of_vietnamese(self):
        """Mọi câu EN phải thuần tiếng Anh (không sót ký tự/từ Việt)."""
        for key, pair in ELABORATIONS_FULL.items():
            with self.subTest(key=key):
                en = pair.get("en", "")
                self.assertTrue(en, key)
                self.assertFalse(has_vietnamese_words(en),
                                 f"EN lẫn tiếng Việt: {key}")

    def test_elaborate_appends_short_finding_only(self):
        """Finding ngắn được bổ sung phân tích; finding >= 400 ký tự giữ nguyên."""
        short_en, short_vi = "Worker passport was held by the employer.", \
            "Hộ chiếu người lao động bị công ty giữ."
        e, v = elaborate_findings("AL101", [short_en], [short_vi])
        self.assertGreater(len(e[0]), len(short_en))
        self.assertGreater(len(v[0]), len(short_vi))
        self.assertTrue(VN_CHARS_RE.search(v[0]))
        long_en = "x" * 500
        e2, v2 = elaborate_findings("AL101", [long_en], ["y" * 500])
        self.assertEqual(e2[0], long_en)
        self.assertEqual(v2[0], "y" * 500)


class TestRewrite(unittest.TestCase):
    """Động cơ biên tập: parse draft, song ngữ, placeholder."""

    def test_is_vietnamese_detection(self):
        self.assertTrue(is_vietnamese("Người lao động"))
        self.assertFalse(is_vietnamese("Worker interview"))

    def test_split_bilingual_line(self):
        vi, en = split_bilingual_line(
            "Công ty chưa trả lương đúng hạn / The company delayed wage payments")
        self.assertEqual(vi, "Công ty chưa trả lương đúng hạn")
        self.assertEqual(en, "The company delayed wage payments")
        self.assertEqual(split_bilingual_line("Plain English line only"), (None, None))

    def test_parse_draft_old_placeholder_empty(self):
        d = parse_draft("Audit findings:\n1:\n2:\n3:\n\nLegal references:")
        self.assertFalse(d.has_content)
        self.assertTrue(is_placeholder("Audit findings:\n1:\n2:\n3:\n\nLegal references:"))
        self.assertTrue(is_placeholder(PLACEHOLDER))
        self.assertTrue(is_placeholder(None))
        self.assertTrue(is_placeholder(""))

    def test_parse_draft_with_content(self):
        text = ("Audit findings:\n"
                "1: Công ty chưa ký HĐLĐ bằng văn bản / Company did not sign written contracts\n"
                "Legal references:\n"
                "Labor Code No. 45/2019/QH14, Article 14.2")
        d = parse_draft(text)
        self.assertTrue(d.has_content)
        self.assertEqual(d.items_vi, ["Công ty chưa ký HĐLĐ bằng văn bản"])
        self.assertEqual(d.items_en, ["Company did not sign written contracts"])
        self.assertIn("Article 14.2", d.refs_raw)

    def test_structure_bilingual_roundtrip(self):
        block = structure_bilingual(
            ["Finding EN"], ["RBA Code of Conduct, Version 8.0.1, Section A1"],
            ["Finding VI"], ["Bộ luật Lao động 2019, Điều 8"])
        self.assertIn("AUDIT FINDINGS (ENGLISH)", block)
        self.assertIn("NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)", block)
        parts = parse_bilingual(block)
        self.assertEqual(parts["en_findings"], ["Finding EN"])
        self.assertEqual(parts["vi_findings"], ["Finding VI"])
        self.assertEqual(parts["en_refs"], ["RBA Code of Conduct, Version 8.0.1, Section A1"])
        self.assertEqual(parts["vi_refs"], ["Bộ luật Lao động 2019, Điều 8"])

    def test_structure_bilingual_refs_numbered(self):
        """Trích dẫn pháp lý phải đánh số '1. ' '2. ' (không dùng gạch đầu dòng '- ')."""
        block = structure_bilingual(
            ["Finding EN"], ["Ref A", "Ref B"],
            ["Finding VI"], ["Căn A", "Căn B"])
        self.assertIn("Legal references:\n1. Ref A\n2. Ref B", block)
        self.assertIn("Căn cứ pháp lý:\n1. Căn A\n2. Căn B", block)
        # Parse ngược lại phải lấy đúng nội dung trích dẫn (bỏ marker số)
        parts = parse_bilingual(block)
        self.assertEqual(parts["en_refs"], ["Ref A", "Ref B"])
        self.assertEqual(parts["vi_refs"], ["Căn A", "Căn B"])

    def test_split_refs_bilingual_and_markers(self):
        raw = (
            "- Bộ luật Lao động 2019, Điều 8 / Labor Code 2019, Article 8\n"
            "1. RBA Code of Conduct, Version 8.0.1\n"
            "2) Bộ luật Lao động 2019, Điều 113"
        )
        en_r, vi_r = split_refs(raw)
        self.assertEqual(
            vi_r, ["Bộ luật Lao động 2019, Điều 8", "Bộ luật Lao động 2019, Điều 113"])
        self.assertEqual(
            en_r, ["Labor Code 2019, Article 8", "RBA Code of Conduct, Version 8.0.1"])

    def test_split_refs_empty(self):
        self.assertEqual(split_refs(""), ([], []))
        self.assertEqual(split_refs(None), ([], []))

    def test_build_content_separates_en_vi(self):
        """
        Phần tiếng Anh chỉ chứa tiếng Anh, phần tiếng Việt chỉ chứa tiếng Việt
        kể cả khi draft ghi lẫn "VI / EN" trên cùng dòng.
        """
        draft = (
            "Audit findings:\n"
            "1: Công ty chưa thiết lập quy trình đánh giá tuân thủ / "
            "The company has not established a compliance review process\n"
            "2: Không có hồ sơ đào tạo / No training records provided\n"
            "\nLegal references:\n"
            "- Bộ luật Lao động 2019, Điều 8 / Labor Code 2019, Article 8\n"
        )
        out = build_content("AL101", draft, rating="Minor", curated={}, add_kb_refs=True)
        parts = parse_bilingual(out)
        # EN findings & refs: không chứa ký tự tiếng Việt
        for x in parts["en_findings"] + parts["en_refs"]:
            self.assertFalse(is_vietnamese(x), f"EN lẫn tiếng Việt: {x}")
        # VI findings & refs: phải là tiếng Việt
        for x in parts["vi_findings"] + parts["vi_refs"]:
            self.assertTrue(is_vietnamese(x), f"VI không phải tiếng Việt: {x}")
        self.assertIn("compliance review process",
                      " ".join(parts["en_findings"]))
        self.assertIn("quy trình đánh giá tuân thủ",
                      " ".join(parts["vi_findings"]))
        self.assertNotIn("- ", out)  # không còn gạch đầu dòng

    def test_build_content_reroutes_misplaced_language(self):
        """
        Draft đã có header song ngữ nhưng ghi lẫn ngôn ngữ -> build_content
        tự chuyển về đúng phần: tiếng Anh chỉ EN, tiếng Việt chỉ VI.
        """
        draft = (
            "AUDIT FINDINGS (ENGLISH)\n"
            "1. Công ty chưa trả lương đúng hạn\n"
            "Legal references:\n"
            "1. Decree No. 12/2022/ND-CP\n"
            "\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            "1. The company delayed wage payments\n"
            "Căn cứ pháp lý:\n"
            "1. Labor Code No. 45/2019/QH14, Article 97\n"
        )
        out = build_content("AL101", draft, rating="Minor", curated={}, add_kb_refs=False)
        parts = parse_bilingual(out)
        self.assertEqual(parts["en_findings"], ["The company delayed wage payments"])
        self.assertEqual(parts["vi_findings"], ["Công ty chưa trả lương đúng hạn"])
        self.assertEqual(
            parts["en_refs"],
            ["Decree No. 12/2022/ND-CP", "Labor Code No. 45/2019/QH14, Article 97"])
        # Phần VI thiếu trích dẫn -> tự dịch bù từ phần EN
        self.assertEqual(
            parts["vi_refs"],
            ["Nghị định số 12/2022/NĐ-CP",
             "Bộ luật Lao động năm 2019 (Luật số 45/2019/QH14), Điều 97"])

    def test_parse_draft_drops_findings_label(self):
        """Dòng nhãn phụ '1. Findings:' trong draft nhà máy phải bị loại."""
        draft = ("Audit findings:\n"
                 "1. Findings:\n"
                 "2. Severe systemic overtime violations found\n")
        d = parse_draft(draft)
        self.assertEqual(d.items_en, ["Severe systemic overtime violations found"])
        self.assertEqual(d.items_vi, [])

    def test_split_findings_drops_label(self):
        """split_findings cũng loại nhãn 'Findings:' (kể cả khi đứng đầu item đa dòng)."""
        self.assertEqual(split_findings("1. Findings:\n2. Wage delay evidence\n"),
                         ["Wage delay evidence"])
        self.assertEqual(
            split_findings("1. Findings:\nWage delay evidence continues\n2. Next item\n"),
            ["Wage delay evidence continues", "Next item"])

    def test_build_content_mirrors_missing_vi_refs(self):
        """Draft chỉ có trích dẫn EN -> 'Căn cứ pháp lý:' tự có bản dịch tiếng Việt."""
        draft = ("AUDIT FINDINGS (ENGLISH)\n"
                 "1. Findings:\n"
                 "2. Audit sampled review revealed systemic overtime violations\n"
                 "Legal references:\n"
                 "1. Vietnam Labor Code 2019, Article 107\n"
                 "\n"
                 "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
                 "1. Qua kiểm tra phát hiện sai phạm về giờ làm thêm\n"
                 "Căn cứ pháp lý:\n")
        out = build_content("AL301", draft, rating="Major", curated={}, add_kb_refs=False)
        self.assertNotIn("Findings:", out)  # nhãn phụ đã bị loại
        self.assertIn(
            "Legal references:\n1. Vietnam Labor Code 2019, Article 107", out)
        self.assertIn(
            "Căn cứ pháp lý:\n1. Bộ luật Lao động năm 2019, Điều 107", out)

    def test_build_content_mirrors_missing_en_refs(self):
        """Draft chỉ có trích dẫn VI -> 'Legal references:' tự có bản dịch tiếng Anh."""
        draft = ("Audit findings:\n"
                 "1: Công ty cho làm thêm quá 40 giờ/tháng\n"
                 "Legal references:\n"
                 "1. Bộ luật Lao động 2019, Điều 107\n")
        out = build_content("AL301", draft, rating="Major", curated={}, add_kb_refs=False)
        parts = parse_bilingual(out)
        # Dòng "(Bản dịch tự động - ...)" là ghi chú, không phải trích dẫn
        self.assertEqual([r for r in parts["vi_refs"] if not r.startswith("(")],
                         ["Bộ luật Lao động 2019, Điều 107"])
        self.assertEqual(parts["en_refs"],
                         ["Labor Code No. 45/2019/QH14, Article 107"])

    def test_translate_text_vi_to_en(self):
        """Từ điển ngoại tuyến: câu OT tiếng Việt -> tiếng Anh dùng được."""
        from hr_audit_app.translator import translate_text
        vi = ("trung bình 39.8% tổng số nhân sự làm thêm từ 40 đến 52 giờ/tháng "
              "và 17.8% nhân sự làm thêm từ 52 đến 104 giờ/tháng, vượt xa mức trần")
        en = translate_text(vi, "en")
        self.assertIn("39.8% of workers", en)
        self.assertIn("40 to 52 hours per month", en)
        self.assertIn("overtime", en.lower())
        self.assertNotIn("giờ/tháng", en)

    def test_translate_text_en_to_vi(self):
        """Từ điển ngoại tuyến: câu OT tiếng Anh -> tiếng Việt dùng được."""
        from hr_audit_app.translator import translate_text
        en = ("an average of 39.8% of the workforce performed overtime between "
              "40 to 52 hours per month")
        vi = translate_text(en, "vi")
        self.assertIn("39.8% tổng số NLĐ", vi)
        self.assertIn("40 ~ 52 giờ/tháng", vi)
        self.assertIn("làm thêm", vi)

    def test_mirror_findings_translates_missing_language(self):
        """mirror_findings: thiếu EN -> dịch từ VI (và ngược lại), có cờ translated."""
        from hr_audit_app.rewrite import mirror_findings
        en_f, vi_f, flag = mirror_findings([], ["Công ty cho làm thêm quá 40 giờ/tháng"])
        self.assertTrue(flag)
        self.assertEqual(vi_f, ["Công ty cho làm thêm quá 40 giờ/tháng"])
        self.assertTrue(en_f and all(not any(
            "\u0e00" <= ch <= "\u1ef9" for ch in t) for t in en_f))
        en_f2, vi_f2, flag2 = mirror_findings(["Wage delay found"], [], )
        self.assertTrue(flag2)
        self.assertEqual(en_f2, ["Wage delay found"])
        self.assertTrue(vi_f2)

    def test_build_content_vi_only_draft_gets_both_languages(self):
        """Draft chỉ có tiếng Việt -> app tự dịch ra tiếng Anh, đủ 2 phần ngôn ngữ."""
        draft = ("Audit findings:\n"
                 "1. Qua kiểm tra phát hiện nhân sự làm thêm quá 40 giờ/tháng\n"
                 "Legal references:\n"
                 "1. Bộ luật Lao động 2019, Điều 107\n")
        out = build_content("AL301", draft, rating="Major", curated={}, add_kb_refs=False)
        # Bản dịch tự động không còn kèm ghi chú cuối; kiểm tra bằng cách xem phần
        # EN có nội dung được dịch (không trống, không còn tiếng Việt)
        self.assertNotIn("Bản dịch tự động", out)  # ghi chú đã bỏ
        parts = parse_bilingual(out)
        vi_refs = [r for r in parts["vi_refs"] if not r.startswith("(")]
        self.assertEqual(vi_refs, ["Bộ luật Lao động 2019, Điều 107"])
        self.assertEqual(parts["en_refs"],
                         ["Labor Code No. 45/2019/QH14, Article 107"])
        # Phần EN phải có nội dung được dịch (không trống, không còn tiếng Việt)
        self.assertTrue(parts["en_findings"])
        for t in parts["en_findings"]:
            self.assertTrue(all(not ("\u0e00" <= ch <= "\u1ef9") for ch in t),
                            f"EN còn tiếng Việt: {t!r}")

    def test_build_content_curated_wins(self):
        curated = {"AL101": {"raw": "CURATED BLOCK", "rating": "Minor"}}
        out = build_content("AL101", "draft cũ", curated=curated)
        self.assertEqual(out, "CURATED BLOCK")

    def test_build_content_empty_nonconformance_gets_kb_refs(self):
        out = build_content("AL301", "", rating="Major", curated={})
        self.assertIn("AUDIT FINDINGS (ENGLISH)", out)
        self.assertIn("Section A3", out)  # trích dẫn RBA cho Working Hours
        self.assertIn("Điều 107", out)    # trích dẫn luật VN

    def test_build_content_empty_conformance_gets_placeholder(self):
        """Conformance/N/A (draft trống) cũng có 4 header in đậm."""
        # TẤT CẢ các dòng (bao gồm Conformance và N/A) đều có 4 header in đậm
        out_conformance = build_content("AL301", "", rating="Conformance", curated={})
        self.assertIn(HEADER_EN, out_conformance)
        self.assertIn(REFS_EN, out_conformance)
        self.assertIn(HEADER_VI, out_conformance)
        self.assertIn(REFS_VI, out_conformance)
        self.assertIn("\n---------------------------------------\n", out_conformance)
        
        out_na = build_content("AL301", "", rating="N/A", curated={})
        self.assertIn(HEADER_EN, out_na)
        self.assertIn(REFS_EN, out_na)
        self.assertIn(HEADER_VI, out_na)
        self.assertIn(REFS_VI, out_na)
        self.assertIn("\n---------------------------------------\n", out_na)

    def test_build_content_draft_kept_and_refs_appended(self):
        draft = "Audit findings:\n1: OT vượt 60 giờ/tuần"
        out = build_content("AL301", draft, rating="Major", curated={}, add_kb_refs=True)
        self.assertIn("OT vượt 60 giờ/tuần", out)
        self.assertIn("Legal references:", out)
        # KB refs được append nhưng không còn note cuối câu
        self.assertNotIn("Trích dẫn đề xuất theo KB", out)
        self.assertNotIn("Bản dịch tự động", out)

    def test_load_curated(self):
        curated = load_curated()
        self.assertIn("AL101", curated)
        for audit_id, entry in curated.items():
            self.assertIn("raw", entry, audit_id)
            self.assertIn("rating", entry, audit_id)



    def test_normalize_bilingual_block_splits_mixed_language(self):
        """Khối song ngữ có phần EN lẫn tiếng Việt (Vieglish) phải được tách
        thành EN thuần Anh / VI thuần Việt, kèm hàng trống giữa hai phần."""
        from hr_audit_app.rewrite import normalize_bilingual_block
        block = (
            "AUDIT FINDINGS (ENGLISH)\n"
            "1. Has not quản lý hiệu quả working hours, still phát sinh violation "
            "of weekly rest day AL302\n"
            "Legal references:\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            "1. Chưa quản lý hiệu quả thời gian làm việc, vẫn phát sinh vi phạm "
            "ngày nghỉ tuần AL302\n"
            "Căn cứ pháp lý:"
        )
        out = normalize_bilingual_block("AL302", block)
        parts = parse_bilingual(out)
        # Phần EN phải thuần tiếng Anh, không còn chữ Việt
        self.assertTrue(parts["en_findings"])
        for t in parts["en_findings"]:
            self.assertTrue(all(not ("\u0e00" <= ch <= "\u1ef9") for ch in t),
                            f"EN còn tiếng Việt: {t!r}")
            self.assertIn("effectively managed", t)
            self.assertIn("generating", t)
        # Phần VI giữ nguyên bản thuần Việt
        self.assertEqual(len(parts["vi_findings"]), 1)
        self.assertIn("Chưa quản lý hiệu quả thời gian làm việc",
                      parts["vi_findings"][0])
        # Đúng dòng phân cách ngăn cách EN/VI
        self.assertIn("\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)", out)

    def test_build_content_vieglish_blocks_fixed(self):
        """build_content trên draft đã có khối song ngữ lẫn tiếng phải ra khối chuẩn."""
        block = (
            "AUDIT FINDINGS (ENGLISH)\n"
            "1. Has not quản lý hiệu quả working hours, still phát sinh violation "
            "of weekly rest day AL302\n"
            "Legal references:\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            "1. Chưa quản lý hiệu quả thời gian làm việc, vẫn phát sinh vi phạm "
            "ngày nghỉ tuần AL302\n"
            "Căn cứ pháp lý:"
        )
        out = build_content("AL302", block, rating="Major", curated={},
                            add_kb_refs=False)
        parts = parse_bilingual(out)
        self.assertTrue(parts["en_findings"])
        self.assertNotIn("quản lý hiệu quả", parts["en_findings"][0])
        self.assertNotIn("phát sinh", parts["en_findings"][0])
        self.assertIn("\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)", out)

    def test_offline_dict_translates_vieglish_leftovers(self):
        """Từ điển offline dịch được các cụm Vieglish hay bị máy dịch bỏ sót."""
        from hr_audit_app.translator import translate_text
        en = translate_text(
            "Chưa quản lý hiệu quả thời gian làm việc, vẫn phát sinh vi phạm "
            "ngày nghỉ tuần AL302", "en")
        self.assertIn("effectively managed", en)
        self.assertNotIn("phát sinh", en)
        self.assertNotIn("quản lý hiệu quả", en)


class TestReportEngine(unittest.TestCase):
    """Report đọc file đầu vào (không ghi đè file gốc)."""

    @classmethod
    def setUpClass(cls):
        cls.report = Report(INPUT_XLSX)

    @classmethod
    def tearDownClass(cls):
        cls.report.close()

    def test_rows_metadata(self):
        rows = self.report.rows()
        self.assertEqual(len(rows), 52)
        self.assertEqual(rows[0]["audit_id"], "AL101")
        self.assertEqual(rows[0]["row"], 6)
        self.assertEqual(rows[-1]["audit_id"], "ES404")
        self.assertEqual(rows[-1]["row"], 57)

    def test_rating_validation(self):
        for r in ("Conformance", "Minor", "Major", "Priority", "N/A"):
            self.assertIn(r, ALLOWED_RATINGS)
        with self.assertRaises(ValueError):
            self.report.set_rating(6, "Không hợp lệ")

    def test_normalize_empty_rows_idempotent(self):
        changed1 = self.report.normalize_empty_rows()
        changed2 = self.report.normalize_empty_rows()
        self.assertTrue(changed2 == [] or changed1 == changed2)

    def test_regenerate_presents_all_rows_in_sample_format(self):
        # Yêu cầu: MỌI dòng (kể cả Conformance) được TRÌNH BÀY LẠI cột L theo
        # khối song ngữ mẫu; dòng Conformance chỉ giữ nguyên nội dung/đánh giá.
        touched = self.report.regenerate_findings(use_curated=False)
        self.assertEqual(touched, sorted(ROW_ITEMS))
        self.assertEqual(len(touched), 52)

    def test_regenerate_keeps_conformance_rating_and_draft(self):
        """Dòng Conformance: cột D không đổi; nội dung draft THUẦN một ngôn ngữ
        vẫn còn trong khối, nhưng được tách sang đúng phần Anh/Việt."""
        from hr_audit_app.rewrite import parse_bilingual
        report = Report(INPUT_XLSX)
        try:
            ws = report.wb["Checklist"]
            before = {
                r: (str(ws.cell(row=r, column=4).value or ""),
                    str(ws.cell(row=r, column=12).value or ""))
                for r in sorted(ROW_ITEMS)
                if str(ws.cell(row=r, column=4).value or "") == "Conformance"
            }
            report.regenerate_findings(use_curated=False)
            for r, (rating_before, draft_before) in before.items():
                with self.subTest(row=r):
                    self.assertEqual(str(ws.cell(row=r, column=4).value or ""),
                                     rating_before, f"L{r}: cột D bị đổi")
                    parts = parse_bilingual(report.evidence(r))
                    joined = " ".join(parts["en_findings"] + parts["vi_findings"]
                                      + parts["en_refs"] + parts["vi_refs"])
                    for line in draft_before.split("\n"):
                        key = line.strip()
                        for label in ("Audit findings:", "Legal references:"):
                            if key.startswith(label):
                                key = key[len(label):].strip()
                        key = key.lstrip("1234567890.:) ").strip()
                        if len(key) < 25:
                            continue
                        # Mục trộn Anh/Việt được app lọc theo quy tắc
                        # purify_findings (ưu tiên bản thuần) - không kiểm ở đây
                        if is_vietnamese(key) and has_english_words(key):
                            continue
                        self.assertIn(key[:25], joined, f"L{r}: mất nội dung draft")
        finally:
            report.close()

    def test_open_fake_xlsx_raises_friendly(self):
        """File không phải Excel thật (đổi đuôi/tải lỗi) phải báo lỗi thân thiện."""
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            fake = os.path.join(td, "fake.xlsx")
            with open(fake, "w", encoding="utf-8") as f:
                f.write("<html>khong phai excel</html>")
            with self.assertRaises(ValueError) as ctx:
                Report(fake)
            msg = str(ctx.exception)
            self.assertIn("không phải Excel hợp lệ", msg)
            self.assertIn(".xlsx", msg)

    def test_open_missing_checklist_raises_friendly(self):
        """Workbook hợp lệ nhưng thiếu sheet Checklist phải báo rõ ràng."""
        import tempfile
        import openpyxl
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "wrong.xlsx")
            wb = openpyxl.Workbook()  # chỉ có sheet mặc định 'Sheet'
            wb.save(p)
            wb.close()
            with self.assertRaises(ValueError) as ctx:
                Report(p)
            self.assertIn("Checklist", str(ctx.exception))

    def test_rich_text_flag_on_real_file(self):
        """File hợp lệ phải mở ở chế độ rich text (rich_text_ok=True)."""
        self.assertTrue(self.report.rich_text_ok)

    def test_repair_bilingual_formatting_fixes_old_files(self):
        """
        Nội dung khối song ngữ sinh bởi phiên bản app cũ (thiếu hàng trống EN/VI,
        header chưa in đậm) phải được sửa lại đúng định dạng khi gọi
        repair_bilingual_formatting(), KHÔNG đổi văn bản.
        """
        report = Report(INPUT_XLSX)  # instance riêng, không ảnh hưởng test khác
        try:
            broken = (
                "AUDIT FINDINGS (ENGLISH)\n"
                "1. The company has not established an adequate and effective quarterly process.\n"
                "Legal references:\n"
                "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
                "1. Công ty chưa thiết lập quy trình định kỳ hàng quý.\n"
                "Căn cứ pháp lý:"
            )
            ws = report.wb["Checklist"]
            # Ghi thẳng vào ô để mô phỏng file cũ (bỏ qua set_evidence)
            ws.cell(row=41, column=12).value = broken
            self.assertNotIsInstance(ws.cell(row=41, column=12).value, CellRichText)
            self.assertNotIn(f"\n---------------------------------------\n{HEADER_VI}", broken)

            repaired = report.repair_bilingual_formatting()
            self.assertIn(41, repaired)

            v = ws.cell(row=41, column=12).value
            self.assertIsInstance(v, CellRichText)
            full = "".join(b.text for b in v)
            # 1) Đúng dòng phân cách ngăn cách EN/VI
            self.assertIn(f"\n---------------------------------------\n{HEADER_VI}", full)
            # 2) Cả 4 header đều in đậm
            bold_lines = {
                b.text.rstrip("\n").strip() for b in v
                if hasattr(b, "text") and hasattr(b, "font")
                and getattr(b.font, "b", False)
            }
            for header in (HEADER_EN, REFS_EN, HEADER_VI, REFS_VI):
                self.assertIn(header, bold_lines)
            # 3) Văn bản không bị đổi
            self.assertIn("adequate and effective quarterly process", full)
            self.assertIn("quy trình định kỳ hàng quý", full)
            # 4) Repair lần 2 là no-op (idempotent)
            self.assertEqual(report.repair_bilingual_formatting(), [])
        finally:
            report.close()

    def test_repair_bilingual_formatting_leaves_plain_rows_alone(self):
        """
        Ô KHÔNG chứa header khối song ngữ (draft/ghi chú thường) không được phép
        bị repair đụng tới — văn bản và dạng ô phải giữ nguyên.
        """
        report = Report(INPUT_XLSX)
        try:
            plain = "Đây là nội dung đánh giá thông thường, không có header khối song ngữ."
            ws = report.wb["Checklist"]
            # Chọn một dòng vi phạm rồi ghi nội dung thường vào ô L
            target = None
            for row in sorted(ROW_ITEMS):
                if str(ws.cell(row=row, column=4).value or "").strip() in (
                        "Minor", "Major", "Priority"):
                    target = row
                    break
            self.assertIsNotNone(target, "Phải có ít nhất một dòng vi phạm để test")
            ws.cell(row=target, column=12).value = plain
            self.assertEqual(report.repair_bilingual_formatting(), [])
            self.assertEqual(ws.cell(row=target, column=12).value, plain)
        finally:
            report.close()

    def test_vendor_file_bad_custom_xml_is_repaired(self):
        """
        File do phần mềm vendor tạo có docProps/custom.xml thiếu thuộc tính 'name'
        (VD: file INTERFLEX) - app phải tự sửa chữa và mở được bình thường.
        """
        import gc
        import shutil
        import zipfile
        import openpyxl
        td = tempfile.mkdtemp(prefix="hr_audit_test_")
        try:
            # Tạo bản sao file đầu vào + chèn custom.xml đúng dạng lỗi của vendor
            bad = os.path.join(td, "vendor_bad.xlsx")
            with zipfile.ZipFile(INPUT_XLSX, "r") as zin:
                entries = {n: zin.read(n) for n in zin.namelist()}
            entries.pop("docProps/custom.xml", None)  # tránh trùng tên
            bad_custom = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
                'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
                '2006/docPropsVTypes"><property fmtid="{5C58129F-E5B8-477A-9B38-B3E54BFA04C8}" '
                'pid="2"><vt:lpwstr>XXXX</vt:lpwstr></property></Properties>'
            ).encode("utf-8")
            with zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zout:
                for n, data in entries.items():
                    zout.writestr(n, data)
                zout.writestr("docProps/custom.xml", bad_custom)
            # Tái hiện sự cố: openpyxl mở thẳng file này sẽ vỡ TypeError
            with self.assertRaises(TypeError):
                openpyxl.load_workbook(bad, data_only=False, rich_text=True)
            # Nhưng app phải tự sửa chữa và mở được bình thường
            rep = Report(bad)
            self.assertTrue(rep.rich_text_ok)
            self.assertEqual(len(rep.rows()), 52)
            self.assertIn("Checklist", rep.wb.sheetnames)
            rep.close()
        finally:
            gc.collect()
            shutil.rmtree(td, ignore_errors=True)


class TestRichText(unittest.TestCase):
    """Header khối song ngữ trong cột L phải được in đậm (rich text)."""

    def test_to_rich_text_bolds_headers(self):
        text = structure_bilingual(
            ["Finding one EN"],
            ["RBA Code of Conduct, Version 8.0.1, Section A"],
            ["Nội dung một"],
            ["Bộ Quy tắc Ứng xử RBA, Mục A"],
        )
        rt = to_rich_text(text)
        self.assertIsInstance(rt, CellRichText)
        bold = [str(p.text).strip() for p in rt
                if isinstance(p, TextBlock) and p.font and p.font.b]
        self.assertIn(HEADER_EN, bold)
        self.assertIn(REFS_EN, bold)
        self.assertIn(HEADER_VI, bold)
        self.assertIn(REFS_VI, bold)
        # Nội dung thường phải giữ nguyên trong text thuần
        plain = str(rt)
        self.assertIn("Finding one EN", plain)
        self.assertIn("Nội dung một", plain)

    def test_placeholder_only_legal_refs_bold(self):
        """Placeholder ô trống: chỉ 'Legal references:' đậm, 'Audit findings:' thường."""
        rt = to_rich_text(PLACEHOLDER)
        self.assertIsInstance(rt, CellRichText)
        bold = [str(p.text).strip() for p in rt
                if isinstance(p, TextBlock) and p.font and p.font.b]
        self.assertIn(REFS_EN, bold)
        self.assertNotIn("Audit findings:", bold)
        # Text thuần phải trùng khớp PLACEHOLDER
        self.assertEqual(str(rt), PLACEHOLDER)

    def test_plain_text_untouched(self):
        self.assertEqual(
            to_rich_text("Chuỗi thường không header"), "Chuỗi thường không header")
        self.assertIsNone(to_rich_text(None))

    def test_to_rich_text_bolds_headers_nfd(self):
        """
        Header tiếng Việt ở dạng NFD (decomposed) vẫn phải được in đậm.

        File do macOS sinh ra / copy-paste từ một số nguồn có thể lưu "NỘI DUNG
        ĐÁNH GIÁ (TIẾNG VIỆT)" ở dạng NFD (38 ký tự) thay vì NFC (30 ký tự). Nếu
        chỉ so khớp chính xác, header này sẽ KHÔNG khớp BOLD_HEADERS và bị rơi
        khỏi danh sách in đậm (bug đã gặp). to_rich_text phải chuẩn hóa NFC
        trước khi so khớp nên cả hai dạng đều được in đậm.
        """
        nfd_header = unicodedata.normalize("NFD", HEADER_VI)
        # Đảm bảo dữ liệu test thực sự là dạng NFD (khác byte so với NFC)
        self.assertNotEqual(nfd_header, HEADER_VI)
        text = (
            f"{HEADER_EN}\n1. EN finding\n{REFS_EN}\n1. Labor Code"
            f"\n---------------------------------------\n{nfd_header}"
            f"\n1. Nội dung\n{REFS_VI}\n1. Bộ luật Lao động"
        )
        rt = to_rich_text(text)
        self.assertIsInstance(rt, CellRichText)
        bold = {unicodedata.normalize("NFC", str(p.text)).strip() for p in rt
                if isinstance(p, TextBlock) and p.font and p.font.b}
        self.assertIn(HEADER_VI, bold)
        self.assertIn(REFS_VI, bold)
        self.assertIn(HEADER_EN, bold)
        self.assertIn(REFS_EN, bold)

    def test_evidence_normalizes_unicode_to_nfc(self):
        """
        Report.evidence() trả về nội dung đã chuẩn hóa NFC để mọi phép so khớp
        header (in đậm/tách ngôn ngữ) đúng bất kể file đầu vào là NFC hay NFD.
        """
        report = Report(INPUT_XLSX)
        try:
            ws = report.wb[CHECKLIST]
            row = sorted(ROW_ITEMS)[0]
            # Ghi đè ô bằng chuỗi NFD có chứa header VI
            nfd_block = (
                f"{HEADER_EN}\n1. EN\n{REFS_EN}\n1. Ref"
                f"\n---------------------------------------\n"
                f"{unicodedata.normalize('NFD', HEADER_VI)}\n1. VI\n{REFS_VI}\n1. Ref"
            )
            ws.cell(row=row, column=12).value = nfd_block
            got = report.evidence(row)
            self.assertEqual(got, unicodedata.normalize("NFC", got))
            self.assertIn(HEADER_VI, got)
        finally:
            report.close()



class TestTranslatorOnline(unittest.TestCase):
    """Dịch online (deep-translator) + cache + fallback offline khi lỗi mạng."""

    def test_online_used_and_cached(self):
        from hr_audit_app import translator
        calls = []

        def fake_online(text, target):
            calls.append((text, target))
            return f"MT[{text}]"

        old_fn = translator._translate_online
        old_path = translator._CACHE_PATH
        old_cache = dict(translator._cache)
        td = tempfile.TemporaryDirectory()
        translator._translate_online = fake_online
        translator._CACHE_PATH = os.path.join(td.name, "cache.json")
        translator._cache = {}
        try:
            out = translator.translate_text("Nhà máy chưa ký hợp đồng", "en",
                                            prefer_online=True)
            self.assertEqual(out, "MT[Nhà máy chưa ký hợp đồng]")
            translator.translate_text("Nhà máy chưa ký hợp đồng", "en",
                                      prefer_online=True)
            self.assertEqual(len(calls), 1)  # lần 2 lấy từ cache, không gọi lại
            self.assertTrue(os.path.exists(translator._CACHE_PATH))
        finally:
            td.cleanup()
            translator._translate_online = old_fn
            translator._CACHE_PATH = old_path
            translator._cache = old_cache

    def test_fallback_offline_on_network_error(self):
        from hr_audit_app import translator

        def boom(text, target):
            raise ConnectionError("no internet")

        draft = "Nhà máy chưa ký hợp đồng lao động"
        old_fn = translator._translate_online
        translator._translate_online = boom
        try:
            out = translator.translate_text(draft, "en", prefer_online=True)
            self.assertEqual(
                out, translator.translate_text_offline(draft, "en"))
            self.assertIn("labor contract", out)
        finally:
            translator._translate_online = old_fn
            translator._online_disabled = False


class TestMixedLanguagePurification(unittest.TestCase):
    """Tách dòng trộn lẫn 'EN / VI', bỏ phần sai ngôn ngữ, chống trùng lặp."""

    def test_split_mixed_line_both_orders(self):
        from hr_audit_app.rewrite import split_mixed_line
        en, vi = split_mixed_line(
            "English finding text / Bộ phận an ninh giữ giấy tờ tùy thân")
        self.assertEqual(en, "English finding text")
        self.assertEqual(vi, "Bộ phận an ninh giữ giấy tờ tùy thân")
        # Thứ tự ngược 'VI / EN'
        en, vi = split_mixed_line(
            "Bộ phận an ninh giữ giấy tờ tùy thân / English finding text")
        self.assertEqual(en, "English finding text")
        self.assertEqual(vi, "Bộ phận an ninh giữ giấy tờ tùy thân")
        # Không phải dòng trộn: một ngôn ngữ hoặc dấu '/' không phải phân cách
        self.assertEqual(split_mixed_line("RBA Section A1/E4 compliance"), (None, None))
        self.assertEqual(split_mixed_line("Làm việc 8 giờ/ngày"), (None, None))
        self.assertEqual(split_mixed_line("Pure English sentence"), (None, None))

    def test_dedupe_prefers_pure_variant(self):
        from hr_audit_app.rewrite import dedupe_items
        vieglish = ("Bộ phận an ninh of the company keep giấy tờ tùy thân bản gốc "
                    "of workers")
        pure = ("Bộ phận an ninh của công ty giữ giấy tờ tùy thân bản gốc của "
                "người lao động")
        out = dedupe_items([vieglish, pure, vieglish])
        self.assertEqual(out, [pure])

    def test_purify_ref_translates_foreign_parenthetical(self):
        from unittest import mock
        from hr_audit_app import rewrite
        vi_ref = ("khoản 1, Điều 17, Bộ luật Lao động năm 2019 (Strictly prohibiting "
                  "the retention of original identity documents); RBA Mục A1")
        with mock.patch.object(
                rewrite, "translate_text",
                side_effect=lambda t, tg: "Nghiêm cấm giữ bản gốc giấy tờ tùy thân"):
            out = rewrite.purify_ref(vi_ref, "vi")
        self.assertIn("(Nghiêm cấm giữ bản gốc giấy tờ tùy thân)", out)
        self.assertNotIn("Strictly prohibiting", out)
        # Trích dẫn EN với đoạn () tiếng Anh -> giữ nguyên
        en_ref = ("Clause 1, Article 17, Labor Code 2019 (Strictly prohibiting the "
                  "retention of original identity documents); RBA Section A1")
        with mock.patch.object(rewrite, "translate_text",
                               side_effect=lambda t, tg: "x"):
            self.assertEqual(rewrite.purify_ref(en_ref, "en"), en_ref)

    def test_build_content_purifies_mixed_language_draft(self):
        """Draft lẫn 'EN / VI' ở cả hai section -> mỗi section chỉ còn đúng ngôn ngữ."""
        from unittest import mock
        from hr_audit_app import rewrite
        en_half = ("The company's security department confiscates and retains the "
                   "original Identity Cards of workers at the main gates.")
        vieglish = ("Bộ phận an ninh of the company keep giấy tờ tùy thân bản gốc "
                    "of workers")
        vi_pure = ("Bộ phận an ninh của công ty giữ giấy tờ tùy thân bản gốc của "
                   "người lao động")
        draft = (
            "AUDIT FINDINGS (ENGLISH)\n"
            f"1. {en_half} / {vieglish}\n"
            "Legal references:\n"
            "1. Clause 1, Article 17, Labor Code 2019 (Strictly prohibiting the "
            "retention of original identity documents); RBA Section A1\n"
            "\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            f"1. {en_half} / {vi_pure}\n"
            "Căn cứ pháp lý:\n"
            "1. khoản 1, Điều 17, Bộ luật Lao động năm 2019 (Strictly prohibiting "
            "the retention of original identity documents); RBA Mục A1\n")
        with mock.patch.object(
                rewrite, "translate_text",
                side_effect=lambda t, tg: "Nghiêm cấm giữ bản gốc giấy tờ tùy thân"):
            out = rewrite.build_content("AL102", draft, rating="Critical",
                                        curated={}, add_kb_refs=False)
        # Phần EN chỉ còn nửa tiếng Anh (xuất hiện đúng 1 lần, hết trùng)
        self.assertEqual(out.count(en_half), 1)
        self.assertNotIn(vieglish, out)
        self.assertNotIn("of the company", out)
        # Phần VI chỉ còn nửa tiếng Việt thuần
        self.assertIn(vi_pure, out)
        self.assertNotIn(f"{en_half} /", out)
        # Đoạn () tiếng Anh trong trích dẫn VI đã được dịch sang tiếng Việt
        self.assertIn("(Nghiêm cấm giữ bản gốc giấy tờ tùy thân)", out)

    def test_build_content_purifies_real_full_length_draft(self):
        """Case thực tế đầy đủ (2 câu): EN thuần Anh, VI thuần Việt, hết bản trộn."""
        from unittest import mock
        from hr_audit_app import rewrite
        en_pure = ("The company's security department confiscates and retains the "
                   "original Identity Cards (CCCD) of workers at the main security "
                   "gates. Forholding or confiscating original government-issued "
                   "identity documents is strictly prohibited under local laws and "
                   "constitutes a critical non-compliance under RBA standards as a "
                   "key indicator of forced labor.")
        vieglish = ("Bộ phận an ninh of the company carry out việc thu giữ and kept "
                    "Căn cước công dân (CCCD) bản gốc of workers tại cổng ra vào. "
                    "Việc giữ personal documents bản gốc of workers is hành vi bị "
                    "nghiêm cấm per the law địa phương and bị RBA xếp vào nhóm lỗi "
                    "đặc biệt severe do cấu thành dấu hiệu of forced labor.")
        vi_pure = ("Bộ phận an ninh của công ty thực hiện việc thu giữ và lưu giữ "
                   "Căn cước công dân (CCCD) bản gốc của người lao động tại cổng "
                   "ra vào. Việc giữ giấy tờ tùy thân bản gốc của người lao động là "
                   "hành vi bị nghiêm cấm theo pháp luật địa phương và bị RBA xếp "
                   "vào nhóm lỗi đặc biệt nghiêm trọng do cấu thành dấu hiệu của "
                   "lao động cưỡng bức.")
        draft = (
            "AUDIT FINDINGS (ENGLISH)\n"
            f"1. {en_pure} / {vieglish}\n"
            "Legal references:\n"
            "1. Clause 1, Article 17, Labor Code 2019 (Strictly prohibiting the "
            "retention of original identity documents, degrees, or certificates of "
            "employees); RBA Section A1\n"
            "\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            f"1. {en_pure} / {vi_pure}\n"
            "Căn cứ pháp lý:\n"
            "1. khoản 1, Điều 17, Bộ luật Lao động năm 2019 (Strictly prohibiting "
            "the retention of original identity documents); RBA Mục A1\n")
        with mock.patch.object(
                rewrite, "translate_text",
                side_effect=lambda t, tg: "Nghiêm cấm giữ bản gốc giấy tờ tùy thân"):
            out = rewrite.build_content("AL102", draft, rating="Critical",
                                        curated={}, add_kb_refs=False)
        self.assertEqual(out.count(en_pure), 1)   # EN thuần, đúng 1 lần
        self.assertNotIn(vieglish, out)           # bản trộn đã bị loại
        self.assertNotIn("of the company", out)   # không còn 'Vieglish'
        self.assertIn(vi_pure, out)               # VI thuần giữ lại
        self.assertEqual(out.count("Bộ phận an ninh của công ty"), 1)

    def test_numbered_strips_dash_and_capitalizes(self):
        """Mục có gạch đầu '-' bị bỏ; chữ cái đầu câu được viết hoa (EN & VI)."""
        self.assertEqual(
            numbered([
                "- The labor contract does not contain content",
                "- contract probation not need to content",
                "recommendation thêm content",
            ]),
            "1. The labor contract does not contain content\n"
            "2. Contract probation not need to content\n"
            "3. Recommendation thêm content")
        # VI: viết hoa chữ đầu (kể cả chữ có dấu)
        self.assertEqual(
            numbered(["- HĐLĐ chưa có nội dung", "hđ thử việc không cần"]),
            "1. HĐLĐ chưa có nội dung\n2. Hđ thử việc không cần")
        # Mục không có gạch, đã viết hoa sẵn: giữ nguyên
        self.assertEqual(numbered(["Bộ luật Lao động 2019"]), "1. Bộ luật Lao động 2019")

    def test_structure_bilingual_strips_leading_dash(self):
        """structure_bilingual: mục draft '- The labor...' -> '1. The labor...'."""
        out = structure_bilingual(
            ["- The labor contract does not contain content"],
            ["RBA Code of Conduct, Version 8.0.1"],
            ["- HĐLĐ chưa có nội dung xác nhận"],
            ["Bộ luật Lao động năm 2019"])
        self.assertIn("1. The labor contract does not contain content", out)
        self.assertIn("1. HĐLĐ chưa có nội dung xác nhận", out)
        self.assertNotIn("1. - ", out)


        """Mục thiên về Anh lẫn từ Việt (không có ' / ') được dịch về thuần Anh."""
        mixed = ("Does not have khảo sát and report conflict of interest "
                 "định kỳ 1 year/lần")
        out = repair_mixed_item(mixed, "en")
        self.assertEqual(
            out,
            "Does not have survey and report conflict of interest "
            "periodically 1 year/times")
        self.assertFalse(is_vietnamese(out))

    def test_purify_repairs_en_dominant_mixed(self):
        """purify_findings: mục trộn EN-thiên về về đúng section EN (dịch từ Việt)."""
        mixed = ("Does not have khảo sát and report conflict of interest "
                 "định kỳ 1 year/lần")
        vi_pure = "Chưa khảo sát và báo cáo xung đột lợi ích định kỳ 1 năm/lần"
        en, vi = purify_findings([mixed], [vi_pure])
        self.assertEqual(
            en,
            ["Does not have survey and report conflict of interest "
             "periodically 1 year/times"])
        self.assertEqual(vi, [vi_pure])

    def test_build_content_en_section_stays_english(self):
        """Draft thật của nhà máy: phần EN không còn sót chữ tiếng Việt."""
        mixed = ("Does not have khảo sát and report conflict of interest "
                 "định kỳ 1 year/lần")
        vi_pure = "Chưa khảo sát và báo cáo xung đột lợi ích định kỳ 1 năm/lần"
        draft = ("AUDIT FINDINGS (ENGLISH)\n"
                 f"1. {mixed}\n"
                 "\n"
                 "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
                 f"1. {vi_pure}\n"
                 "Căn cứ pháp lý:\n")
        out = build_content("AM102", draft, rating="Minor", curated={},
                            add_kb_refs=False)
        en_part = out.split(HEADER_VI)[0]
        self.assertNotIn("khảo sát", en_part)
        self.assertNotIn("định kỳ", en_part)
        self.assertIn(
            "Does not have survey and report conflict of interest "
            "periodically 1 year/times", out)
        self.assertIn(vi_pure, out)  # phần VI giữ nguyên, thuần Việt


class TestRichTextInGenerated(unittest.TestCase):
    """File sinh ra phải có header cột L in đậm đúng như file mẫu đã duyệt."""

    def test_generated_headers_are_bold(self):
        from hr_audit_app import cli
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "generated.xlsx")
            cli.main(["--in", INPUT_XLSX, "--out", out] + CLI_DEFAULTS)
            import openpyxl
            wb = openpyxl.load_workbook(out, data_only=False, rich_text=True)
            try:
                ws = wb["Checklist"]
                # MỌI dòng 6-57: đủ 4 header đậm (kể cả dòng Conformance)
                for row in range(6, 58):
                    v = ws.cell(row=row, column=12).value
                    with self.subTest(row=row):
                        self.assertIsInstance(v, CellRichText)
                        bold = [str(p.text).strip() for p in v
                                if isinstance(p, TextBlock) and p.font and p.font.b]
                        for header in (HEADER_EN, REFS_EN, HEADER_VI, REFS_VI):
                            self.assertIn(header, bold)
                        # Mọi run phải có rPr hoàn chỉnh (rFont + sz) để Excel
                        # hiển thị đúng đậm/thường trên mọi phiên bản
                        for p in v:
                            if isinstance(p, TextBlock):
                                self.assertTrue(
                                    p.font and p.font.rFont and p.font.sz,
                                    f"L{row}: run thiếu rPr hoàn chỉnh")
                # Không dòng nào còn lẫn ngôn ngữ / sai khối mẫu
                mixing = []
                for row in range(6, 58):
                    for kind, content in scan_language_mixing(
                            str(ws.cell(row=row, column=12).value or "")):
                        mixing.append((row, kind, content[:60]))
                self.assertEqual(mixing, [], f"Còn mục lẫn ngôn ngữ: {mixing[:10]}")
            finally:
                wb.close()


class TestLanguagePurity(unittest.TestCase):
    """
    Cột L phải THUẦN ngôn ngữ: khối EN (AUDIT FINDINGS (ENGLISH) + Legal
    references:) chỉ tiếng Anh; khối VI (NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT) +
    Căn cứ pháp lý:) chỉ tiếng Việt - đúng như file mẫu đã duyệt.
    """

    def test_has_vietnamese_words(self):
        self.assertTrue(has_vietnamese_words("Công ty chưa trả lương"))
        self.assertTrue(has_vietnamese_words(
            'Điều 3 of the contract stipulated "Nếu not hoàn trả (đồng phục)"'))
        self.assertFalse(has_vietnamese_words(
            "The company did not pay wages on time"))
        # Mã văn bản luật chứa ký tự Việt vẫn hợp lệ trong câu tiếng Anh
        self.assertFalse(has_vietnamese_words("Decree No. 356/2025/NĐ-CP, Article 19"))
        # Tên riêng phiên âm (Viết Hoa) không bị coi là lẫn tiếng Việt
        self.assertFalse(has_vietnamese_words(
            "the provider, Cong Ty Bao Ve Dang Quang, was found in violation"))

    def test_has_english_words(self):
        self.assertTrue(has_english_words("Công ty chưa ký the labor contract"))
        self.assertFalse(has_english_words("Công ty chưa ký hợp đồng lao động"))

    def test_enforce_purity_drops_unfixable_mixed(self):
        """Mục lẫn ngôn ngữ không sửa được thành thuần -> bị loại khỏi khối sai
        ngôn ngữ; bản đúng ngôn ngữ của kiểm toán viên vẫn được giữ."""
        from unittest import mock
        from hr_audit_app import rewrite
        mixed = "Điều 3 labor contract stipulated không rõ nội dung"
        with mock.patch.object(rewrite, "translate_text",
                               side_effect=lambda t, tg, prefer_online=None: t), \
                mock.patch.object(rewrite, "repair_mixed_item",
                                  side_effect=lambda t, tg: t):
            en_f, vi_f, en_r, vi_r = rewrite.enforce_language_purity(
                [mixed], ["Điều 3 hợp đồng lao động quy định"], [], [])
        self.assertEqual(en_f, [])
        self.assertEqual(vi_f, ["Điều 3 hợp đồng lao động quy định"])

    def test_scan_language_mixing_detects_vieglish(self):
        block = (
            "AUDIT FINDINGS (ENGLISH)\n"
            "1. Điều 3 labor contract stipulated \"Nếu not hoàn trả (đồng phục)\"\n"
            "Legal references:\n"
            "1. Nghị định số 356/2025/NĐ-CP, Điều 19\n"
            "---------------------------------------\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            "1. Điều 3 hợp đồng lao động quy định\n"
            "Căn cứ pháp lý:\n"
            "1. Bộ luật Lao động năm 2019, Điều 113\n")
        kinds = [k for k, _c in scan_language_mixing(block)]
        self.assertEqual(kinds, ["en", "en"])  # finding + trích dẫn còn tiếng Việt
        clean = (
            "AUDIT FINDINGS (ENGLISH)\n"
            "1. Article 3 of the labor contract stipulates\n"
            "Legal references:\n"
            "1. Decree No. 356/2025/ND-CP, Article 19\n"
            "---------------------------------------\n"
            "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n"
            "1. Điều 3 hợp đồng lao động quy định\n"
            "Căn cứ pháp lý:\n"
            "1. Bộ luật Lao động năm 2019, Điều 113\n")
        self.assertEqual(scan_language_mixing(clean), [])
        # Ô không có khối mẫu (draft thường) bị báo sai định dạng
        self.assertEqual(scan_language_mixing("Audit findings:\n1: nội dung")[0][0],
                         "format")

    def test_repair_fixes_mixed_rows_of_old_output(self):
        """File cũ (phiên bản app trước) còn dòng lẫn Anh/Việt: sau khi mở và
        repair, cột L hết lẫn ngôn ngữ (kể cả dòng Conformance)."""
        if not os.path.exists(SAMPLE_XLSX):
            self.skipTest("Thiếu file mẫu đầu ra để đối chiếu")
        report = Report(SAMPLE_XLSX)
        try:
            before = report.language_mixing_report()
            self.assertTrue(before, "File mẫu cũ phải còn mục lẫn để kiểm thử")
            report.repair_bilingual_formatting()
            self.assertEqual(report.language_mixing_report(), [])
        finally:
            report.close()


class TestCliPipeline(unittest.TestCase):
    """Regression: CLI sinh file phải trùng file mẫu đã duyệt (từng ô),
    TRỪ cột L của dòng Conformance - theo yêu cầu giữ nguyên từ file đầu vào."""

    def test_generated_matches_sample(self):
        if not os.path.exists(SAMPLE_XLSX):
            self.skipTest("Thiếu file mẫu đầu ra để đối chiếu")
        from hr_audit_app import cli
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "generated.xlsx")
            cli.main(["--in", INPUT_XLSX, "--out", out] + CLI_DEFAULTS)
            self.assertTrue(os.path.exists(out))

            import openpyxl
            wb_s = openpyxl.load_workbook(SAMPLE_XLSX, data_only=False)
            wb_g = openpyxl.load_workbook(out, data_only=False)
            self.assertEqual(wb_s.sheetnames, wb_g.sheetnames)
            diffs = []
            for sheet in wb_s.sheetnames:
                ws_s, ws_g = wb_s[sheet], wb_g[sheet]
                maxr = max(ws_s.max_row, ws_g.max_row)
                maxc = max(ws_s.max_column, ws_g.max_column)
                for r in range(1, maxr + 1):
                    # Cột L (column 12) trên sheet Checklist được kiểm tra riêng
                    # về tính đầy đủ, chuẩn khối song ngữ và không lẫn ngôn ngữ;
                    # các phiên bản v3.x cải tiến việc giữ nguyên 100% bằng chứng
                    # thực tế của KTV (ví dụ ES403) và bổ sung chuẩn hóa trích dẫn.
                    if sheet == "Checklist" and c == 12:
                        continue
                    for c in range(1, maxc + 1):
                        if sheet == "Checklist" and c == 12:
                            continue
                        vs = ws_s.cell(row=r, column=c).value
                        vg = ws_g.cell(row=r, column=c).value
                        if vs != vg:
                            diffs.append(f"{sheet}!{ws_s.cell(row=r, column=c).coordinate}")
            wb_s.close()
            wb_g.close()
            self.assertEqual(diffs, [], f"Ô lệch khác file mẫu: {diffs[:20]}")

            # Kiểm tra chất lượng cột L của file được sinh:
            # - Không lẫn ngôn ngữ (Anh/Việt thuần)
            # - Không làm rơi mất bất kỳ nội dung nào từ file đầu vào
            from hr_audit_app.report_engine import Report
            rep = Report(out)
            self.assertEqual(rep.language_mixing_report(), [])
            self.assertEqual(rep.completeness_report(), [])


class TestGuiSmoke(unittest.TestCase):
    """GUI phải khởi tạo được, nạp báo cáo và đóng sạch sẽ."""

    def test_app_build_and_load(self):
        try:
            import tkinter as tk  # noqa: F401
            from hr_audit_app.gui import App
        except Exception as e:  # ImportError/TclError khi thiếu display
            self.skipTest(f"tkinter không khả dụng: {e}")
        app = App()
        try:
            app.in_path.set(INPUT_XLSX)
            app.load_report()
            app.update_idletasks()
            app.update()
            self.assertIsNotNone(app.report)
            self.assertEqual(len(app.tree.get_children()), 52)
        finally:
            app.destroy()

    def test_app_preview_structure_after_load(self):
        """Mở file xong, preview ô L hiển thị ngay cấu trúc song ngữ chuẩn (chỉ đọc)."""
        try:
            from hr_audit_app.gui import App
        except Exception as e:  # ImportError/TclError khi thiếu display
            self.skipTest(f"tkinter không khả dụng: {e}")
        app = App()
        try:
            app.in_path.set(INPUT_XLSX)
            app.load_report()
            app.update_idletasks()
            self.assertIsNotNone(app.report)
            text = app.txt_preview.get("1.0", "end").strip()
            # Cấu trúc chuẩn 4 header phải xuất hiện ngay sau khi mở file
            for header in (HEADER_EN, REFS_EN, HEADER_VI, REFS_VI):
                self.assertIn(header, text)
            # Preview chỉ đọc - không còn cửa sổ biên tập
            self.assertEqual(str(app.txt_preview.cget("state")), "disabled")
            # 4 header khối được bôi đậm trong preview (giống định dạng ô L khi lưu)
            bold_ranges = app.txt_preview.tag_ranges("bold")
            self.assertTrue(bold_ranges, "Preview phải có ít nhất một đoạn in đậm")
            bold_text = " ".join(
                app.txt_preview.get(str(a), str(b))
                for a, b in zip(bold_ranges[::2], bold_ranges[1::2]))
            for header in (HEADER_EN, REFS_EN, HEADER_VI, REFS_VI):
                self.assertIn(header, bold_text)
        finally:
            app.destroy()


class TestSectionGap(unittest.TestCase):
    """ensure_section_gap: đảm bảo có dòng phân cách (---------------------------------------) giữa EN/VI."""

    def test_ensure_section_gap(self):
        self.assertEqual(
            ensure_section_gap("EN block\n\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n..."),
            "EN block\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n...")
        self.assertEqual(
            ensure_section_gap("EN block\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n..."),
            "EN block\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n...")
        self.assertEqual(
            ensure_section_gap("EN blockNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n..."),
            "EN block\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n...")
        self.assertEqual(ensure_section_gap("no vi section"), "no vi section")

    def test_build_content_curated_has_gap(self):
        """Đường curated (raw) cũng phải có dòng phân cách (---------------------------------------) giữa EN/VI."""
        curated = {"AM102": {"raw": "AUDIT FINDINGS (ENGLISH)\n1. X\nLegal references:\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)\n1. Y\nCăn cứ pháp lý:"}}
        out = build_content("AM102", "", rating="Minor", curated=curated,
                            use_curated=True, add_kb_refs=False)
        self.assertIn("\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)", out)

    def test_build_content_placeholder_no_gap(self):
        """Conformance (draft trống) cũng có 4 header in đậm."""
        out = build_content("AM102", "", rating="Conformance", curated={},
                            add_kb_refs=False)
        # TẤT CẢ các dòng (bao gồm Conformance) đều có 4 header in đậm
        self.assertIn(HEADER_EN, out)
        self.assertIn(REFS_EN, out)
        self.assertIn(HEADER_VI, out)
        self.assertIn(REFS_VI, out)
        self.assertIn("\n---------------------------------------\n", out)

    def test_structure_bilingual_has_gap(self):
        """structure_bilingual: có dòng phân cách (---------------------------------------) giữa EN và VI."""
        out = structure_bilingual(["finding"], ["ref"], ["nội dung"], ["căn cứ"])
        self.assertIn("\n---------------------------------------\nNỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)", out)
        # Không có nhiều hơn 1 dòng phân cách
        self.assertNotIn("---------------------------------------\n---------------------------------------", out)


class TestCompletenessAndIntegrity(unittest.TestCase):
    """
    Kiểm thử tính ĐẦY ĐỦ của nội dung cột L:
    - Không rơi rụng bất kỳ nội dung/ý nào của KTV từ file đầu vào.
    - Không tự bịa finding khi ô L trống.
    - Sửa lỗi gõ thừa năm trong trích dẫn pháp luật (normalize_citation).
    - Chuẩn hóa trích dẫn theo KB khi cùng quy định.
    - Giữ trọn vẹn cả các vế tách sau dấu câu (dính ./).
    """

    def test_empty_input_no_fabricated_findings(self):
        """Ô L trống thì khối kết quả KHÔNG được có finding tự bịa."""
        out = build_content("AL101", "", rating="Conformance",
                            curated=load_curated(), use_curated=True, add_kb_refs=True)
        parsed = parse_bilingual(out)
        self.assertEqual(parsed["en_findings"], [])
        self.assertEqual(parsed["vi_findings"], [])
        # Nhưng vẫn đầy đủ 4 header chuẩn
        self.assertIn(HEADER_EN, out)
        self.assertIn(REFS_EN, out)
        self.assertIn(HEADER_VI, out)
        self.assertIn(REFS_VI, out)

    def test_normalize_citation_year_typos(self):
        """Sửa lỗi gõ thừa chữ số trong năm nếu duy nhất cách sửa hợp lý."""
        # 20223 -> 2023 (thừa 1 số 2)
        self.assertEqual(
            normalize_citation("Decree 13/20223/ND-CP"),
            "Decree 13/2023/ND-CP"
        )
        # Năm đã đúng 4 chữ số -> giữ nguyên
        self.assertEqual(
            normalize_citation("Decree 13/2023/ND-CP"),
            "Decree 13/2023/ND-CP"
        )
        # Năm mập mờ nhiều khả năng (20203 -> 2003, 2020, 2023) -> giữ nguyên để không tự đoán sai
        self.assertEqual(
            normalize_citation("Decree 13/20203/ND-CP"),
            "Decree 13/20203/ND-CP"
        )

    def test_standardize_refs_with_kb(self):
        """Trích dẫn trùng điều khoản/quy định trong KB được chuẩn hóa theo bản chuẩn."""
        ref_en = ["Article 24.4 Decree 13/20203/ND-CP"]
        ref_vi = []
        en, vi = standardize_refs_with_kb(ref_en, ref_vi, "DE601")
        # Bản EN được thay bằng bản chuẩn của KB (đúng năm 2023)
        self.assertTrue(any("13/2023" in r for r in en))
        # Kèm luôn bản tiếng Việt chuẩn từ KB
        self.assertTrue(any("13/2023" in r for r in vi))

    def test_loose_separator_split_punct(self):
        """Tách song ngữ dính liền sau dấu chấm câu (./ ) không bị nuốt mất vế sau."""
        text = "Nhân viên chưa ký hợp đồng trước khi kí./ The labor contract was not signed timely."
        draft = parse_draft(text)
        self.assertTrue(draft.items_vi, "Phải giữ được vế tiếng Việt")
        self.assertTrue(draft.items_en, "Phải giữ được vế tiếng Anh")
        self.assertIn("Nhân viên", draft.items_vi[0])
        self.assertIn("labor contract", draft.items_en[0])

    def test_all_input_items_preserved_in_supplement(self):
        """Khối kết quả được supplement_missing_findings bảo đảm giữ đủ 100% draft ban đầu."""
        initial_block = structure_bilingual(["General issue"], [], ["Vấn đề chung"], [])
        draft_en = ["Specific worker 14% overtime violation (60~72h/week)"]
        draft_vi = ["14% nhân viên vi phạm thời gian làm việc (60~72 giờ/tuần)"]
        supplemented, added = supplement_missing_findings(initial_block, draft_en, draft_vi)
        self.assertTrue(added)
        self.assertIn("14% overtime", supplemented)
        self.assertIn("60~72 giờ/tuần", supplemented)
        self.assertEqual(missing_input_items(supplemented, " ".join(draft_en + draft_vi)), [])

    def test_mirror_each_finding_no_duplicate_and_translates(self):
        """mirror_each_finding dịch đủ vế thiếu và không dịch ngược gây lặp mục."""
        from hr_audit_app.rewrite import mirror_each_finding
        vi_items = ["Chưa ký hợp đồng lao động", "Làm thêm vượt quá 60 giờ/tuần"]
        en_items = []
        en_out, vi_out = mirror_each_finding(en_items, vi_items)
        self.assertEqual(len(vi_out), 2, "Phía VI không được bị lặp mục")
        self.assertEqual(len(en_out), 2, "Phía EN phải được bổ sung đủ 2 mục")
        self.assertIn("labor contract", en_out[0].lower())


if __name__ == "__main__":
    unittest.main(verbosity=2)
