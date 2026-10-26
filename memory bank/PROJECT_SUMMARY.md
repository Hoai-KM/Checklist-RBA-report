# Memory Bank - App Báo cáo Checklist HR Audit RBA

## Thông tin dự án
- **Tên:** App Biên tập & Sinh Báo cáo HR Audit RBA (Samsung SEVT) v3.2
- **Mục đích:** Tự động sinh nội dung cột L (Description of Supporting Evidence of Violation) song ngữ Anh/Việt trong file Excel báo cáo HR audit theo chuẩn RBA Code of Conduct 8.0.1 và pháp luật Việt Nam
- **Ngày tạo memory bank:** 25/10/2026
- **Remote Git:** https://github.com/Hoai-KM/Checklist-RBA-report.git

## Cấu trúc thư mục sau khi dọn dẹp

```
2. APP Báo cáo checklist/
├── hr_audit_app/              # Mã nguồn chính ứng dụng
│   ├── gui.py                 # Giao diện tkinter
│   ├── cli.py                 # Xử lý batch dòng lệnh
│   ├── report_engine.py       # Đọc/ghi Excel, giữ định dạng & công thức
│   ├── rewrite.py             # Động cơ biên tập, khối song ngữ EN/VI
│   ├── knowledge_base.py      # KB: RBA 8.0.1 + Luật Lao động 2019, NĐ 145/2020
│   ├── curated_findings.json  # Nội dung finding đã kiểm chuẩn
│   ├── translation_cache.json # Cache dịch thuật
│   ├── elaborations.py        # Chi tiết nội dung
│   ├── elaborations_full.py
│   ├── elab_a.py, elab_b.py, elab_c.py
│   └── __init__.py
├── tests/                     # Bộ kiểm thử (49 test)
│   └── test_pipeline.py
├── memory bank/               # Thư mục lưu trữ thông tin dự án (này)
│   └── PROJECT_SUMMARY.md
├── generate_report.py         # Entry point CLI
├── run_app.py                 # Entry point GUI
├── validate_output.py         # Công cụ so sánh kết quả
├── Chay_App_Moi.bat           # Script chạy app
├── Sinh_Bao_Cao.bat           # Script sinh báo cáo
├── requirements.txt           # openpyxl, deep-translator
├── README.md                  # Hướng dẫn sử dụng
├── *.xlsx                     # File dữ liệu đầu vào / đầu ra mẫu
└── *.pdf                      # Tài liệu tham khảo pháp luật (RBA V8, Luật Lao động, NĐ 145/2020)
```

## File rác đã xóa
- `__pycache__/` (gốc, hr_audit_app, tests) — cache Python bytecode
- `archive/` — script trích xuất/đối chiếu KB (extract*, grep*, verify*), file .txt trung gian, file rỗng (chỉ mang tính tham khảo, không cần chạy app)
- `_q2.py` — file test tạm (`print('q2-ok')`)

## Sửa lỗi ES403 bị lặp và trộn lẫn nội dung tiêu chí khác (25/10/2026)
- **Lỗi 1 (trùng lặp):** `enforce_language_purity()` và `build_content()` chỉ `dedupe_items` cho `refs` (trích dẫn) nhưng KHÔNG cho `findings` (nội dung vi phạm). Khi `mirror_each_finding()` dịch bù hoặc `supplement_pools()` chèn lại nội dung bị rơi, cùng nội dung ES403 xuất hiện nhiều lần.
- **Lỗi 2 (trộn lẫn tiêu chí):** Khi `curated_covers_draft()` trả về True, `supplement_pools()` thêm TẤT CẢ draft fragment không có trong curated — bao gồm nội dung của tiêu chí KHÁC (vd AL301/AL302 về giờ làm việc bị trộn vào ô ES403).
- **Sửa tại `hr_audit_app/rewrite.py`:**
  1. Dòng 1164: `enforce_language_purity()` trả về `dedupe_items(en_f), dedupe_items(vi_f), dedupe_items(en_r), dedupe_items(vi_r)`.
  2. Dòng ~1667: `build_content()` thêm `dedupe_items()` cho `en_f`, `vi_f` sau `mirror_each_finding()`.
  3. Dòng ~1466: `supplement_missing_findings()` thêm `dedupe_items()` cho `en_f2`, `vi_f2` sau `supplement_pools()`.
  4. Mới: hàm `_mentions_other_audit(text, current_audit_id)` phát hiện đoạn văn nhắc mã tiêu chí khác (vd AL301 trong ô ES403).
  5. Dòng ~1630: `build_content()` lọc `draft_en`/`draft_vi` bằng `_mentions_other_audit` trước khi `supplement_pools()` — chặn nội dung chéo giữa các tiêu chí.
- **Kết quả:** Tất cả 79 test (234 subtests) vẫn pass.

## Cách chạy
```bash
pip install -r requirements.txt
python run_app.py          # Giao diện
python generate_report.py  # Dòng lệnh
```

## Git remote
`https://github.com/Hoai-KM/Checklist-RBA-report.git`
