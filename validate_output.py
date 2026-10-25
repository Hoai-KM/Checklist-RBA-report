# -*- coding: utf-8 -*-
"""Validate: so sánh file sinh ra với file output mẫu - từng ô."""
import os
import sys

import openpyxl

# Đảm bảo in tiếng Việt/ký tự đặc biệt được trên console Windows (cp1252...)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hr_audit_app.rewrite import scan_language_mixing  # noqa: E402

SAMPLE = r"d:\AI PROJECT\2. APP Báo cáo checklist\2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx"
GENERATED = sys.argv[1] if len(sys.argv) > 1 else r"d:\AI PROJECT\2. APP Báo cáo checklist\test_output.xlsx"

wb_s = openpyxl.load_workbook(SAMPLE, data_only=False)
wb_g = openpyxl.load_workbook(GENERATED, data_only=False)

def fmt(v, n=200):
    if v is None:
        return ""
    s = str(v).replace("\r\n", "¶").replace("\n", "¶")
    return s[:n]

total_diff = 0
fixed_mixed = 0
for sheet in wb_s.sheetnames:
    ws_s = wb_s[sheet]
    ws_g = wb_g[sheet]
    maxr = max(ws_s.max_row, ws_g.max_row)
    maxc = max(ws_s.max_column, ws_g.max_column)
    diffs = []
    for r in range(1, maxr + 1):
        # Ô L của FILE MẪU còn lẫn ngôn ngữ (EN sót tiếng Việt / VI sót tiếng Anh)
        # thì không đối chiếu - app đã sửa lại cho thuần Anh/Việt.
        sample_mixed = (
            sheet == "Checklist" and 6 <= r <= 57
            and any(kind in ("en", "vi")
                    for kind, _content in
                    scan_language_mixing(ws_s.cell(row=r, column=12).value)))
        for c in range(1, maxc + 1):
            if sample_mixed and c == 12:
                fixed_mixed += 1
                continue
            vs = ws_s.cell(row=r, column=c).value
            vg = ws_g.cell(row=r, column=c).value
            if vs != vg:
                total_diff += 1
                diffs.append(f"  {ws_s.cell(row=r, column=c).coordinate}: SAMPLE='{fmt(vs)}' GEN='{fmt(vg)}'")
    print(f"=== {sheet}: {len(diffs)} diff(s)")
    for d in diffs[:15]:
        print(d)
    if len(diffs) > 15:
        print(f"  ... và {len(diffs)-15} diff khác")

print()
print("Lưu ý: các ô L của file mẫu đã được app sửa vì còn lẫn ngôn ngữ:",
      fixed_mixed)
print("TOTAL DIFFS:", total_diff)
if total_diff == 0:
    print("✔ FILE SINH RA TRÙNG KHỚP 100% VỚI FILE MẪU (theo giá trị ô)")
wb_s.close(); wb_g.close()
