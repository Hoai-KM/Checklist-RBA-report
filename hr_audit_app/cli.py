# -*- coding: utf-8 -*-
"""
CLI sinh báo cáo HR audit từ file đầu vào.

Ví dụ:
    python generate_report.py ^
        --in "Copy of 2026.03.11_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
        --out "2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
        --audit-date 2026.08.26 --audit-type "Corrective Action Verification" ^
        --team-leader "KHUẤT MẠNH HOÀI" --team-member "NGUYỄN TIẾN THỊNH (SEVT)" ^
        --closing-date 2026-08-26 --interview-count 11 --interview-score 0.894 ^
        --profile-year 2025
"""
import argparse
import datetime
import os
import sys

# Đảm bảo in tiếng Việt được trên mọi console Windows (cp1252/cp650)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from report_engine import ALLOWED_RATINGS, Report  # noqa: E402
from knowledge_base import ROW_ITEMS, APP_VERSION  # noqa: E402
from rewrite import load_curated  # noqa: E402

DEFAULTS = {
    "audit_date": None,
    "audit_type": None,
    "team_leader": None,
    "team_member": None,
    "closing_date": None,
    "interview_count": None,
    "interview_score": None,
    "profile_year": None,
}


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Sinh báo cáo HR audit RBA từ file checklist")
    p.add_argument("--in", dest="inp", required=True, help="File Excel đầu vào")
    p.add_argument("--out", dest="out", required=True, help="File Excel đầu ra")
    p.add_argument("--audit-date", help="Ngày audit hiển thị (Checklist!D1), vd 2026.08.26")
    p.add_argument("--audit-type", help="Loại audit (General Information!D8)")
    p.add_argument("--team-leader", help="Trưởng đoàn audit (GI!D6)")
    p.add_argument("--team-member", help="Thành viên audit (GI!D7)")
    p.add_argument("--closing-date", help="Ngày closing audit (GI!D11), YYYY-MM-DD")
    p.add_argument("--interview-count", type=int, help="Số Worker interview (GI!D31)")
    p.add_argument("--interview-score", type=float, help="Interview Score (Company Profile!F11)")
    p.add_argument("--profile-year", help="Năm profile (Company Profile!F2/F3), vd 2025")
    p.add_argument("--keep-profile", action="store_true",
                   help="Giữ dữ liệu kỳ cũ trong Company Profile (mặc định sẽ reset)")
    p.add_argument("--no-curation", action="store_true",
                   help="Không dùng nội dung curated, chỉ thêm trích dẫn KB vào draft")
    p.add_argument("--no-kb-refs", action="store_true",
                   help="Không bổ sung trích dẫn pháp luật vào các draft vi phạm")
    p.add_argument("--set-rating", action="append", default=[],
                   metavar="ID=RATING",
                   help="Ghi đè đánh giá của một tiêu chí, vd --set-rating AL101=Minor (dùng nhiều lần)")
    return p.parse_args(argv)


def run(args):
    if not os.path.exists(args.inp):
        raise SystemExit(f"Không tìm thấy file đầu vào: {args.inp}")

    print(f"HR Audit App v{APP_VERSION}")
    curated = load_curated()
    use_curated = not args.no_curation

    # Ghi đè rating theo tham số: --set-rating AL101=Minor
    apply_ratings = {}
    row_by_id = {ROW_ITEMS[r]: r for r in ROW_ITEMS}
    for spec in args.set_rating:
        if "=" not in spec:
            raise SystemExit(f"--set-rating sai cú pháp: {spec} (dạng ID=RATING)")
        audit_id, rating = spec.split("=", 1)
        audit_id, rating = audit_id.strip().upper(), rating.strip()
        if audit_id not in row_by_id:
            raise SystemExit(f"Audit ID không tồn tại: {audit_id}")
        if rating not in ALLOWED_RATINGS:
            raise SystemExit(f"Rating không hợp lệ: {rating}")
        apply_ratings[row_by_id[audit_id]] = rating

    report = Report(args.inp)

    # 1) Metadata
    if args.audit_date:
        report.set_checklist_date(args.audit_date)
    closing = None
    if args.closing_date:
        closing = datetime.datetime.strptime(args.closing_date, "%Y-%m-%d")
    report.set_general_info(
        audit_type=args.audit_type,
        team_leader=args.team_leader,
        team_member=args.team_member,
        closing_date=closing,
        worker_interview_count=args.interview_count,
    )
    report.set_company_profile(
        year=args.profile_year,
        interview_score=args.interview_score,
        clear_old=not args.keep_profile,
    )

    # 2) Nội dung cột L cho TẤT CẢ 52 dòng theo khối song ngữ mẫu
    #    (dòng Conformance giữ nguyên nội dung draft + đánh giá của file đầu vào)
    touched = report.regenerate_findings(
        curated=curated,
        use_curated=use_curated,
        add_kb_refs=not args.no_kb_refs,
        apply_ratings=apply_ratings or None,
    )

    # 3) Chuẩn hóa các dòng trống/placeholder
    normalized = report.normalize_empty_rows()

    # 3b) Rà soát lại định dạng các ô đã có khối song ngữ (hàng trống EN/VI +
    #     bôi đậm cả 4 header, tách ngôn ngữ) để sửa cả nội dung của phiên bản
    #     app cũ (kể cả dòng Conformance)
    repaired = report.repair_bilingual_formatting()

    # 3c) Hậu kiểm cột L: không còn mục lẫn Anh/Việt, mọi dòng đúng khối mẫu
    mixing = report.language_mixing_report()

    # 3d) Hậu kiểm ĐẦY ĐỦ: bảo đảm không rơi mất bất kỳ nội dung nào từ file đầu vào
    missing = report.completeness_report()

    # 4) Lưu
    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    report.save(args.out)
    report.close()

    print(f"Đã sinh báo cáo: {args.out}")
    print(f"  - Số dòng đã trình bày lại cột L theo khối mẫu: {len(touched)}/52")
    print(f"  - Dòng trống đã chuẩn hóa (chỉ N/A): {len(normalized)} dòng")
    print(f"  - Dòng đã sửa lại định dạng (hàng trống + in đậm header): {len(repaired)} dòng")
    print(f"  - Nội dung curated sử dụng: {sorted(curated)}")
    if mixing:
        print(f"  ! CÒN {len(mixing)} mục cần rà lại (lẫn ngôn ngữ / sai định dạng):")
        for row, kind, content in mixing[:20]:
            label = {"en": "EN còn tiếng Việt", "vi": "VI còn tiếng Anh",
                     "format": "không đúng khối mẫu"}.get(kind, kind)
            print(f"      - L{row} [{label}]: {content[:100]}")
    else:
        print("  ✔ Kiểm tra ngôn ngữ: 0 mục lẫn Anh/Việt, 52/52 dòng đúng khối mẫu")
    if missing:
        print(f"  ! CÒN {len(missing)} nội dung file đầu vào bị thiếu sau khi biên tập:")
        for row, item_text in missing[:20]:
            print(f"      - L{row}: {item_text[:100]}")
    else:
        print("  ✔ Kiểm tra đầy đủ: Giữ đủ 100% nội dung đánh giá từ file đầu vào")


def main(argv=None):
    args = parse_args(argv)
    run(args)


if __name__ == "__main__":
    main()
