# -*- coding: utf-8 -*-
"""
GUI app sinh báo cáo HR audit RBA (tkinter).

Cách dùng:
    python run_app.py

Luồng làm việc:
1. Chọn file Excel đầu vào (Copy of ...xlsx), bấm "Mở báo cáo"
2. App tự sinh nội dung tiêu chuẩn (curated + trích dẫn KB) và hiển thị
   ngay "Xem trước ô L" theo cấu trúc song ngữ chuẩn (EN/VI)
3. Chọn từng dòng trong bảng Checklist để xem kết quả tương ứng
4. Lưu file đầu ra
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from hr_audit_app.knowledge_base import APP_VERSION, NON_CONFORMANCE_RATINGS  # noqa: E402
from hr_audit_app.report_engine import Report  # noqa: E402
from hr_audit_app.rewrite import (  # noqa: E402
    HEADER_EN,
    HEADER_VI,
    PLACEHOLDER,
    REFS_EN,
    REFS_VI,
    ensure_section_gap,
    is_placeholder,
    load_curated,
)

APP_TITLE = f"HR Audit Report Editor - RBA 8.0.1 - v{APP_VERSION}"
AUDIT_TYPES = ["Initial", "Corrective Action Verification", "Follow-up", "Special Audit"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x820")
        self.minsize(1100, 700)

        self.report = None
        self.in_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.curated = load_curated()
        self.current_row = None

        self._build_ui()
        self._bind_keys()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        top = ttk.Frame(self, padding=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="File đầu vào:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.in_path).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="Chọn...", command=self.choose_input).grid(row=0, column=2)
        ttk.Button(top, text="Mở báo cáo", command=self.load_report).grid(row=0, column=3, padx=(8, 0))

        ttk.Label(top, text="File đầu ra:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(top, textvariable=self.out_path).grid(row=1, column=1, sticky="ew", padx=4, pady=(4, 0))
        ttk.Button(top, text="Chọn...", command=self.choose_output).grid(row=1, column=2, pady=(4, 0))
        ttk.Button(top, text="Lưu báo cáo", command=self.save_report).grid(row=1, column=3, padx=(8, 0), pady=(4, 0))

        main = ttk.Frame(self)
        main.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # ---- Cột trái: metadata + bảng ----
        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        left.rowconfigure(1, weight=1)

        meta = ttk.LabelFrame(left, text="Thông tin đợt audit", padding=6)
        meta.grid(row=0, column=0, sticky="ew")
        self._build_meta(meta)

        tree_frame = ttk.LabelFrame(left, text="Checklist (52 tiêu chí)", padding=4)
        tree_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self.tree = ttk.Treeview(tree_frame, columns=("id", "subcat", "rating"),
                                 show="headings", height=22)
        self.tree.heading("id", text="Audit ID")
        self.tree.heading("subcat", text="Tiêu chí")
        self.tree.heading("rating", text="Đánh giá")
        self.tree.column("id", width=70, anchor="center")
        self.tree.column("subcat", width=190)
        self.tree.column("rating", width=90, anchor="center")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        # ---- Cột phải: thông tin dòng + preview kết quả ----
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        info = ttk.Frame(right)
        info.grid(row=0, column=0, sticky="ew")
        info.columnconfigure(1, weight=1)
        ttk.Label(info, text="Dòng đang xem:").grid(row=0, column=0, sticky="w")
        self.lbl_id = ttk.Label(info, text="-", font=("", 10, "bold"))
        self.lbl_id.grid(row=0, column=1, sticky="w")
        self.lbl_rating = ttk.Label(info, text="", font=("", 10))
        self.lbl_rating.grid(row=0, column=2, sticky="e", padx=(16, 0))

        prev_frame = ttk.LabelFrame(
            right, text="Xem trước ô L (kết quả ghi vào Excel)", padding=4)
        prev_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        prev_frame.columnconfigure(0, weight=1)
        prev_frame.rowconfigure(0, weight=1)
        self.txt_preview = tk.Text(prev_frame, wrap="word", height=22,
                                   font=("Consolas", 10))
        self.txt_preview.tag_configure("bold", font=("Consolas", 10, "bold"))
        self.txt_preview.grid(row=0, column=0, sticky="nsew")
        pv = ttk.Scrollbar(prev_frame, orient="vertical", command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=pv.set)
        pv.grid(row=0, column=1, sticky="ns")
        self.txt_preview.configure(state="disabled")  # chỉ đọc

        action = ttk.Frame(right)
        action.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action, text="Auto-generate dòng vi phạm",
                   command=self.auto_generate).pack(side="left", padx=(0, 6))
        ttk.Button(action, text="Chuẩn hóa dòng trống",
                   command=self.normalize_rows).pack(side="left", padx=(0, 6))
        ttk.Button(action, text="Lưu báo cáo",
                   command=self.save_report).pack(side="right")

        self.status = tk.StringVar(value="Mở file đầu vào để bắt đầu.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w",
                  padding=(6, 2)).grid(row=2, column=0, sticky="ew")

    def _build_meta(self, meta):
        self.meta_vars = {
            "audit_date": tk.StringVar(value=""),
            "audit_type": tk.StringVar(value="Corrective Action Verification"),
            "team_leader": tk.StringVar(value=""),
            "team_member": tk.StringVar(value=""),
            "closing_date": tk.StringVar(value=""),
            "interview_count": tk.StringVar(value=""),
            "interview_score": tk.StringVar(value=""),
            "profile_year": tk.StringVar(value=""),
        }
        rows = [
            ("Ngày audit (D1):", "audit_date"),
            ("Loại audit:", "audit_type"),
            ("Trưởng đoàn:", "team_leader"),
            ("Thành viên:", "team_member"),
            ("Ngày closing (YYYY-MM-DD):", "closing_date"),
            ("Số Worker interview:", "interview_count"),
            ("Interview Score:", "interview_score"),
            ("Profile year:", "profile_year"),
        ]
        for i, (label, key) in enumerate(rows):
            ttk.Label(meta, text=label).grid(row=i, column=0, sticky="w", pady=1)
            if key == "audit_type":
                w = ttk.Combobox(meta, textvariable=self.meta_vars[key],
                                 values=AUDIT_TYPES, state="readonly", width=28)
            else:
                w = ttk.Entry(meta, textvariable=self.meta_vars[key], width=30)
            w.grid(row=i, column=1, sticky="ew", padx=4)

        self.chk_clear = tk.BooleanVar(value=True)
        ttk.Checkbutton(meta, text="Reset dữ liệu Company Profile kỳ cũ",
                        variable=self.chk_clear).grid(
            row=len(rows), column=0, columnspan=2, sticky="w")
        self.chk_curated = tk.BooleanVar(value=True)
        ttk.Checkbutton(meta, text="Dùng nội dung đã kiểm chuẩn (curated)",
                        variable=self.chk_curated).grid(
            row=len(rows) + 1, column=0, columnspan=2, sticky="w")
        self.chk_kbrefs = tk.BooleanVar(value=True)
        ttk.Checkbutton(meta, text="Tự động bổ sung trích dẫn pháp luật",
                        variable=self.chk_kbrefs).grid(
            row=len(rows) + 2, column=0, columnspan=2, sticky="w")

    def _bind_keys(self):
        self.bind("<Control-o>", lambda e: self.load_report())
        self.bind("<Control-s>", lambda e: self.save_report())

    # ------------------------------------------------------------------
    # File
    # ------------------------------------------------------------------
    def choose_input(self):
        p = filedialog.askopenfilename(
            title="Chọn file audit đầu vào",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if p:
            self.in_path.set(p)
            if not self.out_path.get():
                base, name = os.path.split(p)
                out = os.path.join(base, name.replace("Copy of ", ""))
                # Tránh ghi đè file nguồn khi tên không có tiền tố "Copy of "
                if os.path.abspath(out) == os.path.abspath(p):
                    root, ext = os.path.splitext(name)
                    out = os.path.join(base, root + "_edited" + ext)
                self.out_path.set(out)

    def choose_output(self):
        p = filedialog.asksaveasfilename(
            title="Chọn file đầu ra", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if p:
            self.out_path.set(p)

    def load_report(self):
        p = self.in_path.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showwarning(APP_TITLE, "Hãy chọn file đầu vào hợp lệ.")
            return
        try:
            if self.report:
                self.report.close()
            self.report = Report(p)
        except ValueError as e:
            # Lỗi đã được chuyển thành thông báo thân thiện (file sai định dạng...)
            self._log_load_error(p, e)
            messagebox.showerror(APP_TITLE, f"Không mở được file:\n\n{e}")
            return
        except Exception as e:
            self._log_load_error(p, e)
            messagebox.showerror(
                APP_TITLE,
                f"Không mở được file:\n{e}\n\n"
                "Chi tiết lỗi đã ghi vào file app_error.log (trong thư mục dự án).\n"
                "Mẹo: đảm bảo file là .xlsx thật (không đổi đuôi từ .xls/.txt), "
                "không tải dở dang, và không mở cùng lúc trong Excel.")
            return
        self._fill_tree()
        self._fill_meta_from_report()
        # Tự sinh nội dung tiêu chuẩn và hiển thị preview ngay
        self.auto_generate()
        self.normalize_rows()
        self._select_first_violation()
        self.status.set(f"Đã mở: {p}")

    def _log_load_error(self, path, exc):
        """Ghi chi tiết lỗi mở file vào app_error.log để dễ chẩn đoán."""
        import datetime
        import traceback
        try:
            with open(
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                             "app_error.log"),
                "a", encoding="utf-8",
            ) as f:
                f.write("=" * 60 + "\n")
                f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] File: {path}\n")
                f.write(f"Tồn tại: {os.path.exists(path)}  "
                        f"Dung lượng: {os.path.getsize(path) if os.path.exists(path) else 'N/A'}\n")
                f.write(f"Lỗi: {type(exc).__name__}: {exc}\n")
                f.write(traceback.format_exc() + "\n")
        except Exception:
            pass  # không để việc ghi log làm hỏng luồng chính

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.report.rows():
            tags = ("nonconf",) if r["rating"] in NON_CONFORMANCE_RATINGS else ()
            self.tree.insert("", "end", iid=str(r["row"]),
                             values=(r["audit_id"], r["sub_category"], r["rating"]),
                             tags=tags)
        self.tree.tag_configure("nonconf", background="#ffe6e6")

    def _fill_meta_from_report(self):
        gi = self.report.wb["General Information"]
        cl = self.report.wb["Checklist"]
        cp = self.report.wb["Company Profile"]
        self.meta_vars["audit_date"].set(cl.cell(row=1, column=4).value or "")
        self.meta_vars["audit_type"].set(gi.cell(row=8, column=4).value or "Initial")
        self.meta_vars["team_leader"].set(gi.cell(row=6, column=4).value or "")
        self.meta_vars["team_member"].set(gi.cell(row=7, column=4).value or "")
        d11 = gi.cell(row=11, column=4).value
        self.meta_vars["closing_date"].set(
            d11.strftime("%Y-%m-%d") if hasattr(d11, "strftime") else "")
        ic = gi.cell(row=31, column=4).value
        self.meta_vars["interview_count"].set(str(ic) if ic is not None else "")
        iscore = cp.cell(row=11, column=6).value
        self.meta_vars["interview_score"].set(str(iscore) if iscore is not None else "")
    # ------------------------------------------------------------------
    # Chọn dòng & preview kết quả
    # ------------------------------------------------------------------
    def on_row_select(self, _event=None):
        sel = self.tree.selection()
        if not sel or not self.report:
            return
        row = int(sel[0])
        self.current_row = row
        audit_id = self.report.audit_id(row)
        rating = self.report.rating(row) or ""
        self.lbl_id.config(text=f"{audit_id} (dòng {row})")
        self.lbl_rating.config(text=f"Đánh giá: {rating or '—'}")
        self._show_preview(row)

    def _show_preview(self, row):
        evidence = (self.report.evidence(row) or "").strip()
        if not evidence or is_placeholder(evidence):
            self._set_preview(
                "(Dòng chưa có nội dung - sẽ ghi placeholder chuẩn khi lưu.)\n\n"
                + PLACEHOLDER)
        else:
            self._set_preview(ensure_section_gap(evidence))

    def _set_preview(self, text):
        self.txt_preview.configure(state="normal")
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text or "")
        # In đậm 4 dòng header khối trong preview (giống định dạng ô L khi lưu Excel)
        self.txt_preview.tag_remove("bold", "1.0", "end")
        for ln in (HEADER_EN, REFS_EN, HEADER_VI, REFS_VI):
            idx = "1.0"
            while True:
                idx = self.txt_preview.search(ln, idx, "end", exact=True)
                if not idx:
                    break
                self.txt_preview.tag_add("bold", idx, f"{idx}+{len(ln)}c")
                idx = f"{idx}+{len(ln)}c"
        self.txt_preview.configure(state="disabled")

    def _select_row(self, row):
        if row is not None and str(row) in self.tree.get_children():
            self.tree.selection_set(str(row))
            self.tree.focus(str(row))
            self.tree.see(str(row))
            self.on_row_select()

    def _select_first_violation(self):
        """Chọn dòng vi phạm đầu tiên (Minor/Major/Priority), nếu không có thì dòng đầu."""
        children = self.tree.get_children()
        if not children:
            return
        target = None
        for iid in children:
            values = self.tree.item(iid, "values") or []
            if len(values) >= 3 and values[2] in NON_CONFORMANCE_RATINGS:
                target = iid
                break
        if target is None:
            target = children[0]
        self.tree.selection_set(target)
        self.tree.focus(target)
        self.tree.see(target)
        self.on_row_select()

    # ------------------------------------------------------------------
    # Auto-generate / chuẩn hóa / lưu
    # ------------------------------------------------------------------
    def auto_generate(self):
        if not self.report:
            return
        touched = self.report.regenerate_findings(
            curated=self.curated if self.chk_curated.get() else None,
            use_curated=self.chk_curated.get(),
            add_kb_refs=self.chk_kbrefs.get(),
        )
        self._fill_tree()
        if touched:
            self._select_row(touched[0])
        else:
            self._select_first_violation()
        self.status.set(f"Auto-generate hoàn tất: trình bày lại {len(touched)}/52 dòng "
                        f"cột L theo khối song ngữ mẫu (EN thuần Anh / VI thuần Việt); "
                        f"dòng Conformance giữ nguyên nội dung và đánh giá của file đầu vào.")

    def normalize_rows(self):
        if not self.report:
            return
        changed = self.report.normalize_empty_rows()
        self._fill_tree()
        if self.current_row is not None:
            self._select_row(self.current_row)
        self.status.set(f"Đã chuẩn hóa {len(changed)} dòng N/A trống về placeholder chuẩn.")

    def save_report(self):
        if not self.report:
            messagebox.showwarning(APP_TITLE, "Chưa mở báo cáo nào.")
            return
        out = self.out_path.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, "Hãy chọn nơi lưu file đầu ra.")
            return
        # Rà soát định dạng các ô đã có khối song ngữ trước khi lưu: đảm bảo
        # đúng 1 hàng trống ngăn cách EN/VI và cả 4 header khối đều in đậm
        # (sửa kể cả nội dung do phiên bản app cũ sinh ra, không đổi văn bản)
        repaired = self.report.repair_bilingual_formatting()
        mixing = self.report.language_mixing_report()
        missing = self.report.completeness_report()
        self._apply_meta()
        try:
            self.report.save(out)
        except PermissionError:
            messagebox.showerror(
                APP_TITLE,
                "Không ghi được file (có thể đang mở trong Excel).\nĐóng file rồi thử lại.")
            return
        if mixing or missing:
            warn_parts = []
            if mixing:
                rows_mix = ", ".join(f"L{r}" for r, _k, _c in mixing[:10])
                warn_parts.append(
                    f"- Còn {len(mixing)} mục lẫn Anh/Việt hoặc chưa đúng mẫu tại: {rows_mix}")
            if missing:
                rows_miss = ", ".join(f"L{r}" for r, _c in missing[:10])
                warn_parts.append(
                    f"- Còn {len(missing)} nội dung của file đầu vào bị thiếu tại: {rows_miss}")
            messagebox.showwarning(
                APP_TITLE,
                f"Đã lưu báo cáo:\n{out}\n\nLưu ý:\n" + "\n".join(warn_parts))
            self.status.set(f"Đã lưu báo cáo: {out} - cần rà soát lại một số ô")
            return
        if repaired:
            self.status.set(
                f"Đã lưu báo cáo: {out} (đã sửa định dạng {len(repaired)} dòng)")
        else:
            self.status.set(f"Đã lưu báo cáo: {out}")
        messagebox.showinfo(
            APP_TITLE,
            f"Đã lưu báo cáo:\n{out}\n\n"
            f"Kiểm tra ngôn ngữ cột L: 0 mục lẫn Anh/Việt (52/52 dòng đúng khối mẫu).")

    def _apply_meta(self):
        import datetime
        m = self.meta_vars
        if m["audit_date"].get().strip():
            self.report.set_checklist_date(m["audit_date"].get().strip())
        closing = None
        if m["closing_date"].get().strip():
            try:
                closing = datetime.datetime.strptime(m["closing_date"].get().strip(), "%Y-%m-%d")
            except ValueError:
                pass
        ic = None
        if m["interview_count"].get().strip().isdigit():
            ic = int(m["interview_count"].get().strip())
        score = None
        try:
            score = (float(m["interview_score"].get().strip())
                     if m["interview_score"].get().strip() else None)
        except ValueError:
            score = None
        self.report.set_general_info(
            audit_type=m["audit_type"].get().strip() or None,
            team_leader=m["team_leader"].get().strip() or None,
            team_member=m["team_member"].get().strip() or None,
            closing_date=closing,
            worker_interview_count=ic,
        )
        year = m["profile_year"].get().strip() or None
        self.report.set_company_profile(year=year, interview_score=score,
                                        clear_old=self.chk_clear.get())


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
