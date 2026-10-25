# App Biên tập & Sinh Báo cáo HR Audit RBA (Samsung SEVT) v3.2

App tự sinh nội dung chuẩn cột **L** (Description of Supporting Evidence of
Violation) song ngữ Anh/Việt trong file Excel báo cáo HR audit theo chuẩn
**RBA Code of Conduct 8.0.1** và **pháp luật Việt Nam** (cấu trúc khối chuẩn,
trích dẫn pháp luật + luật Việt Nam), đồng thời tự động trích dẫn luật phù hợp
theo từng finding.

**Nguyên tắc cốt lõi: ĐẦY ĐỦ NỘI DUNG & THUẦN NGÔN NGỮ:**
- **Không rơi rụng nội dung:** Cột L của file đầu vào là *sự thật thực tế* (ground truth).
  Mọi ý, bằng chứng, dữ liệu số liệu (ví dụ % vi phạm, giờ làm, số ngày liên tục)
  của kiểm toán viên đều được giữ nguyên 100%, bổ sung bản dịch đối ứng, và trình bày
  vào cấu trúc khối chuẩn 4 header.
- **Không tự bịa finding:** Ô L của file đầu vào để trống thì khối kết quả để trống
  phần findings (chỉ gồm 4 header chuẩn và trích dẫn pháp luật tương ứng từ KB).
- **Chuẩn hóa & Sửa lỗi trích dẫn:**
  - Sửa lỗi gõ phím lặp thừa chữ số trong năm văn bản (vd `13/20223/ND-CP` → `13/2023/ND-CP`,
    `20119` → `2019`, `20225` → `2025`). Nếu mập mờ nhiều khả năng thì giữ nguyên văn.
  - Tự động thay thế bằng trích dẫn chuẩn từ KB (kèm đủ bản dịch đối ứng) khi KTV viện dẫn
    cùng một quy định/điều khoản đã có trong cơ sở dữ liệu pháp luật.
- **Bảo đảm thuần ngôn ngữ (không lẫn Anh/Việt):** MỌI dòng của cột L (6-57) đều
  được trình bày theo **đúng mẫu báo cáo đã duyệt** — khối 4 header in đậm, phần
  `AUDIT FINDINGS (ENGLISH)` + `Legal references:` **chỉ tiếng Anh**, phần
  `NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)` + `Căn cứ pháp lý:` **chỉ tiếng Việt**.
- **Hậu kiểm 2 lớp tự động khi Lưu / CLI:**
  1. `language_mixing_report()`: Báo cáo các ô còn lẫn ngôn ngữ hoặc sai cấu trúc.
  2. `completeness_report()`: Báo cáo nếu phát hiện bất kỳ nội dung nào của file đầu
     vào bị thiếu sau khi biên tập.
- Dòng có **Final Significance (cột D) = Conformance** giữ nguyên nội dung đánh giá
  của file đầu vào (không ghi đè bằng curated), chỉ chuẩn hóa lại cách trình bày đúng khối mẫu.

## Cấu trúc thư mục

```
2. APP Báo cáo checklist/
├── Chay_App.bat                   # ★ NHẤP ĐÚP ĐỂ MỞ APP (tự kiểm tra Python & thư viện)
├── run_app.py                     # Chạy app giao diện (GUI)
├── generate_report.py             # Chạy bằng dòng lệnh (CLI) - xử lý nhanh
├── validate_output.py             # Công cụ so sánh kết quả với file mẫu
├── requirements.txt               # Thư viện cần cài (openpyxl)
├── hr_audit_app/
│   ├── gui.py                     # Giao diện tkinter
│   ├── cli.py                     # Xử lý batch bằng dòng lệnh
│   ├── report_engine.py           # Đọc/ghi Excel, giữ nguyên định dạng & công thức
│   │                              #   + tự sửa file vendor (custom.xml hỏng)
│   ├── rewrite.py                 # Động cơ biên tập, cấu trúc khối song ngữ EN/VI
│   ├── knowledge_base.py          # KB: RBA 8.0.1 (A1-A6, A.M, D1-D6, D.M, E1-E4)
│   │                              #      + Luật Lao động 2019, NĐ 145/2020...
│   └── curated_findings.json      # Nội dung finding đã kiểm chuẩn
├── tests/
│   └── test_pipeline.py           # 49 test: KB, rewrite, engine, regression CLI, GUI
├── archive/                       # Script/txt trung gian dùng khi xây KB (tham khảo)
├── Copy of 2026.03.11_...xlsx     # File ĐẦU VÀO
└── 2026.08.26_...xlsx             # File ĐẦU RA mẫu (đã duyệt - dùng kiểm chứng)
```

## Cài đặt

Cần Python 3.10+ (tkinter có sẵn trên Windows). Cài thư viện:

```
pip install -r requirements.txt
```

## Cách 1 - Dùng giao diện (khuyên dùng)

**Nhanh nhất:** nhấp đúp vào file **`Chay_App.bat`** — app tự kiểm tra Python,
tự cài thư viện nếu thiếu và mở giao diện mà không hiện cửa sổ console đen.

Hoặc chạy bằng lệnh:

```
python run_app.py
```

1. **Chọn file đầu vào** → `Copy of 2026.03.11_...xlsx`, bấm **Mở báo cáo** (Ctrl+O).
   Bảng 52 tiêu chí hiện ra; dòng vi phạm tô màu hồng. Ngay sau khi mở, app tự
   sinh nội dung chuẩn và hiển thị **"Xem trước ô L"** theo cấu trúc song ngữ
   (phần tiếng Anh chỉ tiếng Anh, phần tiếng Việt chỉ tiếng Việt) — khung chỉ đọc.
2. Sửa **Thông tin đợt audit** (ngày, loại audit CAV, trưởng đoàn...).
3. Chọn từng dòng trong bảng Checklist để xem nội dung cột L tương ứng
   trong khung xem trước.
4. **Auto-generate**: trình bày lại cột L của **cả 52 dòng** theo đúng khối
   song ngữ mẫu (EN thuần Anh / VI thuần Việt). Dòng vi phạm dùng curated nếu có
   (thiếu thì giữ draft + bổ sung trích dẫn pháp luật); dòng **Conformance giữ
   nguyên nội dung và đánh giá** của file đầu vào, chỉ được chuẩn hóa cách trình
   bày (không thêm trích dẫn cho phần đã có nội dung).
5. **Chuẩn hóa dòng trống**: đưa các ô **N/A** trống về placeholder chuẩn.
6. **Lưu báo cáo** (Ctrl+S). Khi lưu, app tự rà lại định dạng (hàng trống EN/VI,
   4 header in đậm) và **hậu kiểm lẫn ngôn ngữ**: nếu còn mục lẫn Anh/Việt thì
   hiện cảnh báo kèm danh sách dòng cần rà lại.

## Cách 2 - Dòng lệnh (xử lý nhanh, tái lập được)

```
python generate_report.py ^
    --in "Copy of 2026.03.11_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
    --out "2026.08.26_L10PO5_YOUSUNG_HR audit result report in 2H.2026.xlsx" ^
    --audit-date 2026.08.26 ^
    --audit-type "Corrective Action Verification" ^
    --team-leader "KHUẤT MẠNH HOÀI" ^
    --team-member "NGUYỄN TIẾN THỊNH (SEVT)" ^
    --closing-date 2026-08-26 ^
    --interview-count 11 ^
    --interview-score 0.894 ^
    --profile-year 2025
```

Tùy chọn thêm:
- `--set-rating AL101=Minor` — ghi đè đánh giá (dùng được nhiều lần)
- `--no-curation` — không dùng nội dung curated, chỉ bổ sung trích dẫn KB
- `--no-kb-refs` — không tự thêm trích dẫn vào draft
- `--keep-profile` — giữ dữ liệu kỳ cũ trong Company Profile

## Logic sinh nội dung cột L

| Tình huống | Hành vi |
|---|---|
| Dòng vi phạm có nội dung kiểm chuẩn | Ghi nguyên văn đã duyệt (kèm rating nếu curated chỉ định, vd AL101 Major→Minor) |
| Dòng vi phạm có draft tiếng Việt | Giữ draft, **bổ sung trích dẫn** RBA 8.0.1 + luật VN tương ứng, dịch bù phần còn thiếu ngôn ngữ |
| Dòng vi phạm trống | Sinh khối chuẩn rỗng kèm trích dẫn KB |
| Dòng Conformance có draft thật | **Giữ nguyên nội dung + đánh giá**, chỉ TRÌNH BÀY LẠI theo khối mẫu (tách EN/VI, không thêm trích dẫn cho phần đã có nội dung) |
| Dòng Conformance trống | Khối rỗng + trích dẫn KB (đúng cách trình bày của file mẫu) |
| Dòng N/A trống (nút "Chuẩn hóa dòng trống") | Chuẩn hóa về placeholder chuẩn `Audit findings: / 1: / 2: / 3: / Legal references:` |
| Mọi dòng lẫn Anh/Việt (Vieglish) | Dịch lại cho thuần ngôn ngữ của khối; không đạt thì loại khỏi khối sai ngôn ngữ và báo trong hậu kiểm |

Khối chuẩn đầu ra:

```
AUDIT FINDINGS (ENGLISH)
1. <finding tiếng Anh chuẩn RBA>
Legal references:
1. RBA Code of Conduct, Version 8.0.1, Section A1: Prohibition of Forced Labor
2. Labor Code No. 45/2019/QH14, Article 8.1 (prohibition of forced labor)

NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)
1. <nội dung tiếng Việt>
Căn cứ pháp lý:
1. Bộ Quy tắc Ứng xử RBA phiên bản 8.0.1, Mục A1: Cấm lao động cưỡng bức
2. Bộ luật Lao động năm 2019, Điều 8 khoản 1 (...)
```

> ✍️ **Định dạng rich text trong cột L:** khi ghi vào Excel, các dòng header khối
> `AUDIT FINDINGS (ENGLISH)`, `Legal references:`, `NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)`,
> `Căn cứ pháp lý:` được **in đậm** (rich text), phần findings và trích dẫn giữ chữ
> thường — đúng định dạng file mẫu đã duyệt. Với placeholder ô trống, chỉ
> `Legal references:` được in đậm (khớp chuẩn báo cáo). Khung **"Xem trước ô L"**
> trong app cũng **in đậm 4 header** này, và phần tiếng Anh được phân cách với
> phần tiếng Việt bằng **1 hàng trống**.

## Cơ sở kiến thức trích dẫn (knowledge_base.py)

Mỗi Audit ID được ánh xạ tới section RBA 8.0.1 chính xác (đối chiếu tài liệu
`3. RBA_V8 Code Standard Document.pdf`):

- AL101-103 → **A1** Prohibition of Forced Labor
- AL201-203 → **A2** Young Workers
- AL301-303 → **A3** Working Hours
- AL401-403 → **A4** Wages and Benefits
- AL501-504 → **A5** Non-Discrimination / Non-Harassment / Humane Treatment
- AL601-603 → **A6** Freedom of Association and Collective Bargaining
- AM101-403 → **A.M.1.1 - A.M.4.3** Labor Management System
- DE101-601 → **D1-D6** Ethics (Business Integrity... Privacy)
- DM101-403 → **D.M.1.1 - D.M.4.3** Ethics Management System
- ES101-404 → **E1-E4** Supply Chain Management

Pháp luật Việt Nam theo chủ đề: Bộ luật Lao động 2019 (Đ.8, 14, 20-22, 94, 102,
104, 107, 111-113, 118, 125, 143-149, 183), NĐ 145/2020/NĐ-CP, NĐ 12/2022/NĐ-CP,
TT 10/2020/TT-BLDTBXH, NĐ 13/2023 & NĐ 356/2025 (dữ liệu cá nhân), Luật Công đoàn
2012, Luật Tín ngưỡng tôn giáo 2016, Luật Người khuyết tật 2010, Luật Cạnh tranh
2018, Luật PCTN 2018, Hiến pháp 2013.

## Bảo toàn định dạng Excel

App đọc/ghi bằng openpyxl nên **giữ nguyên**: font/màu/viền/căn lề, merged cells,
độ rộng cột, data validation, **toàn bộ công thức** (RBA Score, VLOOKUP điểm số,
ngày audit liên kết Company Profile) và **external links** tới bảng Table.
Cột L được ghi dưới dạng **rich text** (header khối in đậm) như file mẫu đã duyệt.

### Tương thích file từ nhà máy/đối tác khác

App **mở được cả các file do phần mềm khác tạo ra** (VD: file INTERFLEX báo cáo có
`docProps/custom.xml` thiếu thuộc tính `name`) bằng cơ chế **tự sửa chữa** — không
cần người dùng làm gì thêm. Nếu sau khi mở, app hiện thông báo lỗi, hãy mở file
`app_error.log` trong thư mục dự án và gửi nội dung cho người quản trị để khắc phục.

## Kiểm chứng

### 1. Bộ test tự động (68 test)

```python
python -m unittest discover -s tests -v
```

Bộ test bao gồm:

| Nhóm | Nội dung |
|---|---|
| `TestKnowledgeBase` | Đủ 52 Audit ID, ánh xạ đúng dòng 6-57, mọi mục có trích dẫn RBA + luật VN |
| `TestRewrite` | Parse draft, tách song ngữ VI/EN, nhận diện placeholder, build_content (curated/draft trống/placeholder). **Lọc nhãn phụ rỗng `Findings:`** trong draft nhà máy; **tự dịch bù trích dẫn thiếu ngôn ngữ** (EN↔VI, dự phòng bằng trích dẫn KB) |
| `TestMixedLanguagePurification` | **Tách dòng trộn lẫn "EN / VI"** (cả hai thứ tự), loại bản "Vieglish" (Việt lẫn từ Anh) khi có biến thể thuần tương đương, dịch đoạn ngoặc () còn sót ngôn ngữ kia trong trích dẫn — gồm cả case thực tế dài 2 câu |
 | `TestTranslator` | **Tự động dịch finding thiếu ngôn ngữ** (`translator.py`): ưu tiên **Google Translate online** (deep-translator, có cache `translation_cache.json`), không có mạng tự chuyển từ điển ngoại tuyến 300+ cặp chuyên ngành; quy tắc tháng/năm/khoảng số/%; VI-only draft → đủ cả EN và VI |
| `TestTranslatorOnline` | Dịch online: kết quả được **cache** (gọi lại không hỏi mạng), **fallback offline** khi lỗi mạng |
| `TestRichText` | `to_rich_text()`: header khối song ngữ in đậm; placeholder chỉ đậm `Legal references:` |
| `TestReportEngine` | Đọc metadata 52 dòng, validate rating, chuẩn hóa dòng trống N/A (idempotent), **regenerate trình bày lại cả 52 dòng** theo khối mẫu; dòng Conformance **giữ nguyên cột D và nội dung draft thuần ngôn ngữ** |
| `TestReportEngine` | Lỗi mở file: file giả (không phải .xlsx) báo thân thiện; thiếu sheet Checklist; **file vendor có custom.xml hỏng tự sửa chữa được** |
| `TestRichTextInGenerated` | Rich text: file CLI sinh ra có đủ 4 header đậm ở **mọi dòng 6-57**; **mọi run mang rPr hoàn chỉnh**; **không dòng nào còn lẫn Anh/Việt hoặc sai khối mẫu** |
| `TestLanguagePurity` | Bộ nhận diện thuần ngôn ngữ (`has_vietnamese_words` / `has_english_words` — bỏ qua mã văn bản luật `NĐ-CP` và tên riêng Viết Hoa); `enforce_language_purity` loại bản lẫn không sửa được khỏi khối sai ngôn ngữ; `scan_language_mixing` phát hiện Vieglish + ô sai định dạng; **mở file cũ còn lẫn Anh/Việt rồi repair → hết lẫn** |
| `TestCliPipeline` | **Regression**: CLI sinh file → so từng ô với file mẫu đã duyệt = 0 diff (chỉ bỏ qua ô mẫu còn lẫn ngôn ngữ đã được app sửa) |
| `TestGuiSmoke` | GUI khởi tạo, nạp báo cáo 52 dòng, đóng sạch sẽ |

### 2. So sánh thủ công với file mẫu

File sinh ra được so từng ô với file đầu ra mẫu đã duyệt:

```
python generate_report.py --in "Copy of 2026.03.11_..." --out test_output.xlsx ^
    --audit-date 2026.08.26 --audit-type "Corrective Action Verification" ^
    --team-leader "KHUẤT MẠNH HOÀI" --team-member "NGUYỄN TIẾN THỊNH (SEVT)" ^
    --closing-date 2026-08-26 --interview-count 11 --interview-score 0.894 --profile-year 2025
python validate_output.py test_output.xlsx
# TOTAL DIFFS: 0  ✔ TRÙNG KHỚP 100%
# (chỉ bỏ qua các ô cột L của FILE MẪU còn lẫn Anh/Việt - app đã sửa lại cho
#  thuần ngôn ngữ; script in rõ số ô như vậy)
```

### 3. Hậu kiểm lẫn ngôn ngữ trên file sinh ra

CLI tự in kết quả hậu kiểm sau mỗi lần sinh báo cáo:

```
  ✔ Kiểm tra ngôn ngữ: 0 mục lẫn Anh/Việt, 52/52 dòng đúng khối mẫu
```

Nếu còn mục lẫn, app liệt kê từng dòng kèm loại lỗi (`EN còn tiếng Việt`,
`VI còn tiếng Anh`, `không đúng khối mẫu`). GUI hiện cảnh báo tương tự khi lưu.
Có thể quét thủ công bằng `_scan_mixed.py <file.xlsx>`.

## Lưu ý

- Phần dịch EN cho finding mới (chưa có trong curated) cần kiểm toán viên hoàn
  thiện — app hỗ trợ: cấu trúc khối chuẩn, tách song ngữ "VI / EN" có sẵn trong
  draft, tự điền trích dẫn luật.
- Nên xác nhận lại hiệu lực/số hiệu văn bản mới nhất (vd NĐ 356/2025/NĐ-CP) trước
  khi phát hành báo cáo chính thức.
- Thư mục `archive/` chứa các script trích xuất/đối chiếu (extract*, grep*, verify*)
  đã dùng để xây dựng KB pháp luật từ file PDF gốc — chỉ mang tính tham khảo,
  không cần thiết khi chạy app.
