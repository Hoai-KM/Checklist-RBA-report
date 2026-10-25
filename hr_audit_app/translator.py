# -*- coding: utf-8 -*-
"""
Tự động dịch finding kiểm toán HR giữa tiếng Việt và tiếng Anh.

Ưu tiên dịch online qua Google Translate (thư viện deep-translator) để đạt độ
chính xác cao; khi không có Internet (hoặc đặt HR_AUDIT_TRANSLATOR_ONLINE=0) tự
chuyển sang từ điển chuyên ngành audit ngoại tuyến. Kết quả online được lưu cache
(translation_cache.json) để chạy lại nhanh và ổn định. Kết quả mang tính DRAFT —
ô L kèm ghi chú "(Bản dịch tự động - kiểm toán viên rà soát)" để KTV rà lại.
"""
import hashlib
import json
import os
import re
import socket

MONTHS_EN = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

# ------------------------------------------------------------------ từ điển ---
# Cặp (tiếng Việt, English). Cụm dài được xử lý trước cụm ngắn (tự sắp xếp).
PAIRS = [
    # --- cụm chuyên ngành dài (VI -> EN) ---
    ("hợp đồng không xác định thời hạn", "labor contract with indefinite term"),
    ("không xác định thời hạn", "indefinite term"),
    ("xác định thời hạn", "fixed-term"),
    ("thỏa ước lao động tập thể", "collective bargaining agreement"),
    ("thỏa thuận lao động tập thể", "collective bargaining agreement"),
    ("giấy phép môi trường", "environmental permit"),
    ("giấy tờ tùy thân", "personal documents"),
    ("giấy khai sinh", "birth certificate"),
    ("lương tối thiểu vùng", "regional minimum wage"),
    ("lương tối thiểu", "minimum wage"),
    ("bảng chấm công", "timekeeping records"),
    ("bảng tính lương", "payroll sheet"),
    ("nội quy lao động", "internal labor regulations"),
    ("hợp đồng lao động", "labor contract"),
    ("điều kiện lao động", "working conditions"),
    ("an toàn lao động", "occupational safety"),
    ("vệ sinh lao động", "occupational hygiene"),
    ("tai nạn lao động", "work-related accident"),
    ("bảo hộ lao động", "personal protective equipment"),
    ("phòng cháy chữa cháy", "fire prevention and fighting"),
    ("pccc", "fire prevention and fighting"),
    ("lối thoát hiểm", "emergency exit"),
    ("báo cháy", "fire alarm"),
    ("khám sức khỏe", "health examination"),
    ("nghỉ thai sản", "maternity leave"),
    ("nghỉ không lương", "unpaid leave"),
    ("ngày nghỉ tuần", "weekly rest day"),
    ("nghỉ tuần", "weekly rest"),
    ("nghỉ bù", "compensatory rest"),
    ("ngày nghỉ", "rest day"),
    ("nghỉ lễ", "public holiday"),
    ("ngày lễ", "public holiday"),
    ("nghỉ phép", "annual leave"),
    ("phép năm", "annual leave"),
    ("nghỉ ốm", "sick leave"),
    ("làm việc ban đêm", "night work"),
    ("làm đêm", "night work"),
    ("ca đêm", "night shift"),
    ("tăng ca", "overtime"),
    ("thời gian làm thêm", "overtime"),
    ("khảo sát", "survey"),
    ("định kỳ", "periodically"),
    ("lần", "times"),
    ("làm thêm từ", "worked overtime from"),
    ("làm thêm giờ", "overtime"),
    ("làm thêm", "overtime"),
    ("thời giờ làm việc", "working hours"),
    ("thời gian làm việc", "working hours"),
    ("giờ làm việc", "working hours"),
    ("giờ công", "working hours"),
    ("môi trường làm việc", "working environment"),
    ("quan hệ lao động", "labor relations"),
    ("lao động trẻ em", "child labor"),
    ("chưa thành niên", "juvenile workers"),
    ("lao động cưỡng bức", "forced labor"),
    ("quấy rối tình dục", "sexual harassment"),
    ("phân biệt đối xử", "discrimination"),
    ("công đoàn cơ sở", "grassroots labor union"),
    ("kênh khiếu nại", "grievance mechanism"),
    ("báo cáo ẩn danh", "anonymous reporting"),
    ("bảo mật thông tin", "information confidentiality"),
    ("dữ liệu cá nhân", "personal data"),
    ("xung đột lợi ích", "conflict of interest"),
    ("đánh giá rủi ro", "risk assessment"),
    ("hồ sơ nhân sự", "personnel records"),
    ("hồ sơ bảng lương", "payroll records"),
    ("chất thải nguy hại", "hazardous waste"),
    ("hành động khắc phục", "corrective action"),
    ("tự đánh giá", "self-assessment"),
    ("trợ cấp thôi việc", "severance allowance"),
    ("chấm dứt hợp đồng", "contract termination"),
    ("đơn xin nghỉ việc", "resignation application"),
    ("thông báo nghỉ việc", "resignation notice"),
    ("chậm trả lương", "delayed wage payment"),
    ("bảo hiểm xã hội", "social insurance"),
    ("bảo hiểm y tế", "health insurance"),
    ("bảo hiểm thất nghiệp", "unemployment insurance"),
    ("người sử dụng lao động", "employer"),
    ("người lao động", "workers"),
    ("lực lượng lao động", "workforce"),
    ("theo quy định của pháp luật", "as required by law"),
    ("theo quy định", "as stipulated"),
    ("quy định cụ thể", "clearly stipulated"),
    ("chưa quy định", "has not stipulated"),
    ("thời gian báo trước", "notice period"),
    ("bổ sung thêm", "further supplement"),
    ("chưa được", "has not been"),
    ("không được phép", "must not"),
    ("không quá", "no more than"),
    ("chưa có", "does not have"),
    ("không có", "does not have"),
    ("không đầy đủ", "incomplete"),
    ("không đúng", "incorrect"),
    ("không nhất quán", "inconsistent"),
    ("không khớp", "mismatched"),
    ("cần được", "need to be"),
    ("được giải thích", "was explained"),
    ("bổ sung", "supplement"),
    ("quy định", "stipulated"),
    ("giờ/tháng", "hours per month"),
    ("giờ/tuần", "hours per week"),
    ("giờ/ngày", "hours per day"),
    ("giờ/năm", "hours per year"),
    ("mỗi tháng", "per month"),
    ("mỗi tuần", "per week"),
    ("mỗi ngày", "per day"),
    ("mỗi năm", "per year"),
    ("hàng năm", "annually"),
    ("hàng tháng", "monthly"),
    ("hàng tuần", "weekly"),
    ("hàng ngày", "daily"),
    ("tổng số", "total"),
    ("số lượng", "number of"),
    ("phần trăm", "percent"),
    ("liên tục", "consecutive"),
    ("rõ ràng", "clear"),
    ("hợp lệ", "valid"),
    ("hiệu lực", "validity"),
    ("hiện hành", "applicable"),
    ("pháp luật", "the law"),
    ("nghị định", "decree"),
    ("thông tư", "circular"),
    ("chính sách", "policy"),
    ("quy trình", "procedure"),
    ("hướng dẫn", "guideline"),
    ("chứng nhận", "certificate"),
    ("sao y", "certified copy"),
    ("đào tạo", "training"),
    ("huấn luyện", "training"),
    ("giấy phép", "permit"),
    ("đã hết hạn", "expired"),
    ("hết hạn", "expired"),
    ("cập nhật", "updated"),
    ("lưu giữ", "kept"),
    ("nước thải", "wastewater"),
    ("khí thải", "air emissions"),
    ("chất thải", "waste"),
    ("hóa chất", "chemicals"),
    ("hoá chất", "chemicals"),
    ("sự cố", "incident"),
    ("khẩn cấp", "emergency"),
    ("sơ tán", "evacuation"),
    ("nhà cung cấp", "supplier"),
    ("thầu phụ", "subcontractor"),
    ("nhà thầu", "contractor"),
    ("khắc phục", "corrective"),
    ("bằng chứng", "evidence"),
    ("chứng cứ", "evidence"),
    ("kiểm toán", "audit"),
    ("báo cáo", "report"),
    ("thôi việc", "termination"),
    ("nghỉ việc", "resignation"),
    ("xét duyệt", "approval"),
    ("xác nhận", "confirmation"),
    ("tiếp nhận", "receipt"),
    ("thông tin", "information"),
    ("giới tính", "gender"),
    ("khấu trừ", "deduction"),
    ("trừ lương", "wage deduction"),
    ("phạt tiền", "monetary penalty"),
    ("phụ cấp", "allowance"),
    ("chuyên cần", "attendance"),
    ("thâm niên", "seniority"),
    ("trả lương", "wage payment"),
    ("đúng hạn", "on time"),
    ("tính lương", "payroll calculation"),
    ("lương thử việc", "probation wage"),
    ("thử việc", "probation"),
    ("tiền lương", "wages"),
    ("tiền thưởng", "bonus"),
    ("bảo hiểm", "insurance"),
    ("nhân sự", "employees"),
    ("công nhân", "workers"),
    ("nhân viên", "employees"),
    ("lao động", "labor"),
    ("công đoàn", "labor union"),
    ("khiếu nại", "grievance"),
    ("ẩn danh", "anonymous"),
    ("bảo mật", "confidentiality"),
    ("hối lộ", "bribery"),
    ("quà tặng", "gifts"),
    ("minh bạch", "transparency"),
    ("cưỡng bức", "forced"),
    ("quấy rối", "harassment"),
    ("phân biệt", "discrimination"),
    ("lương", "wages"),
    ("thưởng", "bonus"),
    ("mức trần", "cap"),
    ("vượt xa", "far exceed"),
    ("vượt quá", "exceed"),
    ("vượt", "exceed"),
    ("thừa", "exceeding"),
    ("trung bình", "on average"),
    ("tổng", "total"),
    ("thời hạn", "duration"),
    ("thời gian", "time"),
    ("báo trước", "advance notice"),
    ("đối với", "regarding"),
    ("sai lệch", "discrepancy"),
    ("khác biệt", "discrepancy"),
    ("sai", "incorrect"),
    ("nhầm", "mistaken"),
    ("đầy đủ", "complete"),
    ("bắt buộc", "mandatory"),
    ("cần thiết", "necessary"),
    ("yêu cầu", "requirement"),
    ("khuyến nghị", "recommendation"),
    ("kiến nghị", "recommendation"),
    ("đề xuất", "recommendation"),
    ("lưu ý", "note"),
    ("thiếu", "missing"),
    ("chưa", "has not"),
    ("phải", "must"),
    ("nên", "should"),
    ("đề nghị", "request"),
    ("quá", "more than"),
    ("kiểm tra", "review"),
    ("kiểm soát", "control"),
    ("rà soát", "screening"),
    ("phát hiện", "detected"),
    ("ghi nhận", "recorded"),
    ("theo dõi", "monitored"),
    ("thực hiện", "carry out"),
    ("vận hành", "operate"),
    ("áp dụng", "apply"),
    ("tổ chức", "organize"),
    ("cung cấp", "provide"),
    ("đảm bảo", "ensure"),
    ("bảo đảm", "ensure"),
    ("tuân thủ", "comply with"),
    ("chấp hành", "comply with"),
    ("vi phạm", "violation of"),
    ("sai phạm", "violations"),
    ("nghiêm trọng", "severe"),
    ("hệ thống", "systemic"),
    ("nghiêm túc", "serious"),
    ("về", "regarding"),
    ("qua", "through"),
    ("nghỉ", "rest"),
    ("làm việc", "working"),
    ("viết tay", "handwritten"),
    ("nội dung", "content"),
    ("biểu mẫu", "forms"),
    ("dữ liệu", "data"),
    ("an toàn", "safety"),
    ("môi trường", "environment"),
    ("cộng đồng", "community"),
    ("trẻ em", "children"),
    ("y tế", "medical"),
    ("cháy", "fire"),
    ("kỷ luật", "disciplinary"),
    ("sa thải", "dismissal"),
    ("chấm dứt", "terminate"),
    ("ký", "sign"),
    ("phạt", "penalty"),
    ("mẫu", "form"),
    ("hồ sơ", "records"),
    ("tài liệu", "documents"),
    ("danh sách", "list"),
    ("phần", "section"),
    ("tự", "self"),
    ("được", ""),
    ("để", "to"),
    ("trong", "in"),
    ("ngoài", "outside"),
    ("từ", "from"),
    ("đến", "to"),
    ("và", "and"),
    ("hoặc", "or"),
    ("của", "of"),
    ("với", "with"),
    ("cho", "for"),
    ("theo", "per"),
    ("mỗi", "each"),
    ("tất cả", "all"),
    ("mọi", "all"),
    ("vẫn", "still"),
    ("cũng", "also"),
    ("nữa", "more"),
    ("không", "not"),
    ("có", "have"),
    ("là", "is"),
    ("hiện tại", "currently"),
    ("tối đa", "maximum"),
    ("tối thiểu", "minimum"),
    ("khoảng", "approximately"),
    ("dưới", "under"),
    ("trên", "over"),
    ("năm", "year"),
    ("tháng", "months"),
    ("tuần", "weeks"),
    ("ngày", "days"),
    ("đêm", "night"),
    ("giờ", "hours"),
    ("phút", "minutes"),
    ("tuổi", "years old"),
    ("quý", "quarter"),
    ("công ty", "the company"),
    ("nhà máy", "the factory"),
    ("xưởng", "workshop"),
    ("khách hàng", "customer"),
    ("sản phẩm", "products"),
    ("hàng hóa", "goods"),
    ("khu vực", "area"),
    ("nlđ", "workers"),
    ("hđlđ", "labor contract"),
    ("hđ", "contract"),
    ("đã", ""),
    ("cần để", "need to"),
    ("cần", "need to"),
    ("lại", "again"),
    ("đánh giá", "assessment"),
    ("cải tiến", "improvement"),
    ("bhxh", "social insurance"),
    ("bhyt", "health insurance"),
    # --- từ/cụm tiếng Việt hay bị máy dịch bỏ sót khiến câu trở thành
    #     "Vieglish" (Anh lẫn Việt) trong phần EN của báo cáo ---
    ("quản lý hiệu quả", "effectively managed"),
    ("phát sinh", "generating"),
    ("xảy ra", "occurring"),
    ("đang", "currently"),
    ("thường xuyên", "regularly"),
    ("yêu cầu khách hàng", "customer requirements"),
    ("công nhân viên", "employees"),
    # --- liên từ / chỉ dẫn câu phổ biến trong findings ---
    ("tuy nhiên", "however"),
    ("nhưng", "but"),
    ("do đó", "therefore"),
    ("ngoài ra", "in addition"),
    ("bên cạnh đó", "in addition"),
    ("trong quá trình", "during"),
    ("tại", "at"),
    # --- danh từ/cụm danh từ audit còn thiếu ---
    ("tác động", "impact"),
    ("ảnh hưởng", "influence"),
    ("các bên liên quan", "stakeholders"),
    ("bên liên quan", "stakeholders"),
    ("liên quan", "related"),
    ("đánh giá tác động", "impact assessment"),
    ("chưa rõ ràng", "is not clear"),
    ("biên bản", "minutes"),
    ("quan sát", "observed"),
    ("phỏng vấn", "interviewed"),
    ("trình bày", "presented"),
    ("cho biết", "stated"),
    ("xem xét", "reviewed"),
    ("sổ sách", "records"),
    # "việc" là danh tử tiếng Việt, bỏ khi dịch sang Anh (vd "việc đánh giá" -> "assessment")
    ("việc", ""),
    ("đã thực hiện", "conducted"),
    # --- cụm hay gặp trong findings về PPE / hợp đồng / đăng ký / thiết lập ---
    ("trang thiết bị bảo hộ lao động", "personal protective equipment"),
    ("trang thiết bị", "equipment"),
    ("cấp đầy đủ", "fully provided"),
    ("cấp", "provided"),
    ("chưa được ký kết", "has not been signed"),
    ("ký kết bằng văn bản", "signed in writing"),
    ("bằng văn bản", "in writing"),
    ("đăng ký tham gia", "registered for"),
    ("đăng ký", "registered"),
    ("thiết lập", "established"),
    ("chưa trả lương", "has not paid wages"),
    ("vượt quy định", "exceeding the stipulated"),
    ("người sử dụng lao động", "employer"),
    ("nsdlđ", "employer"),
    ("nsdld", "employer"),
    ("thu nhập cá nhân", "personal income"),
    ("thuế tncn", "personal income tax"),
    ("tncn", "personal income tax"),
    ("hợp đồng học việc", "apprenticeship contract"),
    ("học việc", "apprenticeship"),
    ("thực tập", "internship"),
    ("phép tồn", "outstanding annual leave"),
    ("nghỉ việc", "leave of absence"),
    ("tháng liền kề", "adjacent month"),
    ("tháng liên kề", "adjacent month"),
    ("vị trí lao động phổ thông", "general labor positions"),
    ("lao động phổ thông", "general labor"),
    ("chương trình học việc", "apprenticeship program"),
    ("quy trình", "procedure"),
    ("khuyết tật", "disability"),
    ("người khuyết tật", "persons with disabilities"),
    ("lao động khuyết tật", "disabled workers"),
    ("thỏa đáng", "satisfactorily"),
    ("giải đáp", "respond"),
    ("xử lý", "handle"),
    ("ý kiến", "comments"),
    ("đóng góp", "contributions"),
    ("khấu trừ", "deduct"),
    ("khấu trừ thuế", "tax deduction"),
    ("thời hạn báo trước", "notice period"),
    ("báo trước", "advance notice"),
    ("học nghề", "vocational training"),
    ("dạy nghề", "vocational education"),
    ("giáo dục nghề nghiệp", "vocational education"),
]

# Sắp xếp: cụm dài xử lý trước cụm ngắn (cả hai chiều).
# LƯU Ý: loại các cặp có phần tiếng Anh rỗng (VD ("được", "")) khỏi chiều EN->VI,
# nếu không pattern rỗng sẽ khớp mọi vị trí và phá hủy văn bản.
_VI2EN = sorted(PAIRS, key=lambda p: -max(len(p[0]), len(p[1])))
_EN2VI = sorted(((en, vi) for vi, en in PAIRS if en.strip()),
                key=lambda p: -max(len(p[0]), len(p[1])))

# Tập từ tiếng Việt KHÔNG dấu (dùng để nhận diện "Vieglish" — Anh lẫn Việt không dấu).
# Bao gồm: các từ đơn từ PAIRS + bổ sung các từ Việt không dấu phổ biến trong draft nhà máy.
_VI_WORDS_RAW = set()
for _vi, _en in PAIRS:
    for _w in _vi.lower().split():
        if _w.isalpha():
            _VI_WORDS_RAW.add(_w)
# Bổ sung từ Việt không dấu phổ biến (draft nhà máy dùng nhiều, chưa có trong PAIRS)
_VI_WORDS_RAW.update([
    "phép", "chính", "xác", "nay", "hết", "nhưng", "ban", "hành", "tác",
    "động", "xử", "lý", "quản", "hiệu", "quả", "dẫn", "điến", "lỗi", "họp",
    "xem", "xét", "mục", "tiêu", "khoáng", "sản", "xung", "đột", "công",
    "khai", "danh", "tính", "gộp", "hình", "thức", "kéo", "dài", "nâng",
    "chí", "clear", "của", "trong", "được", "không", "có", "và",
    "hoặc", "đã", "sẽ", "đang", "vẫn", "cũng", "nữa", "lại", "thì", "mà",
    "nên", "phải", "cần", "để", "cho", "với", "từ", "trên", "dưới", "giữa",
    "sau", "trước", "theo", "về", "ra", "vào", "lên", "xuống", "này", "kia",
    "đó", "đây", "ai", "gì", "nào", "bao", "nhiêu", "mấy", "sao", "thế",
    "như", "rằng", "là", "mỗi", "tất", "cả", "tổng", "số", "trung",
    "bình", "thời", "gian", "tháng", "ngày", "năm", "tuần", "giờ", "phút",
    "đêm", "sáng", "chiều", "tối", "hôm", "qua", "mai", "mốt",
    "tiếp", "vừa", "mới", "rồi", "xong", "dở", "sắp",
    # Bổ sung từ Việt không dấu phổ biến trong draft nhà máy
    "khong", "hop", "dong", "cong", "tinh", "loi", "nang",
    "thang", "ngay", "tuan", "gio", "phut", "dem",
    "chieu", "toi", "hom", "qua", "mai", "mot",
    "vua", "moi", "roi", "xong", "dang", "sap",
    "phong", "chay", "chua", "chuyen", "sinh", "hoat",
    "hopdong", "laodong", "baohiem", "xahoi", "yte",
    "khac", "nhu", "dac", "biet", "thong", "tin", "quan",
    "trong", "quoc", "te", "hop", "tac", "doi", "thong",
    "nhat", "trong", "so", "luong", "thuong", "phat",
    "tang", "giam", "thap", "cao", "lon", "nho", "gan", "xa",
    "moi", "cu", "nay", "do", "xanh", "den", "trang", "do",
    "tim", "vang", "cam", "hong", "xam", "nau", "re", "dat",
    "thu", "chi", "kho", "nuoc", "gio", "bay", "thuyen",
    "xe", "may", "dien", "tu", "ban", "ghe", "sofa",
    "bep", "nuoi", "trong", "cay", "hoa", "la", "canh",
    "rau", "qua", "thit", "ca", "trung", "sua", "banh",
    "com", "pho", "bun", "chao", "mi", "xoi", "banhmi",
])
VI_WORDS_NO_DIACRITICS = frozenset(_VI_WORDS_RAW)

# ----------------------------------------------------------- quy tắc tiền xử lý
_MONTH_LIST_VI_RE = re.compile(
    r"\btháng\s*((?:\d{1,2}\s*,\s*)+\d{1,2})(?:\s*/\s*(\d{4}))?", re.IGNORECASE)
_MONTH_ONE_VI_RE = re.compile(r"\btháng\s*(\d{1,2})(?:\s*/\s*(\d{4}))?", re.IGNORECASE)
_DUR_MONTHS_VI_RE = re.compile(
    r"\btrong\s+(\d+)\s+tháng\s+((?:\d{1,2}\s*,\s*)+\d{1,2})(?:\s*/\s*(\d{4}))?",
    re.IGNORECASE)
_PCT_VI_RE = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*%\s*(?:tổng\s+)?(?:số\s+)?"
    r"(?:NLĐ|NLD|người lao động|nguoi lao dong|nhân sự|lao động|workers)",
    re.IGNORECASE)
_DATE_VI_RE = re.compile(r"\bngày\s*(\d{1,2}/\d{1,2}/\d{2,4})\b", re.IGNORECASE)
_RANGE_VI_RE = re.compile(r"(\d+)\s*[~]\s*(\d+)")
_MONTH_EN_RE = re.compile(
    r"\b(" + "|".join(MONTHS_EN) + r")\b", re.IGNORECASE)
_BETWEEN_AND_EN_RE = re.compile(r"\bbetween\s+(\d+)\s+and\s+(\d+)", re.IGNORECASE)
_BETWEEN_TO_EN_RE = re.compile(r"\bbetween\s+(\d+)\s+to\s+(\d+)", re.IGNORECASE)
_RANGE_TO_EN_RE = re.compile(r"(\d+)\s+to\s+(\d+)")
_PCT_EN_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*%\s*of\s+(?:the\s+)?(?:total\s+)?"
    r"(?:workers|workforce|employees|labor force)", re.IGNORECASE)


def _month_names(list_str):
    out = []
    for n in re.split(r"\s*,\s*", list_str.strip()):
        i = int(n)
        out.append(MONTHS_EN[i - 1] if 1 <= i <= 12 else f"month {n}")
    return out


def _dur_months_repl(m):
    names = ", ".join(_month_names(m.group(2)))
    year = f" {m.group(3)}" if m.group(3) else ""
    return f"over {m.group(1)} months ({names}{year})"


def _month_list_repl(m):
    names = ", ".join(_month_names(m.group(1)))
    year = f" {m.group(2)}" if m.group(2) else ""
    return f"{names}{year}"


def _month_one_repl(m):
    i = int(m.group(1))
    name = MONTHS_EN[i - 1] if 1 <= i <= 12 else f"month {m.group(1)}"
    year = f" {m.group(2)}" if m.group(2) else ""
    return f"{name}{year}"


def _preprocess_vi_to_en(s):
    s = _DUR_MONTHS_VI_RE.sub(_dur_months_repl, s)
    s = _MONTH_LIST_VI_RE.sub(_month_list_repl, s)
    s = _MONTH_ONE_VI_RE.sub(_month_one_repl, s)
    s = _PCT_VI_RE.sub(r"\1% of workers", s)
    s = _DATE_VI_RE.sub(r"\1", s)
    s = _RANGE_VI_RE.sub(r"\1 to \2", s)
    s = re.sub(r"(\d+),(\d+)\s*%", r"\1.\2%", s)  # "39,8%" -> "39.8%"
    return s


def _preprocess_en_to_vi(s):
    s = _BETWEEN_AND_EN_RE.sub(r"\1 ~ \2", s)
    s = _BETWEEN_TO_EN_RE.sub(r"\1 ~ \2", s)
    s = _RANGE_TO_EN_RE.sub(r"\1 ~ \2", s)
    s = _PCT_EN_RE.sub(r"\1% tổng số NLĐ", s)
    s = re.sub(r"\ban average of\b", "on average", s, flags=re.IGNORECASE)
    s = _MONTH_EN_RE.sub(
        lambda m: f"tháng {MONTHS_EN.index(m.group(1).capitalize()) + 1}", s)
    return s


def _apply_pairs(s, pairs):
    for src, dst in pairs:
        s = re.sub(rf"\b{re.escape(src)}\b", dst, s, flags=re.IGNORECASE)
    return s


def _postprocess(s):
    s = re.sub(r"\s+([,.;:!?%)])", r"\1", s)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    s = re.sub(r"\s*\n\s*", "\n", s)
    s = re.sub(r"(^|\n)([a-zà-ỹ])", lambda m: m.group(1) + m.group(2).upper(), s)
    s = re.sub(r"([.!?]\s+)([a-zà-ỹ])", lambda m: m.group(1) + m.group(2).upper(), s)
    return s.strip()


# ------------------------------------------------------- dịch online & cache ---
_CACHE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "translation_cache.json")
_cache = {}
if os.path.exists(_CACHE_PATH):
    try:
        with open(_CACHE_PATH, encoding="utf-8") as f:
            _cache = json.load(f)
    except Exception:
        _cache = {}
_online_disabled = False  # lỗi mạng -> tắt online cho cả phiên, không chờ mạng nữa
socket.setdefaulttimeout(10)  # tránh treo app khi không có mạng


def _cache_key(text, target):
    return hashlib.sha1(f"{target}|{text}".encode("utf-8")).hexdigest()


_session = None


def _get_session():
    global _session
    if _session is None:
        try:
            import requests
            _session = requests.Session()
            _session.headers.update({
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/128.0.0.0 Safari/537.36"
                ),
                "Accept": "*/*",
            })
        except ImportError:
            _session = False
    return _session


def _translate_online(text, target):
    """
    Dịch online qua Google Translate:
    1. Ưu tiên endpoint Google Translate API (GTX) qua requests.Session:
       Nhanh, tin cậy, không bị lỗi 429 TooManyRequests như scraping web mobile.
    2. Dự phòng: deep-translator GoogleTranslator nếu endpoint GTX không phản hồi.
    """
    source = "vi" if target == "en" else "en"
    # 1. Thử endpoint Google Translate API (GTX)
    session = _get_session()
    if session:
        try:
            url = "https://translate.googleapis.com/translate_a/single"
            params = {
                "client": "gtx",
                "sl": source,
                "tl": target,
                "dt": "t",
                "q": text,
            }
            r = session.get(url, params=params, timeout=10)
            if r.status_code == 200:
                data = r.json()
                if data and data[0]:
                    res = "".join(part[0] for part in data[0] if part and part[0])
                    if res and res.strip():
                        return res.strip()
        except Exception:
            pass

    # 2. Dự phòng: thư viện deep-translator
    from deep_translator import GoogleTranslator
    tr = GoogleTranslator(source=source, target=target)
    return tr.translate(text)


def _default_online():
    """Mặc định dùng online; đặt HR_AUDIT_TRANSLATOR_ONLINE=0 để buộc offline."""
    return os.environ.get("HR_AUDIT_TRANSLATOR_ONLINE", "1") != "0"


def translate_text_offline(text, target):
    """Dịch bằng từ điển ngoại tuyến (dự phòng khi không có Internet)."""
    if target == "en":
        return _postprocess(_apply_pairs(_preprocess_vi_to_en(text), _VI2EN))
    if target == "vi":
        return _postprocess(_apply_pairs(_preprocess_en_to_vi(text), _EN2VI))
    raise ValueError(f"target phải là 'en' hoặc 'vi', nhận: {target!r}")


def translate_text(text, target, prefer_online=None):
    """
    Dịch một dòng finding sang 'en' hoặc 'vi'.

    prefer_online=None -> theo cài đặt môi trường (mặc định: online nếu được).
    Kết quả online được cache ra translation_cache.json; offline không dùng cache
    (để kết quả offline luôn xác định, phục vụ test).
    """
    global _online_disabled
    if not text or not text.strip():
        return text
    use_online = _default_online() if prefer_online is None else prefer_online
    if use_online and not _online_disabled:
        key = _cache_key(text, target)
        if key in _cache:
            return _cache[key]
        result = None
        try:
            result = _translate_online(text, target)
        except Exception:
            _online_disabled = True
        if result and result.strip():
            result = result.strip()
            _cache[key] = result
            try:
                with open(_CACHE_PATH, "w", encoding="utf-8") as f:
                    json.dump(_cache, f, ensure_ascii=False)
            except Exception:
                pass
            return result
    return translate_text_offline(text, target)


def translate_findings(items, target, prefer_online=None):
    """Dịch danh sách finding sang 'en' hoặc 'vi'."""
    return [translate_text(t, target, prefer_online) for t in items]
