# -*- coding: utf-8 -*-
"""
Cơ sở kiến thức RBA 8.0.1 + pháp luật Việt Nam cho app biên tập báo cáo HR audit.

Nguồn tham chiếu:
- RBA VAP Standard V8.0.1 (3. RBA_V8 Code Standard Document.pdf):
    A. Labor: A1-A6; A.M Labor Management System (A.M.1.1 - A.M.4.3)
    D. Ethics: D1-D6; D.M Ethics Management System (D.M.1.1 - D.M.4.3)
    E Supply Chain Management: E1-E4
- Bộ luật Lao động 2019 (Luật số 45/2019/QH14)
- Nghị định 145/2020/NĐ-CP, 12/2022/NĐ-CP, 13/2023/NĐ-CP
- Thông tư 10/2020/TT-BLDTBXH
"""

LABOR_CODE = "Bộ luật Lao động năm 2019 (Luật số 45/2019/QH14)"
LABOR_CODE_URL = "https://thuvienphapluat.vn/van-ban/Lao-dong-Tieu-duong/Bo-luat-Cong-viec-2019-43075.aspx"
LABOR_CODE_EN = "Labor Code No. 45/2019/QH14"
RBA_VERSION = "RBA Code of Conduct, Version 8.0.1"
RBA_VERSION_VI = "Bộ Quy tắc Ứng xử RBA phiên bản 8.0.1"
RBA_VERSION_URL = "https://www.content.ra.org/rba/code-2022.pdf"

RATINGS = ["Conformance", "Minor", "Major", "Priority", "N/A"]
NON_CONFORMANCE_RATINGS = ["Minor", "Major", "Priority"]

# Phiên bản app - được ghi vào thuộc tính file đầu ra (Excel > File > Info > Tags)
# để kiểm tra chắc chắn file được sinh/lưu bởi bản code mới.
APP_VERSION = "3.2"

# ---------------------------------------------------------------------------
# Trích dẫn pháp luật Việt Nam theo chủ đề (chỉ trích dẫn điều luật TRỌNG TÂM NHẤT,
# không trích lan man, bảo đảm 100% chuẩn xác theo pháp luật hiện hành).
# ---------------------------------------------------------------------------
VN = {}

VN["forced_labor"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 8 khoản 1 (nghiêm cấm cưỡng bức lao động, bóc lột lao động)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 8.1 (prohibition of forced labor, exploitation)",
    ],
}

VN["id_documents"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 8 khoản 9 (nghiêm cấm giữ bản chính giấy tờ tùy thân của người lao động)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 8.9 (prohibition of retaining workers' original identity documents)",
    ],
}

VN["labor_contract"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 14 khoản 2 (giao kết hợp đồng lao động bằng văn bản)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 14.2 (written labor contract)",
    ],
}

VN["contract_type"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 20 (loại hợp đồng lao động, chuyển đổi loại hợp đồng)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 20 (types of labor contracts, contract conversion)",
    ],
}

VN["child_labor"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 143 (sử dụng lao động chưa thành niên, nghiêm cấm lao động trẻ em)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 143 (employment of minors, prohibition of child labor)",
    ],
}

VN["juvenile_protection"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 148 (thời giờ làm việc của người chưa thành niên: cấm làm thêm giờ, làm đêm)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 148 (working hours of juvenile workers: no overtime, no night work)",
    ],
}

VN["learner_protection"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 61 (học nghề, tập nghề tại doanh nghiệp)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 61 (vocational training and apprenticeship at enterprises)",
    ],
}

VN["working_hours"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 107 (làm thêm giờ không quá 40 giờ/tháng, 200 giờ/năm)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 107 (overtime: maximum 40 hours/month, 200 hours/year)",
    ],
}

VN["weekly_rest"] = {
    "vi": [f"{LABOR_CODE}, Điều 111 (nghỉ hằng tuần: ít nhất 24 giờ liên tục mỗi tuần)"],
    "en": [f"{LABOR_CODE_EN}, Article 111 (weekly rest: at least 24 consecutive hours per week)"],
}

VN["leave_holidays"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 113 (nghỉ hằng năm có lương)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 113 (annual paid leave)",
    ],
}

VN["wages"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 94 (nguyên tắc trả lương: trực tiếp, đầy đủ, đúng hạn)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 94 (wages shall be paid directly, in full and on time)",
    ],
}

VN["pay_slip"] = {
    "vi": [f"{LABOR_CODE}, Điều 95 khoản 3 (mỗi lần trả lương phải thông báo bảng kê trả lương / phiếu lương cho người lao động)"],
    "en": [f"{LABOR_CODE_EN}, Article 95.3 (pay slip / wage statement provided for each pay period)"],
}

VN["deductions"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 102 (khấu trừ tiền lương: không được quá 30% tiền lương hằng tháng)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 102 (wage deductions: not exceeding 30% of monthly wage)",
    ],
}

VN["harassment"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 8 khoản 2 (nghiêm cấm phân biệt đối xử, quấy rối tình dục tại nơi làm việc)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 8.2 (prohibition of discrimination and sexual harassment in the workplace)",
    ],
}

VN["discipline"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 125 (các hình thức xử lý kỷ luật lao động)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 125 (forms of labor discipline)",
    ],
}

VN["religion"] = {
    "vi": ["Luật Tín ngưỡng, tôn giáo 2016, Điều 6 (quyền tự do tín ngưỡng, tôn giáo)"],
    "en": ["Law on Belief and Religion 2016, Article 6 (freedom of belief and religion)"],
}

VN["disability"] = {
    "vi": ["Luật Người khuyết tật 2010, Điều 16 (quyền làm việc của người khuyết tật)"],
    "en": ["Law on Persons with Disabilities 2010, Article 16 (right to work of persons with disabilities)"],
}

VN["freedom_of_association"] = {
    "vi": ["Luật Công đoàn 2012, Điều 6 (quyền thành lập, gia nhập và hoạt động công đoàn)"],
    "en": ["Law on Trade Unions 2012, Article 6 (right to establish, join and operate trade unions)"],
}

VN["collective_bargaining"] = {
    "vi": [
        f"{LABOR_CODE}, Điều 73 (thương lượng tập thể, thỏa ước lao động tập thể)",
    ],
    "en": [
        f"{LABOR_CODE_EN}, Article 73 (collective bargaining, collective labor agreement)",
    ],
}

VN["peaceful_assembly"] = {
    "vi": ["Hiến pháp 2013, Điều 25 (quyền tự do hội họp, lập hội)"],
    "en": ["Constitution 2013, Article 25 (freedom of assembly and association)"],
}

VN["bribery"] = {
    "vi": ["Luật Phòng, chống tham nhũng 2018, Điều 7 (nghiêm cấm hối lộ, lạm dụng chức vụ)"],
    "en": ["Anti-Corruption Law 2018, Article 7 (prohibited acts: bribery, abuse of position)"],
}

VN["privacy"] = {
    "vi": [
        "Nghị định số 13/2023/NĐ-CP, Điều 24 (hồ sơ đánh giá tác động xử lý dữ liệu cá nhân)",
    ],
    "en": [
        "Decree No. 13/2023/ND-CP, Article 24 (Personal Data Processing Impact Assessment)",
    ],
}

VN["competition"] = {
    "vi": ["Luật Cạnh tranh 2018, Điều 6 (cấm các thỏa thuận hạn chế cạnh tranh)"],
    "en": ["Competition Law 2018, Article 6 (prohibition of anti-competitive agreements)"],
}

VN["none"] = {"vi": [], "en": []}


def rba_ref(section, topic_en, topic_vi):
    return {
        "en": f"{RBA_VERSION}, Section {section}: {topic_en}",
        "vi": f"{RBA_VERSION_VI}, Mục {section}: {topic_vi}",
    }
# ---------------------------------------------------------------------------
# KB cho từng Audit ID của checklist (52 mục, dòng 6-57 sheet Checklist)
# ---------------------------------------------------------------------------
_ITEMS_RAW = {
    "AL101": ("A1", "Prohibition of Forced Labor", "Cấm lao động cưỡng bức", "forced_labor"),
    "AL102": ("A1", "Prohibition of Forced Labor (written contract in native language)", "Cấm lao động cưỡng bức (hợp đồng tiếng mẹ đẻ)", "labor_contract"),
    "AL103": ("A1", "Prohibition of Forced Labor (no holding of original ID)", "Cấm lao động cưỡng bức (không giữ giấy tờ gốc)", "id_documents"),
    "AL201": ("A2", "Young Workers (prohibition of child labor)", "Lao động chưa thành niên (cấm lao động trẻ em)", "child_labor"),
    "AL202": ("A2", "Young Workers (juvenile worker protection)", "Lao động chưa thành niên (bảo vệ chưa thành niên)", "juvenile_protection"),
    "AL203": ("A2", "Young Workers (student worker / intern / apprentice)", "Lao động chưa thành niên (thực tập sinh / học nghề)", "learner_protection"),
    "AL301": ("A3", "Working Hours", "Thời giờ làm việc", "working_hours"),
    "AL302": ("A3", "Working Hours (one day off in seven)", "Thời giờ làm việc (nghỉ 1 ngày mỗi tuần)", "weekly_rest"),
    "AL303": ("A3", "Working Hours (breaks, holidays and leave)", "Thời giờ làm việc (nghỉ giữa giờ, lễ tết, phép năm)", "leave_holidays"),
    "AL401": ("A4", "Wages and Benefits (accurate and timely payment)", "Tiền lương và phúc lợi (trả lương chính xác, đúng hạn)", "wages"),
    "AL402": ("A4", "Wages and Benefits (pay slip)", "Tiền lương và phúc lợi (phiếu lương)", "pay_slip"),
    "AL403": ("A4", "Wages and Benefits (deductions)", "Tiền lương và phúc lợi (khấu trừ)", "deductions"),
    "AL501": ("A5", "Non-Discrimination / Non-Harassment / Humane Treatment", "Không phân biệt đối xử / Quấy rối / Đối xử nhân đạo", "harassment"),
    "AL502": ("A5", "Non-Discrimination / Humane Treatment (disciplinary actions)", "Không phân biệt đối xử / Đối xử nhân đạo (kỷ luật)", "discipline"),
    "AL503": ("A5", "Non-Discrimination / Humane Treatment (religious accommodation)", "Không phân biệt đối xử / Đối xử nhân đạo (tôn giáo)", "religion"),
    "AL504": ("A5", "Non-Discrimination / Humane Treatment (disability accommodation)", "Không phân biệt đối xử / Đối xử nhân đạo (người khuyết tật)", "disability"),
    "AL601": ("A6", "Freedom of Association and Collective Bargaining (trade union)", "Tự do hiệp hội (công đoàn)", "freedom_of_association"),
    "AL602": ("A6", "Freedom of Association and Collective Bargaining", "Tự do hiệp hội và thương lượng tập thể", "collective_bargaining"),
    "AL603": ("A6", "Freedom of Association and Collective Bargaining (peaceful assembly)", "Tự do hiệp hội (họp hành hòa bình)", "peaceful_assembly"),
    "AM101": ("A.M.1.1", "Labor Management System - compliance process", "Hệ thống quản lý lao động - quy trình tuân thủ", "none"),
    "AM102": ("A.M.1.2", "Labor Management System - due diligence", "Hệ thống quản lý lao động - thẩm định rủi ro", "none"),
    "AM201": ("A.M.2.1", "Labor Management System - responsibilities and authorities", "Hệ thống quản lý lao động - trách nhiệm, thẩm quyền", "none"),
    "AM202": ("A.M.2.2", "Labor Management System - policies and control process", "Hệ thống quản lý lao động - chính sách, kiểm soát", "weekly_rest"),
    "AM203": ("A.M.2.3", "Labor Management System - training", "Hệ thống quản lý lao động - đào tạo", "none"),
    "AM301": ("A.M.3.1", "Labor Management System - worker/stakeholder communication", "Hệ thống quản lý lao động - truyền thông", "none"),
    "AM302": ("A.M.3.2", "Labor Management System - worker grievance process", "Hệ thống quản lý lao động - khiếu nại của NLĐ", "none"),
    "AM401": ("A.M.4.1", "Labor Management System - performance review and improvement", "Hệ thống quản lý lao động - đánh giá, cải tiến", "none"),
    "AM402": ("A.M.4.2", "Labor Management System - self-assessment", "Hệ thống quản lý lao động - tự đánh giá", "none"),
    "AM403": ("A.M.4.3", "Labor Management System - corrective action", "Hệ thống quản lý lao động - hành động khắc phục", "none"),
}
_ITEMS_RAW.update({
    "DE101": ("D1", "Business Integrity and No Improper Advantage", "Liêm chính kinh doanh, không lợi thế bất chính", "bribery"),
    "DE201": ("D2", "Disclosure of Information (no record falsification)", "Công khai thông tin (cấm làm giả hồ sơ)", "none"),
    "DE301": ("D3", "Intellectual Property", "Sở hữu trí tuệ", "none"),
    "DE401": ("D4", "Fair Business, Advertising and Competition", "Kinh doanh công bằng, quảng cáo, cạnh tranh", "competition"),
    "DE501": ("D5", "Protection of Identity and Non-Retaliation", "Bảo vệ danh tính và không trả đũa", "none"),
    "DE601": ("D6", "Privacy", "Quyền riêng tư", "privacy"),
    "DM101": ("D.M.1.1", "Ethics Management System - compliance process", "Hệ thống quản lý đạo đức - quy trình tuân thủ", "none"),
    "DM102": ("D.M.1.2", "Ethics Management System - due diligence", "Hệ thống quản lý đạo đức - thẩm định rủi ro", "none"),
    "DM201": ("D.M.2.1", "Ethics Management System - responsibilities and authorities", "Hệ thống quản lý đạo đức - trách nhiệm, thẩm quyền", "none"),
    "DM202": ("D.M.2.2", "Ethics Management System - policies and control process", "Hệ thống quản lý đạo đức - chính sách, kiểm soát", "none"),
    "DM203": ("D.M.2.3", "Ethics Management System - training", "Hệ thống quản lý đạo đức - đào tạo", "none"),
    "DM301": ("D.M.3.1", "Ethics Management System - communication", "Hệ thống quản lý đạo đức - truyền thông", "none"),
    "DM302": ("D.M.3.2", "Ethics Management System - grievance process", "Hệ thống quản lý đạo đức - khiếu nại", "none"),
    "DM401": ("D.M.4.1", "Ethics Management System - performance review and improvement", "Hệ thống quản lý đạo đức - đánh giá, cải tiến", "none"),
    "DM402": ("D.M.4.2", "Ethics Management System - self-assessment", "Hệ thống quản lý đạo đức - tự đánh giá", "none"),
    "DM403": ("D.M.4.3", "Ethics Management System - corrective action", "Hệ thống quản lý đạo đức - hành động khắc phục", "none"),
    "ES101": ("E1", "Supply Chain Management - Company Commitment", "Quản lý chuỗi cung ứng - cam kết công ty", "none"),
    "ES201": ("E2", "Supply Chain Management - Materials Restrictions", "Quản lý chuỗi cung ứng - hạn chế vật liệu", "none"),
    "ES301": ("E3", "Supply Chain Management - Responsible Sourcing of Minerals", "Quản lý chuỗi cung ứng - khoáng sản có trách nhiệm", "none"),
    "ES401": ("E4", "Supplier Responsibility (communication of Code requirements)", "Trách nhiệm nhà cung cấp (truyền thông yêu cầu)", "none"),
    "ES402": ("E4", "Supplier Responsibility (identification of major suppliers)", "Trách nhiệm nhà cung cấp (xác định nhà cung cấp chính)", "none"),
    "ES403": ("E4", "Supplier Responsibility (supplier implementation of the RBA Code)", "Trách nhiệm nhà cung cấp (triển khai Bộ quy tắc RBA)", "none"),
    "ES404": ("E4", "Supplier Responsibility (continuous improvement)", "Trách nhiệm nhà cung cấp (cải tiến liên tục)", "none"),
})

ITEMS = {}
for _id, (_sec, _en, _vi, _topic) in _ITEMS_RAW.items():
    _rba = rba_ref(_sec, _en.split(" - ")[0], _vi.split(" - ")[0])
    _vn = VN[_topic]
    ITEMS[_id] = {
        "rba_section": _sec,
        "rba_topic_en": _en,
        "rba_topic_vi": _vi,
        "rba_citation_en": _rba["en"],
        "rba_citation_vi": _rba["vi"],
        "vn_citations_en": _vn["en"],
        "vn_citations_vi": _vn["vi"],
    }

# Dòng sheet Checklist tương ứng từng Audit ID (dòng 6 = AL101 ... dòng 57 = ES404)
_ORDER = list(_ITEMS_RAW.keys())
ITEM_ROWS = {i: 6 + n for n, i in enumerate(_ORDER)}
ROW_ITEMS = {r: i for i, r in ITEM_ROWS.items()}
