# -*- coding: utf-8 -*-
"""
Động cơ biên tập nội dung cột L (Description of Supporting Evidence of Violation).

Chức năng:
1. parse_draft()          - phân tích draft "Audit findings: ... Legal references: ..." của kiểm toán viên
2. tách song ngữ          - dòng "VI / EN" được tách thành nội dung tiếng Việt + tiếng Anh
3. structure_bilingual()  - dựng khối chuẩn: AUDIT FINDINGS (ENGLISH) + Legal references +
                            NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT) + Căn cứ pháp lý
4. build_content()        - chọn nội dung curated (đã kiểm chuẩn) hoặc dựng từ draft + KB
5. normalize_placeholder()- chuẩn hóa ô trống theo mẫu báo cáo
"""
import json
import os
import re
import unicodedata
from dataclasses import dataclass, field

try:  # chay truc tiep (script) hoac nhu package
    from knowledge_base import ITEMS, NON_CONFORMANCE_RATINGS
    from translator import translate_findings, translate_text, VI_WORDS_NO_DIACRITICS
    from elaborations import _ELABORATE_MIN_LEN, _elaborate_key, ELABORATIONS_FULL
except ImportError:  # pragma: no cover
    from .knowledge_base import ITEMS, NON_CONFORMANCE_RATINGS
    from .translator import translate_findings, translate_text, VI_WORDS_NO_DIACRITICS
    from .elaborations import _ELABORATE_MIN_LEN, _elaborate_key, ELABORATIONS_FULL

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# Mẫu ô trống chuẩn (theo file báo cáo đầu ra được duyệt)
PLACEHOLDER = "Audit findings:\n1: \n2:\n3:\n\nLegal references:"

# Header chuẩn của khối đã biên tập
HEADER_EN = "AUDIT FINDINGS (ENGLISH)"
HEADER_VI = "NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)"
REFS_EN = "Legal references:"
REFS_VI = "Căn cứ pháp lý:"

# Dòng phân cách giữa phần tiếng Anh và tiếng Việt
SEPARATOR = "---------------------------------------"

# Các dòng header được in đậm trong ô L (đúng định dạng file mẫu đã duyệt).
# Lưu ý: "Audit findings:" (mẫu ô trống) KHÔNG đậm — chỉ "Legal references:" đậm.
BOLD_HEADERS = frozenset({HEADER_EN, REFS_EN, HEADER_VI, REFS_VI})


def _norm_unicode(text):
    """
    Chuẩn hóa Unicode dạng NFC (composed).

    Tiếng Việt có THỂ bị lưu ở dạng NFD (decomposed): ví dụ chữ "Ộ" thành "O" +
    dấu nặng + dấu hỏi (3 ký tự thay vì 1). Khi đó chuỗi "NỘI DUNG ĐÁNH GIÁ
    (TIẾNG VIỆT)" dài 38 ký tự thay vì 30 và KHÁC byte so với bản NFC, nên phép
    so khớp chính xác (dùng để bôi đậm header) thất bại -> header tiếng Việt bị
    rơi khỏi danh sách in đậm. File do macOS sinh ra hoặc copy-paste từ một số
    nguồn thường ở dạng NFD. Chuẩn hóa về NFC trước khi so khớp để header luôn
    được nhận diện đúng bất kể dạng Unicode của file đầu vào.
    """
    return unicodedata.normalize("NFC", text)


def _is_bold_header(line):
    """Dòng có phải header khối cần in đậm không (bỏ qua khác biệt NFC/NFD, khoảng trắng, dấu hai chấm)."""
    s = _norm_unicode(line or "").replace(" ", " ").strip()
    if s in BOLD_HEADERS:
        return True
    s_clean = re.sub(r":\s*$", "", s).strip().upper()
    for h in BOLD_HEADERS:
        if s_clean == re.sub(r":\s*$", "", h).strip().upper():
            return True
    return False


def to_rich_text(text, base_font=None):
    """
    Chuyển nội dung cột L thành rich text: các dòng header khối
    (AUDIT FINDINGS (ENGLISH) / Legal references: / NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT) /
    Căn cứ pháp lý:) được in đậm, phần findings và trích dẫn giữ chữ thường.

    Mọi run đều mang rPr hoàn chỉnh (tên font + cỡ, kế thừa từ base_font của ô,
    chỉ khác thuộc tính đậm) để mọi phiên bản Excel/WPS hiển thị đúng đậm/thường.
    """
    if not isinstance(text, str) or not text:
        return text
    clean_text = _norm_unicode(text).replace("\r\n", "\n").replace("\r", "\n")
    lines = clean_text.split("\n")
    if not any(_is_bold_header(ln) for ln in lines):
        return clean_text
    name, size = "Calibri", 10.0
    if base_font is not None:
        name = getattr(base_font, "name", None) or getattr(base_font, "rFont", None) or name
        size = getattr(base_font, "sz", None) or getattr(base_font, "size", None) or size
    bold_font = InlineFont(rFont=name, sz=size, b=True)
    plain_font = InlineFont(rFont=name, sz=size, b=False)
    parts = []
    for i, ln in enumerate(lines):
        sep = "\n" if i < len(lines) - 1 else ""
        parts.append(TextBlock(bold_font if _is_bold_header(ln) else plain_font,
                               ln + sep))
    return CellRichText(*parts)

# Trạng thái draft-placeholder cũ trong file đầu vào (42 ô)
_OLD_PLACEHOLDER_RE = re.compile(
    r"^Audit findings:\s*1:\s*2:\s*3:\s*Legal references:\s*$",
    re.DOTALL,
)

# Thư mục chứa file curated_findings.json
_APP_DIR = os.path.dirname(os.path.abspath(__file__))

LABOR_CODE_EN_ = "Labor Code No. 45/2019/QH14"

# Từ viết tắt phổ biến trong draft -> dạng đầy đủ (dùng khi biên tập/tra cứu)
ABBREVIATIONS = {
    "NLD": "người lao động",
    "NLĐ": "người lao động",
    "NSDLĐ": "người sử dụng lao động",
    "NSDLD": "người sử dụng lao động",
    "CNV": "công nhân viên",
    "HĐLĐ": "hợp đồng lao động",
    "BHXH": "bảo hiểm xã hội",
    "PCCC": "phòng cháy chữa cháy",
    "OT": "overtime (làm thêm giờ)",
    "TTS": "thực tập sinh",
    "HSSV": "học sinh sinh viên",
    "TNCN": "thu nhập cá nhân",
    "BHLĐ": "bảo hộ lao động",
    "ATVSLĐ": "an toàn vệ sinh lao động",
    "KTV": "kiểm toán viên",
}

VN_CHARS_RE = re.compile(
    r"[ăâđêôơưáàảãạấầẩẫậắằẳẵặéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]",
    re.IGNORECASE,
)

# Mã văn bản pháp luật / viết tắt chứa ký tự Việt (NĐ-CP, NLĐ, NĐ, TT-BLDTBXH).
# Chúng xuất hiện hợp lệ trong trích dẫn TIẾNG ANH (vd Decree No. 356/2025/NĐ-CP)
# nên phải loại khỏi văn bản trước khi nhận diện ngôn ngữ.
_LEGAL_ACRONYM_RE = re.compile(
    r"N[ĐD]\s*-\s*CP|TT\s*-\s*BLDTBXH|\bNLĐ\b|\bN[ĐD]\b",
    re.IGNORECASE,
)


def is_vietnamese(text):
    """
    Đoạn văn có chứa tiếng Việt hay không (có dấu hoặc không dấu).
    - Có dấu Việt -> chắc chắn là tiếng Việt.
    - Không dấu: đếm từ thuộc tập từ Việt không dấu; >= 20% -> coi là tiếng Việt
      (threshold thấp để bắt được "Vieglish" — Anh lẫn Việt không dấu).

    LƯU Ý: mã văn bản pháp luật / viết tắt chứa ký tự Việt (NĐ-CP, NLĐ,
    TT-BLDTBXH...) xuất hiện hợp lệ trong TRÍCH DẪN tiếng Anh (vd
    "Decree No. 356/2025/NĐ-CP, Article 19") nên bị loại ra trước khi kiểm tra,
    nếu không trích dẫn EN sẽ bị chuyển nhầm sang phần tiếng Việt.
    """
    if not text:
        return False
    # Bỏ qua mã văn bản pháp luật / viết tắt chứa ký tự Việt
    t = _LEGAL_ACRONYM_RE.sub(" ", text)
    # Nhanh: có ký tự Việt có dấu -> chắc chắn là tiếng Việt
    if VN_CHARS_RE.search(t):
        return True
    # Không dấu: đếm từ thuộc tập từ Việt không dấu
    words = re.findall(r"[A-Za-zÀ-ỹ]+", t)
    if not words:
        return False
    vi_count = sum(1 for w in words if w.lower() in VI_WORDS_NO_DIACRITICS)
    return vi_count / len(words) >= 0.20


# ---------------------------------------------------------------------------
# Kiểm tra "THUẦN MỘT NGÔN NGỮ" cho từng mục trong cột L
# ---------------------------------------------------------------------------
# Từ tiếng Việt KHÔNG DẤU đặc trưng - CHỈ gồm từ KHÔNG thể là từ tiếng Anh.
# (Cố ý loại các từ Việt không dấu trùng từ tiếng Anh: so, do, may, hop, dong,
#  ban, tin, ca, la, re, te, mi, com, chi, den, tim, cam, dat, thu, them, sat...
#  để tránh báo động nhầm trên một câu tiếng Anh hợp lệ. VD "We told them about
#  it" là tiếng Anh thuần - KHÔNG được coi là còn tiếng Việt chỉ vì "them").
_VI_SAFE_TOKENS = frozenset("""
khong duoc nguoi chua cua voi cac nhu phai thuc nghiem hien bao cao luong tien
nghi phep ngay thang tuan gio phut khoan dieu luat thoa uoc xac nhan trach nhiem
quyen khao xung truong phan chuc thiet kiem soat hieu danh khac phuc chinh
sua hoan thien pham tuyen giam nang thao quan tiet cong dinh quy
""".split())


def has_vietnamese_words(text):
    """
    Đoạn văn còn chứa TỪ tiếng Việt hay không (ký tự có dấu tiếng Việt hoặc từ
    Việt không dấu đặc trưng). Dùng để bảo đảm khối
    'AUDIT FINDINGS (ENGLISH)' / 'Legal references:' KHÔNG còn sót tiếng Việt.

    - Mã văn bản luật chứa ký tự Việt (NĐ-CP, NLĐ, TT-BLDTBXH...) được loại bỏ
      trước khi kiểm tra vì chúng hợp lệ trong trích dẫn tiếng Anh.
    - Từ Việt KHÔNG DẤU chỉ bị coi là "lẫn tiếng Việt" khi viết thường trong câu
      (bỏ qua dạng Viết Hoa/IN HOA vì đó thường là TÊN RIÊNG đã phiên âm được
      phép xuất hiện trong câu tiếng Anh, vd: Cong Ty Bao Ve Dang Quang).
    """
    if not text or not isinstance(text, str):
        return False
    t = _LEGAL_ACRONYM_RE.sub(" ", text)
    if VN_CHARS_RE.search(t):
        return True
    for m in re.finditer(r"[A-Za-z]+", t):
        w = m.group(0)
        if w.islower() and w in _VI_SAFE_TOKENS:
            return True
    return False


def has_english_words(text):
    """
    Đoạn văn còn chứa TỪ tiếng Anh đặc trưng hay không. Dùng để bảo đảm khối
    'NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)' / 'Căn cứ pháp lý:' KHÔNG còn sót tiếng Anh.
    """
    return _has_english_words(text)


@dataclass
class Draft:
    """Draft nội dung cột L của một dòng checklist."""
    raw: str = ""
    items_vi: list = field(default_factory=list)
    items_en: list = field(default_factory=list)
    refs_raw: str = ""
    has_content: bool = False

    @property
    def is_empty(self):
        return not self.has_content


def split_bilingual_line(line):
    """
    Tách dòng song ngữ kiểu 'Nội dung tiếng Việt / English content'.
    Trả về (vi_text, en_text). Nếu không phải song ngữ, trả về (None, None).
    """
    if "/" not in line or not is_vietnamese(line):
        return None, None
    parts = [p.strip() for p in line.split(" / ")]
    if len(parts) == 2 and not is_vietnamese(parts[1]):
        return parts[0], parts[1]
    return None, None


# Marker đầu dòng trích dẫn: gạch đầu dòng "- " hoặc đánh số "1. " / "1: " / "1) "
REF_MARKER_RE = re.compile(r"^(?:\d+\s*[.:)]|-)\s*(.*)$")

# Nhãn phụ rỗng kiểu "Findings:" / "Finding:" / "Audit Findings:" trong draft nhà máy
# (chỉ là tiêu đề phụ, không phải finding thật -> phải loại bỏ)
FINDING_LABEL_RE = re.compile(r"^(?:audit\s+)?findings?\s*:?\s*$", re.IGNORECASE)


def is_finding_label(text):
    """Dòng chỉ là nhãn phụ kiểu 'Findings:' / 'Finding' (không mang nội dung)."""
    return bool(FINDING_LABEL_RE.match((text or "").strip()))


def strip_finding_label(text):
    """Bỏ nhãn 'Findings:' đầu dòng; nếu cùng dòng có nội dung thật thì giữ lại."""
    s = (text or "").strip()
    m = re.match(r"^(?:audit\s+)?findings?\s*:\s*(.+)$", s, re.IGNORECASE | re.DOTALL)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return s


def split_refs(raw):
    """
    Tách chuỗi trích dẫn pháp lý thô thành hai danh sách theo ngôn ngữ:
        (en_refs, vi_refs)
    Hỗ trợ dòng đơn ngữ hoặc song ngữ 'VI / EN' (tự tách, kể cả khi dấu '/'
    viết dính). Bỏ marker đầu dòng. Đảm bảo phần tiếng Anh chỉ chứa tiếng Anh,
    phần tiếng Việt chỉ chứa tiếng Việt.
    """
    en_refs, vi_refs = [], []
    if not raw or not isinstance(raw, str):
        return en_refs, vi_refs
    for line in raw.replace("\r\n", "\n").split("\n"):
        s = line.strip()
        if not s:
            continue
        m = REF_MARKER_RE.match(s)
        s = m.group(1).strip() if m else s
        if not s:
            continue
        en, vi = split_mixed_line(s)
        if en and vi:
            vi_refs.append(vi)
            en_refs.append(en)
            continue
        if is_vietnamese(s):
            vi_refs.append(s)
        else:
            en_refs.append(s)
    return en_refs, vi_refs


def is_legal_reference(text):
    """
    Kiểm tra một dòng có phải là trích dẫn pháp lý/quy phạm pháp luật hay không.
    Tránh nhầm lẫn điều khoản hợp đồng của nhà máy (vd: 'Điều 5.1.2 Thời hạn báo trước trong hợp đồng...')
    với điều luật nhà nước (vd: 'Điều 45 khoản 1 Luật lao động 2019', 'Nghị định 145/2020/NĐ-CP').
    """
    if not text or not isinstance(text, str):
        return False
    s = text.strip()
    if re.match(r"^(?:legal\s+references?|căn\s+cứ\s+pháp\s+lý|căn\s+cứ|cơ\s+sở\s+pháp\s+lý)\s*:", s, re.IGNORECASE):
        return True
    
    m = re.match(r"^(?:\d+\s*[.:)]|-)\s*(.*)$", s)
    core = m.group(1).strip() if m else s
    
    narrative_words = [
        "quy định nsdlđ", "phát hiện", "kiểm tra hồ sơ", "phỏng vấn", "công ty đang",
        "chưa thực hiện", "vi phạm", "không phù hợp", "thanh toán", "thay vì",
        "chưa khấu trừ", "vượt quá", "không có quy trình", "thực tiễn này",
        "qua kiểm tra", "hạn chế quyền", "công ty có"
    ]
    core_lower = core.lower()
    if any(nw in core_lower for nw in narrative_words):
        return False
        
    has_law_id = bool(re.search(r"\d+/\d{4}/(?:N[ĐD]-CP|TT-[A-Z]+|QH\d+)", core, re.IGNORECASE))
    has_law_name = bool(re.search(
        r"\b(?:luật|bộ\s+luật|nghị\s+định|thông\s+tư|quyết\s+định|hiến\s+pháp|labor\s+code|decree|circular|law|constitution|code\s+of\s+conduct|rba)\b",
        core, re.IGNORECASE))
    has_article = bool(re.search(r"\b(?:điều|article|khoản|clause|mục|section)\s+\d+", core, re.IGNORECASE))
    
    if has_law_id:
        return True
    if has_law_name and has_article:
        return True
    if has_law_name and len(core.split()) <= 12 and not any(ch in core for ch in [". ", "! ", "? "]):
        return True
    return False


def parse_draft(text):
    """Phân tích nội dung cột L thành Draft."""
    d = Draft(raw=text or "")
    if not text or not isinstance(text, str):
        return d

    lines = text.replace("\r\n", "\n").split("\n")
    in_refs = False
    findings_block = []
    refs_block = []

    for ln in lines:
        s = ln.strip()
        if re.match(r"^(?:legal\s+references?|căn\s+cứ\s+pháp\s+lý|căn\s+cứ|cơ\s+sở\s+pháp\s+lý)\b", s, re.IGNORECASE):
            in_refs = True
            rest = s.split(":", 1)[1] if ":" in s else ""
            if rest.strip():
                refs_block.append(rest.strip())
            continue
        if re.match(r"^(?:audit\s+findings?|nội\s+dung\s+đánh\s+giá)\b", s, re.IGNORECASE):
            continue
        if not s:
            continue
        if in_refs:
            refs_block.append(s)
        elif is_legal_reference(s):
            refs_block.append(s)
        else:
            findings_block.append(s)

    # Loại bỏ item rỗng kiểu "1:", "2:" (chỉ đánh số không có nội dung)
    item_re = re.compile(r"^(\d+)\s*[:.)]\s*(.*)$")
    for s in findings_block:
        m = item_re.match(s)
        body = m.group(2).strip() if m else s
        # Bỏ nhãn phụ kiểu "Findings:" (giữ lại nội dung thật nếu có cùng dòng)
        body = strip_finding_label(body)
        if not body or is_finding_label(body):
            continue
        # Nếu dòng thực chất là căn cứ pháp lý của KTV -> chuyển sang refs_block
        if is_legal_reference(body):
            refs_block.append(body)
            continue
        # Tách dòng trộn lẫn 'EN / VI' (cả hai thứ tự) trước khi phân bổ
        me, mv = split_mixed_line(body)
        if me or mv:
            if me:
                d.items_en.append(me)
            if mv:
                d.items_vi.append(mv)
            continue
        vi, en = split_bilingual_line(body)
        if vi and en:
            d.items_vi.append(vi)
            d.items_en.append(en)
        elif is_vietnamese(body):
            d.items_vi.append(body)
        else:
            d.items_en.append(body)

    d.refs_raw = "\n".join(refs_block)
    d.has_content = bool(d.items_vi or d.items_en or d.refs_raw)
    return d


def is_placeholder(text):
    """Ô L có phải dạng trống/placeholder hay không (cần chuẩn hóa)."""
    if text is None:
        return True
    if not isinstance(text, str):
        return False
    s = text.strip()
    if not s:
        return True
    return bool(_OLD_PLACEHOLDER_RE.match(s)) or s == PLACEHOLDER.strip()


def parse_bilingual(text):
    """
    Phân tích khối đã biên tập (có HEADER_EN/HEADER_VI) thành các phần:
        {en_findings, en_refs, vi_findings, vi_refs}
    Dùng để hiển thị lại nội dung trong GUI cho phép chỉnh sửa từng phần.
    """
    if not isinstance(text, str):
        return {"en_findings": [], "en_refs": [], "vi_findings": [], "vi_refs": []}
    lines = text.replace("\r\n", "\n").split("\n")

    mode = "en"
    en_body, vi_body = [], []
    for ln in lines:
        s = ln.strip()
        if s == HEADER_EN:
            mode = "en"
            continue
        if s == HEADER_VI:
            mode = "vi"
            continue
        if s == SEPARATOR:
            continue
        (vi_body if mode == "vi" else en_body).append(ln)

    def split_section(block, refs_header):
        body, refs, in_refs = [], [], False
        for ln in block:
            s = ln.strip()
            if s == refs_header:
                in_refs = True
                continue
            if not s:
                continue
            if in_refs:
                # Bỏ marker đầu dòng (gạch đầu dòng '- ' hoặc số '1. ') của trích dẫn
                m = re.match(r"^(?:\d+\s*[.:)]|-)\s*(.*)$", s)
                refs.append(m.group(1).strip() if m else s)
            else:
                body.append(s)
        return body, refs

    en_f, en_r = split_section(en_body, REFS_EN)
    vi_f, vi_r = split_section(vi_body, REFS_VI)
    return {
        "en_findings": split_findings("\n".join(en_f)),
        "en_refs": en_r,
        "vi_findings": split_findings("\n".join(vi_f)),
        "vi_refs": vi_r,
    }


def split_findings(block_text):
    """Tách các finding đánh số '1.' / '1:' thành danh sách."""
    items = []
    cur = None
    for line in block_text.split("\n"):
        m = re.match(r"^\s*(\d+)\s*[.:)]\s*(.*)$", line)
        if m:
            if cur is not None:
                items.append(cur.strip())
            cur = m.group(2)
        elif cur is not None:
            cur += "\n" + line
    if cur is not None:
        items.append(cur.strip())
    # Bỏ item chỉ là nhãn "Findings:" và nhãn đầu dòng của item đa dòng
    cleaned = []
    for it in items:
        lines = it.split("\n")
        while lines and is_finding_label(lines[0]):
            lines.pop(0)
        body = strip_finding_label("\n".join(lines))
        if body:
            cleaned.append(body)
    return cleaned


def separate_languages(en_findings, en_refs, vi_findings, vi_refs):
    """Dam bao EN chua Anh / VI chua Viet, KHONG lam mat muc.

    Giu nguyen thu tu va noi dung goc; chi chuyen muc that su lac ngon ngu
    (EN co dau tieng Viet -> ve VI; VI thu`an Anh -> ve EN).
    Draft tieng Viet KHONG DAU duoc giu nguyen o phia VI.
    """
    en_f, en_r, vi_f, vi_r = [], [], [], []
    for t in list(en_findings or []):
        # EN co dau tieng Viet -> thuc chat la muc VI -> chuyen ve VI
        if VN_CHARS_RE.search(t or ""):
            vi_f.append(t)
        else:
            en_f.append(t)
    for t in list(vi_findings or []):
        # VI thuan Anh (khong dau, khong tu Viet) -> ve EN; con lai giu o VI
        # (bao gom Viet khong dau).
        if t and not is_vietnamese(t):
            en_f.append(t)
        else:
            vi_f.append(t)
    for t in list(en_refs or []) + list(vi_refs or []):
        (vi_r if is_vietnamese(t) else en_r).append(t)
    return en_f, en_r, vi_f, vi_r


# --------------------------------------------------------------- thanh lọc ----
# Dòng trộn lẫn Anh/Việt phân tách bằng dấu " / " (thường gặp trong draft nhà máy)
MIXED_SEP_RE = re.compile(r"\s+/\s+")
# Dấu '/' viết DÍNH (không có khoảng trắng quanh dấu - vd "trước khi kí./ The
# labor contract ...") vẫn là ranh giới hai ngôn ngữ -> tách "lỏng" có kiểm soát
# để KHÔNG làm rơi nửa ngôn ngữ của mục (nội dung cột L phải đầy đủ).
_LOOSE_SEP_RE = re.compile(r"\s*/+\s*")
# Đoạn chứa số hiệu văn bản luật / số liệu theo '/' -> KHÔNG tách lỏng
# (tránh cắt nhầm "Thông tư 10/2020/TT-BLDTBXH" hay "8 giờ/ngày").
_CODE_SEG_RE = re.compile(
    r"\d+\s*/\s*\d+|/\s*(?:ND|NĐ|TT|QH)|N[ĐD]\s*-\s*CP|TT\s*-\s*BLDTBXH",
    re.IGNORECASE)
# Ngưỡng an toàn cho tách lỏng: mỗi đoạn phải là câu/mệnh đề thật
_LOOSE_SEG_MIN_LEN = 12
_LOOSE_SPLIT_MIN_LEN = 25
_PAREN_RE = re.compile(r"\(([^()]{10,})\)")

# Từ tiếng Anh phổ biến dùng để nhận diện đoạn "Vieglish" (Việt lẫn từ Anh).
# CHỈ chứa từ KHÔNG BAO GIỜ trùng với từ tiếng Việt (kể cả Việt ngữ không dấu):
# loại "do" (vi: do... = vì), "may" (vi: may mặc), "can" (vi: can thiệp),
# "to/in/on/at/by/as/or/is" (quá ngắn hoặc trùng vi: nhà "to").
_COMMON_EN_WORDS = frozenset("""
the and for with was were has had not this that these those their there
under over between within without into upon upon from during about against
section article clause paragraph point law code conduct standard regulation
company factory workers worker employees employee person people
documents document original certified copy retention confiscate confiscating
prohibited according shall must should will would could but its
labor working hours wage wages payment benefits overtime
""".split())


def _has_english_words(text, minimum=1):
    """Đoạn văn có chứa từ tiếng Anh 'đặc trưng' (không thể là tiếng Việt) hay không."""
    words = set(re.findall(r"[a-z]+", (text or "").lower()))
    return len(words & _COMMON_EN_WORDS) >= minimum


def _loose_split_ok(segs):
    """
    Chỉ chấp nhận tách "lỏng" (dấu '/' viết dính, không có khoảng trắng hai bên)
    khi MỌI đoạn đều là câu/mệnh đề thật và không phải mã văn bản pháp luật/số
    liệu. Nhờ vậy mục song ngữ kiểu "... trước khi kí./ The labor contract ..."
    được tách đúng mà "Thông tư 10/2020/TT-BLDTBXH" hay "8 giờ/ngày" không bị cắt.
    """
    if len(segs) < 2:
        return False
    if any(_CODE_SEG_RE.search(s) for s in segs):
        return False
    if any(len(s) < _LOOSE_SEG_MIN_LEN or len(re.findall(r"\w+", s)) < 3
           for s in segs):
        return False
    return max(len(s) for s in segs) >= _LOOSE_SPLIT_MIN_LEN


def split_mixed_line(text):
    """
    Tách dòng trộn lẫn hai ngôn ngữ phân tách bằng ' / ' thành (en_part, vi_part).
    Hỗ trợ cả thứ tự 'EN / VI' lẫn 'VI / EN' và nhiều đoạn; đoạn trộn kiểu
    'Vieglish' (Việt lẫn từ Anh) được xếp về phần tiếng Việt. Trả về (None, None)
    nếu dòng không chứa đồng thời cả hai ngôn ngữ.

    Tách 2 bước: (1) dấu '/' có khoảng trắng hai bên; (2) dấu '/' viết dính
    (có kiểm soát ở _loose_split_ok) - bảo đảm mục song ngữ không bị rơi nửa
    ngôn ngữ, nội dung cột L không bị thiếu.
    """
    if not text or "/" not in text:
        return None, None
    for sep, loose in ((MIXED_SEP_RE, False), (_LOOSE_SEP_RE, True)):
        segs = [p.strip() for p in sep.split(text.strip()) if p.strip()]
        if len(segs) < 2:
            continue
        if loose and not _loose_split_ok(segs):
            continue
        en, vi = [], []
        for s in segs:
            (vi if is_vietnamese(s) else en).append(s)
        if not en or not vi:
            continue
        return " ".join(en), " ".join(vi)
    return None, None


def _word_set(s):
    return set(w for w in re.findall(r"\w+", s.lower()) if len(w) >= 2)


def _overlap(a, b):
    wa, wb = _word_set(a), _word_set(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / min(len(wa), len(wb))


def _is_mixed_item(t):
    """Mục tiếng Việt còn lẫn từ tiếng Anh (kiểu 'Vieglish') hay không."""
    return is_vietnamese(t) and _has_english_words(t)


# Tách từ: số / chữ cái (hỗ trợ dấu nháy & gạch nối) / khoảng trắng / ký hiệu
_TOKEN_RE = re.compile(r"\d+|[^\W\d_]+(?:['\u2019-][^\W\d_]+)*|\s+|[^\w\s]", re.UNICODE)
_EN_WORD_RE = re.compile(r"[A-Za-z][A-Za-z'\u2019-]+")


def _vn_word(tok):
    """
    Token là từ tiếng Việt (chứa ký tự có dấu HOẶC nằm trong tập từ Việt không dấu).
    Từ viết tắt IN HOA (NĐ, NLĐ, CP, QH, BHXH...) KHÔNG được coi là từ cần dịch,
    để bảo toàn mã trích dẫn pháp luật khi sửa câu Anh lẫn Việt.
    """
    if tok.isupper():
        return False
    if VN_CHARS_RE.search(tok):
        return True
    return tok.lower() in VI_WORDS_NO_DIACRITICS


def _en_word(tok):
    """Token là từ thuần ASCII (khác biệt với từ tiếng Việt có dấu)."""
    return bool(_EN_WORD_RE.fullmatch(tok))


def _translate_run(run, target):
    """
    Dịch một đoạn ngôn ngữ lạc; chỉ nhận kết quả đạt đúng ngôn ngữ đích.

    Ưu tiên từ điển chuyên ngành audit ngoại tuyến (chuẩn ngữ cảnh báo cáo HR):
    nếu từ điển dịch ĐẠT đúng ngôn ngữ đích thì dùng ngay — tránh Google dịch
    lệch ngữ cảnh cụm chuyên ngành (vd 'quản lý hiệu quả' -> 'effective
    management' thay vì 'effectively managed'). Nếu từ điển không có cụm đó
    (dịch không đạt / trả nguyên văn) mới fallback sang Google online; Google
    cũng không đạt thì trả None (đoạn được giữ nguyên ở caller).

    Ngoại lệ: nếu từ điển CHỦ ĐỘNG dịch run thành chuỗi rỗng (các pair dạng
    ("việc", ""), ("được", ""), ("đã", "") — danh tử/trợ động từ tiếng Việt
    thường bị lược khi sang Anh) thì trả về chuỗi rỗng để caller bỏ từ này.
    """
    def _valid(s, allow_empty=False):
        if s is None:
            return None
        if not s.strip():
            return "" if allow_empty else None
        s = s.strip()
        if target == "en" and is_vietnamese(s):
            return None
        if target == "vi" and not is_vietnamese(s):
            return None
        return s

    offline_raw = translate_text(run, target, prefer_online=False)
    allow_empty = bool(offline_raw is not None and offline_raw.strip() == ""
                       and run.strip())
    offline = _valid(offline_raw, allow_empty=allow_empty)
    if offline is not None:
        return offline
    return _valid(translate_text(run, target))


def repair_mixed_item(text, target):
    """
    Sửa mục lẫn hai ngôn ngữ về THUẦN một ngôn ngữ: các run từ thuộc ngôn ngữ
    'lạc' (thiểu số) được dịch sang 'target', số/ký hiệu/phần đúng ngôn ngữ giữ
    nguyên. VD: 'Does not have khảo sát and report conflict of interest định kỳ
    1 year/lần' + target 'en' -> '... survey ... periodically 1 year/times'.
    Nếu không dịch được đoạn nào (offline thiếu mẫu) trả về nguyên văn.
    """
    if not isinstance(text, str) or not text.strip():
        return text
    if target == "en" and not is_vietnamese(text):
        return text
    if target == "vi" and not is_vietnamese(text):
        return text
    want = _vn_word if target == "en" else _en_word
    tokens = _TOKEN_RE.findall(text)
    out, i, n, changed = [], 0, len(tokens), False
    while i < n:
        tok = tokens[i]
        if want(tok) and tok.strip():
            run, j = [tok], i + 1
            while j < n and (
                    want(tokens[j])
                    or (tokens[j].isspace() and j + 1 < n and want(tokens[j + 1]))):
                run.append(tokens[j])
                j += 1
            i = j
            run_text = "".join(run)
            # Run toàn chữ IN HOA là mã/viết tắt (QH14, NĐ-CP, BHXH...) ->
            # giữ nguyên, không dịch để tránh hỏng trích dẫn pháp luật.
            tr = (None if run_text.strip().isupper()
                  else _translate_run(run_text, target))
            if tr is not None:
                if tr:
                    # Giữa câu thì viết thường (trừ đầu câu / sau dấu chấm)
                    if re.search(r"[.!?]\s*$", "".join(out)) or not "".join(out).strip():
                        tr = tr
                    else:
                        tr = tr[0].lower() + tr[1:]
                # tr == "" -> từ điển chủ động bỏ từ (danh tử "việc", trợ từ
                # "được"/"đã"...) -> bỏ khỏi câu Anh
                out.append(tr)
                changed = True
            else:
                out.append(run_text)
        else:
            out.append(tok)
            i += 1
    if not changed:
        return text
    return re.sub(r"[ \t]{2,}", " ", "".join(out)).strip()


def dedupe_items(items):
    """
    Loại mục trùng lặp (giống hệt) và mục 'lẫn ngôn ngữ' khi tồn tại biến thể
    thuần tiếng tương đương (trùng >= 60% từ) — giữ lại bản thuần.
    """
    out, seen = [], set()
    for t in items:
        k = re.sub(r"\s+", " ", (t or "").strip()).lower()
        if k and k not in seen:
            seen.add(k)
            out.append(t.strip())
    kept = []
    for i, t in enumerate(out):
        if _is_mixed_item(t) and any(
                j != i and not _is_mixed_item(u) and _overlap(t, u) >= 0.6
                for j, u in enumerate(out)):
            continue
        kept.append(t)
    return kept


def purify_findings(en_findings, vi_findings):
    """
    Thanh lọc findings về 'EN thuần Anh / VI thuần Việt':
    1. Tách dòng trộn có dấu ' / ' -> nửa đúng ngôn ngữ về đúng section.
    2. Mục trộn KHÔNG có dấu ' / ' (thiên về một ngôn ngữ nhưng lẫn từ ngôn ngữ
       kia, VD 'Does not have khảo sát and report conflict of interest định kỳ'):
       dịch các đoạn ngôn ngữ lạc về ngôn ngữ của section gốc.
    3. Bản lẫn lộn bị loại nếu đã có bản thuần tương đương (>= 60% từ trùng).
    """
    en, vi = [], []
    pending = []  # (item_lẫn_lộn, section_gốc 'en'/'vi')

    for origin_tag, items in (("en", en_findings), ("vi", vi_findings)):
        for t in items:
            e, v = split_mixed_line(t)
            if e:
                en.append(e)
            if v:
                vi.append(v)
                if _is_mixed_item(v):
                    pending.append((v, origin_tag))
            if e or v:
                continue
            pool = vi if is_vietnamese(t) else en
            pool.append(t)
            if _is_mixed_item(t):
                pending.append((t, origin_tag))

    for t, tag in pending:
        target_pool = vi if tag == "vi" else en
        other_pool = en if tag == "vi" else vi
        has_pure = any(
            not _is_mixed_item(u) and _overlap(t, u) >= 0.6
            for pool in (target_pool, other_pool) for u in pool if u is not t)
        for pool in (en, vi):
            while t in pool:
                pool.remove(t)
        if not has_pure:
            target_pool.append(repair_mixed_item(t, tag))
    return dedupe_items(en), dedupe_items(vi)


def purify_ref(text, target):
    """
    Làm sạch trích dẫn pháp lý: dịch các đoạn trong ngoặc () còn sót ngôn ngữ kia
    về ngôn ngữ đích (qua translate_text, có cache online). Nếu bản dịch không đạt
    đúng ngôn ngữ đích (mẫu lạ / offline dịch không xong) thì giữ nguyên.
    """
    def repl(m):
        inner = m.group(1)
        inner_vi = is_vietnamese(inner)
        if target == "vi" and not inner_vi:
            tr = translate_text(inner, "vi")
            if tr and is_vietnamese(tr):
                return f"({tr.strip()})"
            return m.group(0)
        if target == "en" and inner_vi:
            tr = translate_text(inner, "en")
            if tr and not is_vietnamese(tr):
                return f"({tr.strip()})"
            return m.group(0)
        return m.group(0)

    return _PAREN_RE.sub(repl, text)


_BULLET_RE = re.compile(r"^[-–—*]\s+")


def _clean_numbered_item(t):
    """Bỏ dấu gạch đầu mục (đã có số) và viết hoa chữ cái đầu câu."""
    t = _BULLET_RE.sub("", (t or "").strip())
    if t and t[0].isalpha() and t[0].islower():
        t = t[0].upper() + t[1:]
    return t


def numbered(items, marker="{}.", suffix=" "):
    """Danh sách -> chuỗi đánh số, đã bỏ gạch đầu mục và viết hoa đầu câu."""
    cleaned = [t for t in items if t and str(t).strip()]
    return "\n".join(
        f"{marker.format(i + 1)}{suffix}{_clean_numbered_item(t)}"
        for i, t in enumerate(cleaned))


def structure_bilingual(en_findings, en_refs, vi_findings, vi_refs):
    """
    Dựng khối chuẩn song ngữ theo format báo cáo:
        AUDIT FINDINGS (ENGLISH)
        1. ...
        Legal references:
        1. ...
        ---------------------------------------
        NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)
        1. ...
        Căn cứ pháp lý:
        1. ...
    """
    en_block = ""
    if en_findings:
        en_block += numbered(en_findings) + "\n"
    en_block += (("\n" if en_findings else "") + f"{REFS_EN}\n{numbered(en_refs)}") if en_refs \
        else (("\n" if en_findings else "") + REFS_EN)

    vi_block = ""
    if vi_findings:
        vi_block += numbered(vi_findings) + "\n"
    vi_block += (("\n" if vi_findings else "") + f"{REFS_VI}\n{numbered(vi_refs)}") if vi_refs \
        else (("\n" if vi_findings else "") + REFS_VI)

    return f"{HEADER_EN}\n{en_block}\n{SEPARATOR}\n{HEADER_VI}\n{vi_block}"


def load_curated(path=None):
    """Nạp nội dung đã kiểm chuẩn (mặc định: curated_findings.json cạnh module)."""
    p = path or os.path.join(_APP_DIR, "curated_findings.json")
    if not os.path.exists(p):
        return {}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def kb_refs_for(audit_id, lang="vi"):
    """Trích dẫn pháp luật từ KB cho một Audit ID."""
    item = ITEMS.get(audit_id)
    if not item:
        return []
    if lang == "vi":
        return [item["rba_citation_vi"]] + item["vn_citations_vi"]
    return [item["rba_citation_en"]] + item["vn_citations_en"]


def translate_citations_vi_to_en(ref):
    """Dịch trích dẫn pháp luật Việt Nam (Điều/khoản/Nghị định) sang dạng tiếng Anh."""
    s = ref
    s = re.sub(r"Bộ luật Lao động năm 2019 \(Luật số 45/2019/QH14\)", LABOR_CODE_EN_, s)
    s = re.sub(r"Bộ luật Lao động số 45/2019/QH14", LABOR_CODE_EN_, s)
    s = re.sub(r"Bộ luật Lao động 2019", LABOR_CODE_EN_, s)
    s = re.sub(r"Luật [Ll]ao động 2019", LABOR_CODE_EN_, s)
    s = re.sub(r"Khoản (\d+)\s+Điều (\d+)", r"Article \2.\1", s, flags=re.IGNORECASE)
    s = re.sub(r"khoản (\d+)\s+Điều (\d+)", r"Article \2.\1", s, flags=re.IGNORECASE)
    s = re.sub(r"Điều (\d+)\s+khoản (\d+)", r"Article \1.\2", s, flags=re.IGNORECASE)
    s = re.sub(r"Điều (\d+)-(\d+)", r"Articles \1-\2", s, flags=re.IGNORECASE)
    s = re.sub(r"Điều (\d+)", r"Article \1", s, flags=re.IGNORECASE)
    s = re.sub(r"Khoản (\d+)", r"Clause \1", s, flags=re.IGNORECASE)
    s = re.sub(r"khoản (\d+)", r"Clause \1", s, flags=re.IGNORECASE)
    s = re.sub(r"Nghị định số (\d+)/(\d+)/NĐ-CP", r"Decree No. \1/\2/ND-CP", s, flags=re.IGNORECASE)
    s = re.sub(r"Nghị định (\d+)/(\d+)/NĐ-CP", r"Decree \1/\2/ND-CP", s, flags=re.IGNORECASE)
    s = re.sub(r"Thông tư (\d+)/(\d+)/TT-BLDTBXH", r"Circular \1/\2/TT-BLDTBXH", s, flags=re.IGNORECASE)
    s = re.sub(r"Luật Giáo dục nghề nghiệp(?: năm 2014| 2014)?", r"Vocational Education Law 2014", s, flags=re.IGNORECASE)
    s = re.sub(r"Luật Người khuyết tật(?: năm 2010| 2010)?", r"Law on Persons with Disabilities 2010", s, flags=re.IGNORECASE)
    s = s.replace("Bộ Công an", "Ministry of Public Security")
    return s


def translate_citations_en_to_vi(ref, item=None):
    """Dịch trích dẫn pháp luật dạng tiếng Anh sang tiếng Việt (ngược chiều trên)."""
    s = ref
    # Trích dẫn RBA -> dùng đúng bản tiếng Việt trong KB của tiêu chí nếu có
    if item and re.match(r"^RBA Code of Conduct", s, re.IGNORECASE):
        return item["rba_citation_vi"]
    s = re.sub(r"Labor Code No\. 45/2019/QH14",
               "Bộ luật Lao động năm 2019 (Luật số 45/2019/QH14)", s)
    s = re.sub(r"Vietnam Labor Code\s*(\d{4})", r"Bộ luật Lao động năm \1", s)
    s = re.sub(r"Vietnam Labor Code", "Bộ luật Lao động", s)
    s = re.sub(r"Labor Code (\d{4})", r"Bộ luật Lao động năm \1", s)
    s = re.sub(r"Articles (\d+)\s*-\s*(\d+)", r"Điều \1-\2", s)
    s = re.sub(r"Article (\d+)\.(\d+)", r"Điều \1 khoản \2", s)
    s = re.sub(r"Article (\d+)", r"Điều \1", s)
    s = re.sub(r"Clause (\d+)", r"khoản \1", s)
    s = re.sub(r"Decree No\. (\d+)/(\d+)/ND-CP", r"Nghị định số \1/\2/NĐ-CP", s)
    s = re.sub(r"Decree (\d+)/(\d+)/ND-CP", r"Nghị định \1/\2/NĐ-CP", s)
    s = re.sub(r"Circular (\d+)/(\d+)/TT-BLDTBXH", r"Thông tư \1/\2/TT-BLDTBXH", s)
    s = s.replace("Ministry of Public Security", "Bộ Công an")
    s = re.sub(r"RBA Code of Conduct,?\s*Version\s*8\.0\.1",
               "Bộ Quy tắc Ứng xử RBA phiên bản 8.0.1", s, flags=re.IGNORECASE)
    s = re.sub(r"\bSection ([A-Z]\d+(?:\.\d+)?)", r"Mục \1", s)
    return s


# Năm trong số hiệu văn bản bị gõ thừa chữ số (vd "13/20203/ND-CP")
_CITATION_YEAR_RE = re.compile(r"(?<=\d)(/)(\d{5,6})(?=/|\b)")


def _plausible_year(y):
    """Chuỗi 4 chữ số có phải năm hợp lý (19xx/20xx)?"""
    return (len(y) == 4 and y[:2] in ("19", "20")
            and 1900 <= int(y) <= 2100)


def normalize_citation(ref):
    """
    Chuẩn hóa nhẹ trích dẫn pháp lý của kiểm toán viên: sửa NĂM trong số hiệu văn
    bản bị gõ lặp phím thừa chữ số (vd "Decree 13/20223/ND-CP" -> "Decree 13/2023/ND-CP",
    "20119" -> "2019", "20225" -> "2025").

    Chỉ sửa khi có cặp chữ số liền kề giống nhau (lỗi lặp phím điển hình) và sau khi
    bỏ 1 chữ số được năm hợp lý (19xx/20xx). Nếu mập mờ (vd "13/20203" không có chữ
    số lặp liền kề) thì GIỮ NGUYÊN văn của KTV - không tự đoán mò để tránh làm sai
    căn cứ pháp lý.
    """
    if not isinstance(ref, str) or not ref.strip():
        return ref

    def _fix(m):
        year = m.group(2)
        if _plausible_year(year):
            return m.group(0)
        # Tìm lỗi gõ lặp phím: hai chữ số liền nhau giống nhau
        for i in range(len(year) - 1):
            if year[i] == year[i + 1]:
                cand = year[:i] + year[i + 1:]
                if _plausible_year(cand):
                    return m.group(1) + cand
        return m.group(0)

    return _CITATION_YEAR_RE.sub(_fix, ref.strip())


def _append_kb_refs_to_block(raw, audit_id):
    """Bổ sung trích dẫn KB vào khối song ngữ đang THIẾU hoàn toàn trích dẫn
    (cả hai phía EN/VI đều rỗng). Nếu khối đã có ít nhất một trích dẫn thì
    trả về nguyên văn."""
    parts = parse_bilingual(raw)
    if parts["en_refs"] or parts["vi_refs"]:
        return raw
    kb_en = kb_refs_for(audit_id, "en")
    kb_vi = kb_refs_for(audit_id, "vi")
    if not kb_en and not kb_vi:
        return raw
    return ensure_section_gap(assemble_block(
        audit_id, parts["en_findings"], kb_en, parts["vi_findings"], kb_vi))


def mirror_refs(en_refs, vi_refs, audit_id=None, fallback_kb=True):
    """
    Đảm bảo trích dẫn pháp lý đủ cả hai ngôn ngữ: phía thiếu được dịch từ phía có sẵn
    (EN<->VI). Nếu không dịch được thành đúng ngôn ngữ (mẫu trích dẫn lạ) thì dùng
    trích dẫn KB của tiêu chí làm dự phòng (chỉ khi fallback_kb=True).

    Khi cả hai phía đã có refs (ví dụ sau bước 4b prepend RBA), dịch bù từng ref VI
    chưa có bản EN tương đương (và ngược lại) để cột EN không bị thiếu.
    Trả về (en_refs, vi_refs).
    """
    item = ITEMS.get(audit_id) if audit_id else None
    if en_refs and not vi_refs:
        mirrored = [translate_citations_en_to_vi(r, item) for r in en_refs]
        if mirrored and all(is_vietnamese(r) for r in mirrored):
            vi_refs = mirrored
        else:
            tr_all = [translate_text(r, "vi") for r in en_refs]
            if tr_all and any(tr_all):
                vi_refs = [t or r for t, r in zip(tr_all, en_refs)]
            elif item and fallback_kb:
                vi_refs = [item["rba_citation_vi"]] + list(item["vn_citations_vi"])
    elif vi_refs and not en_refs:
        mirrored = [translate_citations_vi_to_en(r) for r in vi_refs]
        if mirrored and not any(is_vietnamese(r) for r in mirrored):
            en_refs = mirrored
        else:
            tr_all = [translate_text(r, "en") for r in vi_refs]
            if tr_all and any(tr_all):
                en_refs = [t or r for t, r in zip(tr_all, vi_refs)]
            elif item and fallback_kb:
                en_refs = [item["rba_citation_en"]] + list(item["vn_citations_en"])
    else:
        # Cả hai phía đều có refs: dịch bù từng ref VI chưa có bản EN tương đương
        # (và ngược lại) để tránh bất đối xứng khi 4b prepend RBA vào EN nhưng
        # ref KTV chỉ ghi trong VI (hoặc ngược lại).
        en_list = list(en_refs)
        vi_list = list(vi_refs)
        for r in vi_refs:
            tr = translate_citations_vi_to_en(r)
            if tr and not is_vietnamese(tr) and not pool_covers(tr, en_list):
                en_list.append(tr)
        for r in en_refs:
            tr = translate_citations_en_to_vi(r, item)
            if tr and is_vietnamese(tr) and not pool_covers(tr, vi_list):
                vi_list.append(tr)
        en_refs, vi_refs = en_list, vi_list
    return en_refs, vi_refs


def mirror_findings(en_findings, vi_findings):
    """
    Đảm bảo findings đủ cả hai ngôn ngữ: phía thiếu được dịch tự động từ phía có
    (VI<->EN, từ điển chuyên ngành ngoại tuyến). Trả về (en_findings, vi_findings,
    translated) — translated=True khi đã áp dụng dịch tự động.
    """
    if vi_findings and not en_findings:
        return translate_findings(vi_findings, "en"), list(vi_findings), True
    if en_findings and not vi_findings:
        return list(en_findings), translate_findings(en_findings, "vi"), True
    return list(en_findings), list(vi_findings), False


def mirror_each_finding(en_findings, vi_findings):
    """
    Bổ sung bản dịch cho TỪNG mục finding còn thiếu bản ngôn ngữ kia (không chỉ khi
    cả một phía trống).

    Ghép cặp theo thứ tự: `min(len(en), len(vi))` mục đầu của hai phía coi như đã
    tương ứng nên không dịch thêm; phần dư ở mỗi phía được dịch sang ngôn ngữ kia.
    Nhờ vậy khối EN và khối VI của cột L cùng thể hiện đủ số mục finding như file đầu vào.

    Trả về (en_findings, vi_findings).
    """
    orig_en = list(en_findings or [])
    orig_vi = list(vi_findings or [])
    paired = min(len(orig_en), len(orig_vi))
    en = list(orig_en)
    vi = list(orig_vi)
    for item in orig_vi[paired:]:
        tr = translate_text(item, "en")
        if tr and tr.strip():
            en.append(tr.strip())
    for item in orig_en[paired:]:
        tr = translate_text(item, "vi")
        if tr and tr.strip():
            vi.append(tr.strip())
    return en, vi


# ------------------------------------------------- bảo đảm thuần một ngôn ngữ ---
def _is_impure(text, target):
    """Mục còn lẫn ngôn ngữ kia? (target='en' -> còn tiếng Việt; 'vi' -> còn tiếng Anh)

    Khối EN: chỉ coi là lẫn khi còn từ tiếng Việt CÓ DẤU hoặc Vieglish
    (đoạn Việt lẫn từ Anh) -> mục tiếng Việt KHÔNG DẤU nguyên câu KHÔNG bị
    coi là lẫn (giữ nguyên ở EN cho đúng thứ tự mục gốc của KTV; hậu kiểm
    scan_language_mixing cũng dùng has_vietnamese_words nên nhất quán).
    """
    if not text or not isinstance(text, str):
        return False
    if target == "en":
        return bool(has_vietnamese_words(text))
    return has_english_words(text)


def purify_item_to_language(item, target):
    """
    Đưa một mục finding về THUẦN một ngôn ngữ (`target`).
    Thứ tự xử lý: (1) dịch cả câu (Google online nếu đang bật, từ điển chuyên
    ngành ngoại tuyến nếu không có mạng); (2) dịch riêng các đoạn ngôn ngữ lạc.

    Trả về None nếu KHÔNG đạt bản thuần -> caller loại mục khỏi khối sai ngôn ngữ
    (tuyệt đối không xuất bản bản lẫn Anh/Việt ra báo cáo).
    """
    if not isinstance(item, str) or not item.strip():
        return item
    if not _is_impure(item, target):
        return item
    whole = translate_text(item, target)
    if whole and whole.strip() and not _is_impure(whole, target):
        return whole.strip()
    repaired = repair_mixed_item(item, target)
    if repaired and repaired.strip() and not _is_impure(repaired, target):
        return repaired.strip()
    return None


def enforce_language_purity(en_findings, vi_findings, en_refs, vi_refs, audit_id=None):
    """
    Cổng chốt cuối cùng trước khi ghi ô L: bảo đảm
      - khối EN ('AUDIT FINDINGS (ENGLISH)' + 'Legal references:') thuần tiếng Anh;
      - khối VI ('NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT)' + 'Căn cứ pháp lý:') thuần tiếng Việt.

    Mục lẫn ngôn ngữ được dịch lại cho thuần (online/offline); nếu vẫn không đạt
    thì bị loại khỏi khối SAI ngôn ngữ - bản đúng ngôn ngữ của kiểm toán viên vẫn
    được giữ nguyên nên không mất nội dung gốc.

    Trả về (en_findings, vi_findings, en_refs, vi_refs) đã sạch ngôn ngữ.
    """
    item = ITEMS.get(audit_id) if audit_id else None

    en_f, vi_f = [], []
    for t in list(en_findings or []):
        p = purify_item_to_language(t, "en")
        if p:
            en_f.append(p)
    for t in list(vi_findings or []):
        p = purify_item_to_language(t, "vi")
        if p:
            vi_f.append(p)

    en_r, vi_r = [], []
    # Trích dẫn EN còn tiếng Việt -> dịch sang dạng trích dẫn tiếng Anh; nếu không
    # dịch được thì chuyển hẳn về khối tiếng Việt (không để lẫn trong khối EN).
    for r in list(en_refs or []):
        if not has_vietnamese_words(r):
            en_r.append(r)
            continue
        tr = translate_citations_vi_to_en(r)
        if not has_vietnamese_words(tr):
            en_r.append(tr)
        else:
            vi_r.append(r)
    # Trích dẫn VI còn tiếng Anh -> dịch sang dạng trích dẫn tiếng Việt; nếu không
    # dịch được thì chuyển hẳn về khối tiếng Anh.
    for r in list(vi_refs or []):
        if not has_english_words(r):
            vi_r.append(r)
            continue
        tr = translate_citations_en_to_vi(r, item)
        if not has_english_words(tr):
            vi_r.append(tr)
        else:
            en_r.append(r)
    return en_f, vi_f, dedupe_items(en_r), dedupe_items(vi_r)


def assemble_block(audit_id, en_findings, en_refs, vi_findings, vi_refs):
    """
    Dựng khối song ngữ chuẩn (4 header mẫu) SAU KHI đã bảo đảm khối EN thuần Anh
    và khối VI thuần Việt - mọi đường ghi ô L đều đi qua đây.
    """
    en_f, vi_f, en_r, vi_r = enforce_language_purity(
        en_findings, vi_findings, en_refs, vi_refs, audit_id)
    return structure_bilingual(en_f, en_r, vi_f, vi_r)


def scan_language_mixing(text):
    """
    Quét một ô cột L và trả về danh sách vấn đề:
        [(kind, content), ...]
    kind = 'format' : ô không có khối song ngữ chuẩn (sai định dạng mẫu);
           'en'     : khối EN còn sót tiếng Việt;
           'vi'     : khối VI còn sót tiếng Anh.
    """
    out = []
    if not isinstance(text, str) or not text.strip():
        return out
    if HEADER_EN not in text or HEADER_VI not in text:
        first = text.strip().split("\n")[0]
        out.append(("format", first[:160]))
        return out
    before, after = text.split(HEADER_VI, 1)
    en_body = before.split(HEADER_EN, 1)[1]
    for block, headers, section, checker in (
            (en_body, (REFS_EN,), "en", has_vietnamese_words),
            (after, (REFS_VI,), "vi", has_english_words)):
        for line in block.replace("\r\n", "\n").split("\n"):
            s = line.strip()
            if not s or s in headers or s == SEPARATOR:
                continue
            s = re.sub(r"^(?:\d+\s*[.:)]|-)\s*", "", s).strip()
            if not s:
                continue
            if checker(s):
                out.append((section, s[:200]))
    return out


def ensure_section_gap(text):
    """
    Đảm bảo có dòng phân cách (---------------------------------------) giữa phần EN và VI.
    Chuẩn hóa mềm: nếu đã có SEPARATOR trước HEADER_VI thì giữ nguyên, nếu chỉ có
    khoảng trống hoặc không có gì thì chèn thêm separator.
    """
    if not isinstance(text, str) or HEADER_VI not in text:
        return text
    before, after = text.split(HEADER_VI, 1)
    before = before.rstrip("\n")
    # Nếu đã có separator thì giữ nguyên
    if before.endswith(SEPARATOR):
        return text
    return f"{before}\n{SEPARATOR}\n{HEADER_VI}{after}"


def normalize_bilingual_block(audit_id, text, add_kb_refs=False):
    """
    Chuẩn hóa lại khối song ngữ đã có sẵn (chứa HEADER_EN + HEADER_VI):
    - Tách các mục lẫn ngôn ngữ (Vieglish): EN thuần Anh / VI thuần Việt;
    - Tách trích dẫn pháp lý theo ngôn ngữ;
    - Một phía findings thiếu -> tự dịch bù (offline/online);
    - Đảm bảo đúng 1 hàng trống ngăn cách phần EN và VI.

    add_kb_refs=False (mặc định): KHÔNG tự bổ sung trích dẫn RBA/luật từ KB,
    chỉ dịch bù giữa hai phía EN<->VI đã có. True: có dùng KB làm dự phòng.
    Khối placeholder/draft thường (không có đủ hai header) trả về nguyên văn.
    """
    if not isinstance(text, str) or HEADER_EN not in text or HEADER_VI not in text:
        return text
    parts = parse_bilingual(text)
    en_f, en_r = parts["en_findings"], parts["en_refs"]
    vi_f, vi_r = parts["vi_findings"], parts["vi_refs"]
    en_f, vi_f = purify_findings(en_f, vi_f)
    _ef, en_r, _vf, vi_r = separate_languages([], en_r, [], vi_r)
    en_r = [purify_ref(r, "en") for r in en_r]
    vi_r = [purify_ref(r, "vi") for r in vi_r]
    en_f, vi_f = mirror_each_finding(en_f, vi_f)
    item = ITEMS.get(audit_id) if audit_id else None
    # Cả hai phía đều không có trích dẫn -> dùng trích dẫn KB (nếu cho phép)
    if not en_r and not vi_r and add_kb_refs:
        kb_en = kb_refs_for(audit_id, "en")
        kb_vi = kb_refs_for(audit_id, "vi")
        if kb_en or kb_vi:
            en_r, vi_r = kb_en, kb_vi
    elif add_kb_refs and item:
        # Prepend RBA citation nếu chưa có
        rba_en = item.get("rba_citation_en")
        rba_vi = item.get("rba_citation_vi")
        has_rba_en = any(re.search(r"RBA Code of Conduct", r, re.IGNORECASE) for r in en_r)
        has_rba_vi = any(re.search(r"Quy tắc Ứng xử RBA", r, re.IGNORECASE) for r in vi_r)
        if not has_rba_en and rba_en:
            en_r = [rba_en] + list(en_r)
        if not has_rba_vi and rba_vi:
            vi_r = [rba_vi] + list(vi_r)
    # Luôn dịch bù trích dẫn giữa hai ngôn ngữ (nếu một phía thiếu); chỉ khi
    # add_kb_refs=True mới được dùng trích dẫn KB làm dự phòng nếu dịch không đạt.
    en_r, vi_r = mirror_refs(en_r, vi_r, audit_id, fallback_kb=add_kb_refs)
    return ensure_section_gap(
        assemble_block(audit_id, en_f, en_r, vi_f, vi_r))


# === ĐẦY ĐỦ NỘI DUNG: không được rơi nội dung cột L của file đầu vào ===
# Ngưỡng coi hai đoạn là "cùng nội dung" (dùng cho hậu kiểm đầy đủ + gộp curated)
_COVER_MIN = 0.6
# Ngưỡng xác định bản curated có biên tập đúng nội dung file đầu vào hay không
_CURATED_MATCH_MIN = 0.5
# Ngưỡng coi hai trích dẫn pháp lý là cùng một quy định (để dùng bản chuẩn của KB)
_REF_MATCH_MIN = 0.5


def _norm_item(text):
    """Chuẩn hóa mục để so khớp: bỏ marker số/gạch đầu dòng, gộp khoảng trắng."""
    s = re.sub(r"^\s*(?:\d+\s*[.:)]|-)\s*", "", (text or "").strip())
    return re.sub(r"\s+", " ", s).strip()


def pool_covers(item, pool, threshold=_COVER_MIN):
    """Mục `item` đã có (gần đủ) trong danh sách `pool` chưa?"""
    n = _norm_item(item).lower()
    if not n:
        return True
    for u in pool or []:
        nu = _norm_item(u).lower()
        if not nu:
            continue
        if n in nu or nu in n or _overlap(n, nu) >= threshold:
            return True
    return False


def missing_fragments(item, pool, threshold=_COVER_MIN):
    """
    Các MỆNH ĐỀ của `item` chưa xuất hiện trong `pool` (so khớp mờ >= threshold).
    Giữ NGUYÊN từ ngữ của kiểm toán viên (chỉ bỏ những mệnh đề đã có) - dùng để
    chèn lại phần nội dung bị rơi khi trình bày lại khối song ngữ.
    Trả về [] nếu mục đã có đủ trong pool.
    """
    text = _norm_item(item)
    if not text:
        return []
    if pool_covers(text, pool, threshold):
        return []
    frags = [f.strip() for f in re.split(r"(?<=[.;:!?])\s+", text) if f.strip()]
    if len(frags) <= 1:
        return [text]
    return [f for f in frags if not pool_covers(f, pool, threshold)]


def supplement_pools(draft_en, draft_vi, en_items, vi_items):
    """
    Bổ sung vào khối (en_items/vi_items) mọi mục finding của file đầu vào mà khối
    đang THIẾU (giữ nguyên từ ngữ kiểm toán viên); sau đó dịch bù sang ngôn ngữ
    còn lại nếu bản dịch đạt thuần ngôn ngữ đó. Trả về (en_items, vi_items) mới.
    """
    en, vi = list(en_items or []), list(vi_items or [])
    added = []
    for lang, drafts in (("vi", draft_vi or []), ("en", draft_en or [])):
        for t in drafts:
            frags = missing_fragments(t, vi if lang == "vi" else en)
            if not frags:
                continue
            item = " ".join(frags)
            fixed = purify_item_to_language(item, lang) or item
            (vi if lang == "vi" else en).append(fixed)
            added.append((fixed, lang))
    for item, lang in added:
        other = "en" if lang == "vi" else "vi"
        target = en if other == "en" else vi
        tr = translate_text(item, other)
        if not tr or not tr.strip():
            continue
        if _is_impure(tr, other) or pool_covers(tr, target):
            continue
        target.append(tr.strip())
    return en, vi


def curated_covers_draft(parts, draft_en, draft_vi):
    """
    Bản curated (đã parse thành khối song ngữ) có biên tập ĐÚNG nội dung cột L của
    file đầu vào không? -> Ít nhất một mục của file đầu vào được bản curated đề cập
    (>= _CURATED_MATCH_MIN từ trùng).

    Dùng để KHÔNG áp nội dung curated của nhà máy/kỳ audit khác vào file đang mở,
    và KHÔNG bịa finding cho ô L trống.
    """
    en_pool = list(parts.get("en_findings") or [])
    vi_pool = list(parts.get("vi_findings") or [])
    both = en_pool + vi_pool
    for items, lang in ((draft_en or [], "en"), (draft_vi or [], "vi")):
        pool = (en_pool if lang == "en" else vi_pool) or both
        if not pool:
            continue
        for t in items:
            if not _norm_item(t):
                continue
            if pool_covers(t, pool, _CURATED_MATCH_MIN):
                return True
    return False


def _extract_article_numbers(ref):
    """Trích xuất các số điều/khoản từ trích dẫn để so sánh chính xác hơn."""
    nums = re.findall(r"\b(?:điều|article|khoản|clause|section|mục)\s+(\d+)", ref, re.IGNORECASE)
    compound = re.findall(r"\b(\d+\.\d+)\b", ref)
    return set(nums + compound)


def _best_kb_ref_match(ref, kb_refs):
    """Chỉ số trích dẫn KB khớp nhất với `ref` (trùng từ cao nhất) hoặc None.

    Bổ sung: nếu ref KTV có số điều/khoản cụ thể, ít nhất một số điều đó phải
    xuất hiện trong trích dẫn KB (tránh nhầm Điều 45 với Điều 14 vì cùng bộ luật).
    """
    target = _norm_item(ref).lower()
    ref_arts = _extract_article_numbers(ref)
    best, best_score = None, 0.0
    for i, k in enumerate(kb_refs or []):
        score = _overlap(target, _norm_item(k).lower())
        if score > best_score:
            kb_arts = _extract_article_numbers(k)
            if ref_arts and kb_arts and ref_arts.isdisjoint(kb_arts):
                continue
            best, best_score = i, score
    return best if best_score >= _REF_MATCH_MIN else None


def standardize_refs_with_kb(ref_en, ref_vi, audit_id):
    """
    Chuẩn hóa trích dẫn pháp lý của kiểm toán viên theo KB của tiêu chí: trích dẫn
    nói về CÙNG một quy định đã có trong KB (trùng >= _REF_MATCH_MIN từ) được thay
    bằng bản chuẩn của KB (đúng số hiệu/điều khoản, kèm bản ngôn ngữ còn lại).

    Trích dẫn không khớp KB được GIỮ NGUYÊN (không bỏ sót căn cứ KTV đã nêu).
    Trả về (en_refs, vi_refs).
    """
    pools = {"en": [], "vi": []}
    for lang in ("en", "vi"):
        refs = ref_en if lang == "en" else ref_vi
        other_lang = "vi" if lang == "en" else "en"
        kb = kb_refs_for(audit_id, lang)
        kb_other = kb_refs_for(audit_id, other_lang)
        for r in refs or []:
            idx = _best_kb_ref_match(r, kb)
            if idx is None:
                if not pool_covers(r, pools[lang]):
                    pools[lang].append(r)
                continue
            if not pool_covers(kb[idx], pools[lang]):
                pools[lang].append(kb[idx])
            if idx < len(kb_other) and not pool_covers(kb_other[idx],
                                                      pools[other_lang]):
                pools[other_lang].append(kb_other[idx])
    return pools["en"], pools["vi"]


def merge_ref_pools(base_en, base_vi, add_en, add_vi, item=None):
    """
    Gộp trích dẫn pháp lý của file đầu vào vào khối đã có (không bỏ sót căn cứ mà
    kiểm toán viên đã nêu): mục mới được thêm vào đúng ngôn ngữ, bản ngôn ngữ kia
    được dịch bù (nếu dịch được), trùng lặp bị loại.
    """
    en, vi = list(base_en or []), list(base_vi or [])
    # Xử lý trích dẫn TIẾNG VIỆT trước để bản gốc của KTV (và bản chuẩn KB tiếng
    # Việt) luôn được ưu tiên; bản dịch bù sang ngôn ngữ kia chỉ được thêm khi
    # chưa có trích dẫn tương đương -> tránh ghi đè bằng bản dịch máy thiếu dấu.
    for r in add_vi or []:
        if not r or pool_covers(r, vi):
            continue
        vi.append(r)
        tr = translate_citations_vi_to_en(r)
        if tr and tr.strip() and not is_vietnamese(tr) and not pool_covers(tr, en):
            en.append(tr.strip())
    for r in add_en or []:
        if not r or pool_covers(r, en):
            continue
        en.append(r)
        tr = translate_citations_en_to_vi(r, item)
        if tr and tr.strip() and is_vietnamese(tr) and not pool_covers(tr, vi):
            vi.append(tr.strip())
    return en, vi


def supplement_missing_findings(block, draft_en, draft_vi):
    """
    Chốt chặn ĐẦY ĐỦ: dựng lại khối (nếu cần) sao cho mọi mục finding của file đầu
    vào đều có mặt trong cột L. Mục bị rơi trong quá trình tách/dịch ngôn ngữ được
    chèn lại vào đúng phần ngôn ngữ của nó, giữ nguyên từ ngữ kiểm toán viên.

    Trả về (block, số mục đã chèn lại).
    """
    if not isinstance(block, str) or HEADER_EN not in block or HEADER_VI not in block:
        return block, 0
    parts = parse_bilingual(block)
    en_f, vi_f = list(parts["en_findings"]), list(parts["vi_findings"])
    en_f2, vi_f2 = supplement_pools(draft_en, draft_vi, en_f, vi_f)
    added = (len(en_f2) - len(en_f)) + (len(vi_f2) - len(vi_f))
    if not added:
        return block, 0
    # Dựng lại khối bằng structure_bilingual (KHÔNG lọc lại ngôn ngữ) để mục vừa
    # chèn lại không bị loại một lần nữa -> nội dung cột L đầy đủ.
    return ensure_section_gap(
        structure_bilingual(en_f2, parts["en_refs"], vi_f2, parts["vi_refs"])), added


def missing_input_items(block, draft_text):
    """
    Hậu kiểm "cột L đầy đủ": danh sách mục finding của nội dung file đầu vào KHÔNG
    xuất hiện trong khối đã ghi (dùng cho CLI/GUI báo rõ dòng cần rà lại).
    """
    d = parse_draft(draft_text)
    draft_en, draft_vi = purify_findings(d.items_en, d.items_vi)
    if not draft_en and not draft_vi:
        return []
    if not isinstance(block, str) or HEADER_EN not in block or HEADER_VI not in block:
        return list(draft_en) + list(draft_vi)
    parts = parse_bilingual(block)
    out = []
    for t in draft_vi:
        if missing_fragments(t, parts["vi_findings"]):
            out.append(t)
    for t in draft_en:
        if missing_fragments(t, parts["en_findings"]):
            out.append(t)
    return out


# === ELABORATION: cau phan tich RBA (_Main in elaborations.py) ===


def _join_sentence(fact, add):
    """Nối câu diễn giải vào sau finding gốc với dấu chấm câu chuẩn xác."""
    f = (fact or "").rstrip()
    a = (add or "").strip()
    if not a:
        return f
    if not f:
        return a
    if f[-1] in ".!?:;":
        return f + " " + a
    return f + ". " + a


def elaborate_findings(audit_id, en_items, vi_items):
    """Bo sung cau phan tich RBA cho finding NGAN (giu nguyen su that goc)."""
    out_en = list(en_items)
    out_vi = list(vi_items)
    key = _elaborate_key(audit_id, ITEMS.get) or "general"
    elab = ELABORATIONS_FULL.get(key, ELABORATIONS_FULL.get("general", {}))
    if out_en:
        fact = out_en[0]
        add = elab.get("en", "")
        if fact and add and len(fact) < _ELABORATE_MIN_LEN and add not in fact:
            out_en = [_join_sentence(fact, add)] + out_en[1:]
    if out_vi:
        fact = out_vi[0]
        add = elab.get("vi", "")
        if fact and add and len(fact) < _ELABORATE_MIN_LEN and add not in fact:
            out_vi = [_join_sentence(fact, add)] + out_vi[1:]
    return out_en, out_vi


def build_content(audit_id, draft_text, rating=None, curated=None, use_curated=True,
                  add_kb_refs=True):
    """
    Xây dựng nội dung hoàn chỉnh của cột L cho một dòng.

    NGUYÊN TẮC NỘI DUNG:
    1. NỘI DUNG FILE ĐẦU VÀO LÀ GỐC: đọc nội dung cột L của file đầu vào. Với ô CÓ
       nội dung, app BIÊN TẬP lại chính nội dung đó theo chuẩn RBA + Luật Lao động
       VN (tách song ngữ, dịch bù EN<->VI, bổ sung câu phân tích RBA, bổ sung căn
       cứ pháp lý) rồi trình bày theo khối 4 header in đậm.
    2. KHÔNG LÀM RƠI NỘI DUNG: mọi mục finding/trích dẫn của file đầu vào đều phải
       có mặt trong khối cuối (chốt chặn supplement_missing_findings + hậu kiểm
       missing_input_items).
    3. Bản curated (curated_findings.json) chỉ là BẢN BIÊN TẬP của chính nội dung
       file đầu vào: chỉ dùng khi file đầu vào CÓ nội dung và bản curated thật sự
       đề cập nội dung đó; phần curated bỏ sót vẫn được chèn lại đầy đủ. Ô L trống
       -> KHÔNG bịa finding (chỉ khối rỗng + trích dẫn KB của tiêu chí).
    4. Trình bày theo đúng khối mẫu, phần EN thuần tiếng Anh / phần VI thuần
       tiếng Việt:
        AUDIT FINDINGS (ENGLISH) / 1. ... / Legal references: / 1. ...
        ---------------------------------------
        NỘI DUNG ĐÁNH GIÁ (TIẾNG VIỆT) / 1. ... / Căn cứ pháp lý: / 1. ...
    """
    if curated is None:
        curated = load_curated()
    item = ITEMS.get(audit_id)
    rating = (rating or "").strip()

    # --- (1) Nội dung GỐC đọc từ cột L của file đầu vào ---
    d = parse_draft(draft_text)
    if isinstance(draft_text, str) and HEADER_EN in draft_text and HEADER_VI in draft_text:
        # Ô L đã là khối song ngữ (do app phiên bản trước sinh ra) -> đọc nội
        # dung theo đúng hai phần ngôn ngữ của khối
        parts0 = parse_bilingual(draft_text)
        draft_en, draft_vi = purify_findings(parts0["en_findings"],
                                             parts0["vi_findings"])
        ref_en, ref_vi = list(parts0["en_refs"]), list(parts0["vi_refs"])
    else:
        draft_en, draft_vi = purify_findings(d.items_en, d.items_vi)
        ref_en, ref_vi = split_refs(d.refs_raw)
    ref_en = [normalize_citation(r) for r in ref_en]
    ref_vi = [normalize_citation(r) for r in ref_vi]
    # Chuẩn hóa trích dẫn của KTV theo KB của tiêu chí (đúng số hiệu/điều khoản,
    # đủ cả hai ngôn ngữ); trích dẫn không khớp KB được giữ nguyên.
    if add_kb_refs and item:
        ref_en, ref_vi = standardize_refs_with_kb(ref_en, ref_vi, audit_id)

    cur = curated.get(audit_id) if use_curated else None
    cur_parts = None
    if cur and cur.get("raw"):
        raw = cur["raw"]
        if HEADER_EN in raw and HEADER_VI in raw:
            cur_parts = parse_bilingual(raw)
        elif d.has_content:
            # Bản đã duyệt dạng thô (không phải khối song ngữ): dùng nguyên văn,
            # nhưng chỉ khi ô L của file đầu vào CÓ nội dung (không bịa finding).
            return ensure_section_gap(raw)

    # --- (2) Nội dung đánh giá: draft gốc (mặc định) hoặc bản curated là bản
    #         biên tập của chính draft đó (khi curated thật sự đề cập nội dung) ---
    if cur_parts and curated_covers_draft(cur_parts, draft_en, draft_vi):
        en_f = list(cur_parts["en_findings"])
        vi_f = list(cur_parts["vi_findings"])
        en_r = [normalize_citation(r) for r in cur_parts["en_refs"]]
        vi_r = [normalize_citation(r) for r in cur_parts["vi_refs"]]
        # ĐẦY ĐỦ: chèn lại mục/trích dẫn của file đầu vào mà bản curated bỏ sót
        en_f, vi_f = supplement_pools(draft_en, draft_vi, en_f, vi_f)
        en_r, vi_r = merge_ref_pools(en_r, vi_r, ref_en, ref_vi, item)
    else:
        en_f, vi_f = list(draft_en), list(draft_vi)
        en_r, vi_r = list(ref_en), list(ref_vi)

    if not d.has_content:
        # Ô L TRỐNG: khối rỗng + trích dẫn KB của tiêu chí (4 header in đậm).
        # Không dùng curated -> không bịa nội dung đánh giá cho dòng chưa có draft.
        if item:
            en_r = [item["rba_citation_en"]] + list(item["vn_citations_en"])
            vi_r = [item["rba_citation_vi"]] + list(item["vn_citations_vi"])
        return ensure_section_gap(assemble_block(audit_id, [], en_r, [], vi_r))

    # Ô L đã là khối song ngữ của phiên bản app trước -> chuẩn hóa lại khối đó
    # rồi vẫn đi qua chốt chặn đầy đủ ở bước (5).
    if HEADER_EN in draft_text and HEADER_VI in draft_text:
        return supplement_missing_findings(
            normalize_bilingual_block(audit_id, draft_text, add_kb_refs=add_kb_refs),
            draft_en, draft_vi)[0]

    # --- (3) Trích dẫn pháp lý: dịch bù giữa hai ngôn ngữ ---
    _ef, en_r, _vf, vi_r = separate_languages([], en_r, [], vi_r)
    en_r = [purify_ref(r, "en") for r in en_r]
    vi_r = [purify_ref(r, "vi") for r in vi_r]

    # --- (4a) Cả hai phía đều không có trích dẫn -> dùng trích dẫn KB (dòng vi
    #          phạm) để ô L luôn có căn cứ pháp lý ---
    if not en_r and not vi_r and add_kb_refs and item and rating != "Conformance":
        en_r = kb_refs_for(audit_id, "en")
        vi_r = kb_refs_for(audit_id, "vi")
    # --- (4b) KTV đã có refs nhưng chưa có RBA citation -> prepend RBA của tiêu chí
    #          (1 RBA + giữ đúng luật KTV ghi, KHÔNG nhồi thêm Luật VN từ KB) ---
    elif add_kb_refs and item and rating != "Conformance":
        rba_en = item["rba_citation_en"]
        rba_vi = item["rba_citation_vi"]
        has_rba_en = any(
            re.search(r"RBA Code of Conduct", r, re.IGNORECASE) for r in en_r
        )
        has_rba_vi = any(
            re.search(r"Quy tắc Ứng xử RBA", r, re.IGNORECASE) for r in vi_r
        )
        if not has_rba_en and rba_en:
            en_r = [rba_en] + list(en_r)
        if not has_rba_vi and rba_vi:
            vi_r = [rba_vi] + list(vi_r)
    en_f, vi_f = mirror_each_finding(en_f, vi_f)
    en_r, vi_r = mirror_refs(en_r, vi_r, audit_id)

    # Dòng VI PHẠM + finding ngắn (draft mới từ file import bất kỳ):
    # bổ sung câu phân tích RBA theo văn phong file mẫu (giữ sự thật gốc,
    # chỉ elaborate finding ĐẦU TIÊN của mỗi phía -- các mục 2,3,... giữ nguyên
    # để không làm mất nội dung kiểm toán viên).
    if rating in NON_CONFORMANCE_RATINGS:
        e1, v1 = elaborate_findings(audit_id, en_f[:1], vi_f[:1])
        en_f = e1 + en_f[1:]
        vi_f = v1 + vi_f[1:]

    # --- (5) Dựng khối chuẩn + CHỐT CHẶN ĐẦY ĐỦ: chèn lại mọi mục của file đầu
    #         vào bị rơi trong quá trình biên tập/dịch (cột L không được thiếu) ---
    block = ensure_section_gap(
        assemble_block(audit_id, en_f, en_r, vi_f, vi_r))
    block, _added = supplement_missing_findings(block, draft_en, draft_vi)
    return block

