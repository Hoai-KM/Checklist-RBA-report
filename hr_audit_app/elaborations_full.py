# -*- coding: utf-8 -*-
"""Gop cau phan tich day du tu elab_a/b/c (co dau tieng Viet)."""
try:
    from elab_a import PART_A
except ImportError:
    try:
        from .elab_a import PART_A
    except ImportError:
        PART_A = {}
try:
    from elab_b import PART_B
except ImportError:
    try:
        from .elab_b import PART_B
    except ImportError:
        PART_B = {}
try:
    from elab_c import PART_C
except ImportError:
    try:
        from .elab_c import PART_C
    except ImportError:
        PART_C = {}
ELABORATIONS_FULL = {}
ELABORATIONS_FULL.update(PART_A)
ELABORATIONS_FULL.update(PART_B)
ELABORATIONS_FULL.update(PART_C)

ELABORATIONS_FULL["disability"] = {
    "en": "The facility must ensure reasonable accommodation and effective communication channels for workers with disabilities in accordance with RBA standards and applicable regulations.",
    "vi": "Cơ sở cần bảo đảm các biện pháp hỗ trợ hợp lý và kênh tiếp nhận ý kiến hiệu quả cho lao động khuyết tật theo tiêu chuẩn RBA và quy định pháp luật.",
}
ELABORATIONS_FULL["religion"] = {
    "en": "The facility must provide reasonable accommodation for religious practices in accordance with RBA standards and applicable regulations.",
    "vi": "Cơ sở cần tạo điều kiện thuận lợi và bố trí không gian phù hợp cho nhu cầu thực hành tín ngưỡng, tôn giáo theo tiêu chuẩn RBA và quy định pháp luật.",
}
ELABORATIONS_FULL["discipline"] = {
    "en": "Disciplinary procedures must be clearly documented, transparent, and comply strictly with legal procedures and RBA requirements.",
    "vi": "Quy trình xử lý kỷ luật lao động phải được lập thành văn bản rõ ràng, minh bạch và tuân thủ đúng trình tự pháp luật và yêu cầu RBA.",
}
