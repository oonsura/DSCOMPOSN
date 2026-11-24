"""
คำนวณระดับชั้น +1 ถึง +5
รองรับ ป.1-ป.6, ม.1-ม.6 และ ปวช.1-ปวช.3
ปีที่เกิน ม.6 หรือ ปวช.3 = "หมดสิทธิ์"
"""
import re

EXPIRED_TEXT = 'หมดสิทธิ์'

# ลำดับการเลื่อนชั้น: ป.1 -> ... -> ป.6 -> ม.1 -> ... -> ม.6 -> หมดสิทธิ์
#                     ปวช.1 -> ปวช.2 -> ปวช.3 -> หมดสิทธิ์
SEQUENCES = {
    'ป.': ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6', 'ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'],
    'ม.': ['ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'],
    'ปวช.': ['ปวช.1', 'ปวช.2', 'ปวช.3'],
}

_PATTERNS = [
    (re.compile(r'^(?:ปวช|ประกาศนียบัตรวิชาชีพ(?:ชั้นปีที่)?)\.?\s*(\d)$'), 'ปวช.'),
    (re.compile(r'^(?:ม|มัธยม(?:ศึกษา)?(?:ปีที่)?|m)\.?\s*(\d)$', re.I), 'ม.'),
    (re.compile(r'^(?:ป|ประถม(?:ศึกษา)?(?:ปีที่)?|p)\.?\s*(\d)$', re.I), 'ป.'),
]


def normalize_level(text):
    """แปลงข้อความระดับชั้นให้เป็นรูปแบบมาตรฐาน เช่น 'ม 4', 'มัธยมศึกษาปีที่ 4' -> 'ม.4'
    คืนค่า None ถ้าไม่รู้จัก"""
    if text is None:
        return None
    s = str(text).strip().replace(' ', '')
    if not s:
        return None
    if s == EXPIRED_TEXT:
        return EXPIRED_TEXT
    for pattern, prefix in _PATTERNS:
        m = pattern.match(s)
        if m:
            level = f'{prefix}{m.group(1)}'
            if level in SEQUENCES[prefix]:
                return level
            return None
    return None


def next_levels(current_level, years=5):
    """คืนค่ารายการระดับชั้นในอีก 1..years ปีข้างหน้า"""
    level = normalize_level(current_level)
    if level is None:
        return [''] * years
    if level == EXPIRED_TEXT:
        return [EXPIRED_TEXT] * years
    prefix = 'ปวช.' if level.startswith('ปวช.') else level[:2]
    seq = SEQUENCES[prefix]
    idx = seq.index(level)
    result = []
    for i in range(1, years + 1):
        j = idx + i
        result.append(seq[j] if j < len(seq) else EXPIRED_TEXT)
    return result


def is_within_limit(level):
    """ระดับชั้นไม่เกิน ม.6 / ปวช.3 (และต้องเป็นระดับที่รู้จัก)"""
    lv = normalize_level(level)
    return lv is not None and lv != EXPIRED_TEXT


def apply_levels(user):
    """คำนวณและใส่ค่า year_plus_1..5 ให้ user (ยังไม่ save) คืนค่า True ถ้าคำนวณได้"""
    norm = normalize_level(user.current_level)
    if norm and norm != EXPIRED_TEXT:
        user.current_level = norm
    plus = next_levels(user.current_level)
    user.year_plus_1, user.year_plus_2, user.year_plus_3, user.year_plus_4, user.year_plus_5 = plus
    return norm is not None
