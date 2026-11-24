from django import template
from django.utils import timezone
from django.utils.html import format_html
from datetime import date, datetime

register = template.Library()

THAI_MONTHS = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน',
    'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม',
    'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม',
]
THAI_MONTHS_SHORT = [
    '', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.',
    'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.',
    'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.',
]
THAI_DAYS = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
THAI_DAYS_SHORT = ['จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.', 'อา.']


def _to_ce_date_safe(d):
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    if isinstance(d, date) and d.year >= 2400:
        try:
            return d.replace(year=d.year - 543)
        except ValueError:
            return date(d.year - 543, d.month, min(d.day, 28))
    return d


def get_thai_day_short(weekday_int):
    return THAI_DAYS_SHORT[weekday_int] if 0 <= weekday_int <= 6 else ''


def format_date_thai_short_weekday(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    elif not isinstance(d, date):
        return str(d)
    ce = _to_ce_date_safe(d)
    day_s = get_thai_day_short(ce.weekday())
    year_be = d.year + (543 if d.year < 2400 else 0)
    return f'{day_s} {d.day} {THAI_MONTHS_SHORT[d.month]} {year_be % 100:02d}'


def format_date_thai_short(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    elif not isinstance(d, date):
        return str(d)
    year_be = d.year + (543 if d.year < 2400 else 0)
    return f'{d.day} {THAI_MONTHS_SHORT[d.month]} {year_be}'


def format_date_thai(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    elif not isinstance(d, date):
        return str(d)
    year_be = d.year + (543 if d.year < 2400 else 0)
    return f'{d.day:02d}/{d.month:02d}/{year_be}'


def format_datetime_thai(dt):
    if dt is None:
        return ''
    if isinstance(dt, datetime):
        local_dt = timezone.localtime(dt) if timezone.is_aware(dt) else dt
        return f'{format_date_thai(local_dt.date())} {local_dt.strftime("%H:%M น.")}'
    return format_date_thai(dt)


def get_thai_day_name(weekday_int):
    return THAI_DAYS[weekday_int] if 0 <= weekday_int <= 6 else ''


def format_date_thai_long(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    elif not isinstance(d, date):
        return str(d)
    year_be = d.year + (543 if d.year < 2400 else 0)
    return f'{d.day} {THAI_MONTHS[d.month]} {year_be}'


def format_date_thai_full(d):
    if d is None:
        return ''
    if isinstance(d, datetime):
        d = timezone.localtime(d).date() if timezone.is_aware(d) else d.date()
    elif not isinstance(d, date):
        return str(d)
    ce = _to_ce_date_safe(d)
    day_name = get_thai_day_name(ce.weekday())
    return f'วัน{day_name}ที่ {format_date_thai_long(d)} ({format_date_thai(d)})'


@register.filter(name='date_thai')
def date_thai(value):
    return format_date_thai(value)


@register.filter(name='date_thai_short')
def date_thai_short(value):
    return format_date_thai_short(value)


@register.filter(name='date_thai_short_weekday')
def date_thai_short_weekday(value):
    return format_date_thai_short_weekday(value)


@register.filter(name='datetime_thai')
def datetime_thai(value):
    return format_datetime_thai(value)


@register.filter(name='date_thai_long')
def date_thai_long(value):
    return format_date_thai_long(value)


@register.filter(name='date_thai_full')
def date_thai_full(value):
    return format_date_thai_full(value)


@register.filter(name='day_name_thai')
def day_name_thai(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return get_thai_day_name(_to_ce_date_safe(value).weekday())
    return ''


@register.filter(name='day_short_thai')
def day_short_thai(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return get_thai_day_short(_to_ce_date_safe(value).weekday())
    return ''


@register.filter(name='to_thai_num')
def to_thai_num(value):
    """แปลงตัวเลขอารบิกเป็นเลขไทย เช่น 2569 -> ๒๕๖๙"""
    if value is None:
        return ''
    return str(value).translate(str.maketrans('0123456789', '๐๑๒๓๔๕๖๗๘๙'))


@register.filter(name='get_item')
def get_item(dictionary, key):
    if isinstance(dictionary, dict):
        if key in dictionary:
            return dictionary.get(key)
        # ข้อมูลที่ผ่าน Session (JSON) คีย์ตัวเลขจะถูกแปลงเป็นข้อความ จึงลองค้นหาด้วย str(key) อีกครั้ง
        return dictionary.get(str(key))
    return None


@register.simple_tag
def info_tip(*parts, placement='auto'):
    """แสดงไอคอน (i) สำหรับวางเมาส์เพื่ออ่านข้อความอธิบาย"""
    text = ''.join(str(p) for p in parts if p is not None).replace('\\n', '\n').strip()
    if not text:
        return ''
    return format_html(
        '<span class="info-tip-icon" tabindex="0" role="button" aria-label="คำอธิบายเพิ่มเติม" '
        'data-bs-toggle="tooltip" data-bs-placement="{}" data-bs-title="{}" '
        'onclick="event.stopPropagation();"><i class="bi bi-info-circle-fill"></i></span>',
        placement,
        text,
    )
