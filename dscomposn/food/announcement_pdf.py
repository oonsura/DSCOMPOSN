"""
สร้างเอกสารประกาศผลการคัดเลือกเป็นไฟล์ PDF (ตามรูปแบบหนังสือราชการและไฟล์ Word/Excel)
- แผ่นที่ 1: ประกาศโรงเรียนเทพศิรินทร์ (ตราครุฑตรงกลาง, เนื้อหาตามไฟล์ Word, ลายเซ็นผู้อำนวยการ)
- แผ่นที่ 2: บัญชีรายชื่อผู้ผ่านการคัดเลือกแนบท้ายประกาศ (4 ตราสัญลักษณ์, ตารางตามไฟล์ Excel)
- รองรับการปรับแต่งกั้นหน้า-กั้นหลัง (margins), ขนาดตัวอักษร, และความกว้างแต่ละช่องตาราง
- จัดย่อหน้าข้อความภาษาไทยแบบตัดคำชิดซ้ายปกติ (ไม่ยืดช่องไฟระหว่างตัวอักษร) ตรงกับหน้าเว็บและ Word 1:1
"""
import io
import os
import re

from django.conf import settings
from django.utils import timezone
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from PIL import Image

from composn.models import SystemSetting
from composn.templatetags.composn_tags import THAI_MONTHS
from food.signin_pdf import (
    _ensure_thai_fonts, shape_thai, _load_logo_reader, _draw_logo_box, _fit_text
)

_THAI_TO_ARABIC_TRANS = str.maketrans('๐๑๒๓๔๕๖๗๘๙', '0123456789')


def to_arabic_num(text):
    """แปลงตัวเลขไทย (๐-๙) เป็นเลขอารบิก (0-9) ทั้งหมด"""
    return str(text or '').translate(_THAI_TO_ARABIC_TRANS)


# รายการคำสำหรับตัดคำภาษาไทยให้มีความเป็นธรรมชาติแบบหนังสือราชการ
THAI_COMMON_WORDS = [
    'สำนักงาน', 'คณะกรรมการ', 'การศึกษา', 'ขั้นพื้นฐาน',
    'มูลนิธิ', 'ส่งเสริม', 'โอลิมปิกวิชาการ', 'โอลิมปิก', 'วิชาการ',
    'และ', 'พัฒนา', 'มาตรฐาน', 'วิทยาศาสตร์', 'ศึกษา',
    'ในพระอุปถัมภ์', 'พระอุปถัมภ์', 'สมเด็จ', 'พระพี่นางเธอ',
    'เจ้าฟ้า', 'กัลยาณิวัฒนา',
    'กรมหลวง', 'นราธิวาส', 'ราชนครินทร์',
    'ได้ดำเนินการ', 'ดำเนินการ', 'จัดตั้ง', 'ศูนย์อบรม', 'ศูนย์', 'อบรม',
    'สาขาวิชา', 'สาขา', 'วิชา', 'คอมพิวเตอร์',
    'ศูนย์โรงเรียนเทพศิรินทร์-', 'โรงเรียนเทพศิรินทร์', 'โรงเรียน', 'เทพศิรินทร์',
    'สถาบัน', 'เทคโนโลยี', 'พระจอมเกล้า', 'เจ้าคุณทหาร',
    'ลาดกระบัง', 'สำหรับ', 'ในเขต', 'พื้นที่', 'กรุงเทพมหานคร',
    'จัดการอบรม', 'เป็นที่', 'เรียบร้อยแล้วนั้น', 'เรียบร้อยแล้ว',
    'บัดนี้', 'ตัดสินผล', 'การสอบ', 'คัดเลือก', 'นักเรียน',
    'เพื่อ', 'เข้ารับ', 'การอบรม', 'จึงขอ', 'ประกาศ', 'ประกาศผลการคัดเลือก', 'ผลการคัดเลือก',
    'รายชื่อ', 'ที่ผ่าน', 'การคัดเลือก', 'จำนวน', 'โครงการ', 'ปีการศึกษา',
    'มหาวิทยาลัย', 'กระทรวง', 'ศึกษาธิการ',
    'ตามรายชื่อ', 'แนบท้าย', 'ประกาศนี้',
    'ตามที่', 'ร่วมกับ', '(สพฐ.)', '(สอวน.)', 'สอวน.', 'ค่าย', '1', '2', '3', 'คน',
]
THAI_COMMON_WORDS.sort(key=lambda x: len(x), reverse=True)
_THAI_SPLIT_PATTERN = '|'.join(re.escape(t) for t in THAI_COMMON_WORDS)

_THAI_COMBINING_MARKS = set('\u0e31\u0e34\u0e35\u0e36\u0e37\u0e38\u0e39\u0e3a\u0e47\u0e48\u0e49\u0e4a\u0e4b\u0e4c\u0e4d\u0e4e\u0e33')


def tokenize_thai_text(text):
    """
    ตัดข้อความภาษาไทยออกเป็นหน่วยคำตามคำศัพท์ภาษาไทย, กลุ่มตัวเลข/อังกฤษ
    และจับกลุ่มสระบน-ล่าง/วรรณยุกต์ไว้กับพยัญชนะต้นเสมอเพื่อไม่ให้สระ/วรรณยุกต์หลุดขึ้นบรรทัดใหม่
    """
    text = str(text or '')
    tokens = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] == ' ':
            tokens.append(' ')
            i += 1
            continue
        m = re.match(_THAI_SPLIT_PATTERN, text[i:])
        if m:
            tokens.append(m.group(0))
            i += len(m.group(0))
            continue
        m_ascii = re.match(r'[A-Za-z0-9.,/%()\-]+', text[i:])
        if m_ascii:
            tokens.append(m_ascii.group(0))
            i += len(m_ascii.group(0))
            continue
        cluster = text[i]
        i += 1
        while i < n and (text[i] in _THAI_COMBINING_MARKS or (0xF700 <= ord(text[i]) <= 0xF71F)):
            cluster += text[i]
            i += 1
        tokens.append(cluster)
    return tokens


def _split_thai_clusters(word_str):
    """
    แบ่งคำภาษาไทยออกเป็นกลุ่มตัวอักษรฐาน (Base Character + สระบน/ล่าง/วรรณยุกต์/สระอำ)
    โดยไม่ทำให้สระลอยหรือวรรณยุกต์แยกออกจากพยัญชนะต้น
    """
    if not word_str:
        return []
    clusters = []
    for ch in word_str:
        if clusters and (ch in _THAI_COMBINING_MARKS or (0xF700 <= ord(ch) <= 0xF71F)):
            clusters[-1] += ch
        else:
            clusters.append(ch)
    return clusters


def compute_thai_justified_lines(c, text_str, content_w, font_name, font_size, indent=25 * mm):
    """
    คำนวณรายการบรรทัดย่อยของย่อหน้าภาษาไทยแบบตัดคำชิดซ้ายปกติ (ไม่ยืดช่องไฟระหว่างตัวอักษร)
    รองรับการขึ้นย่อหน้าใหม่ด้วยอักขระขึ้นบรรทัดใหม่ และใช้ ReportLab stringWidth จริง
    เพื่อให้ทั้งไฟล์ PDF และการแสดงผลบนหน้าเว็บ (HTML Preview) ตรงกัน 1:1
    """
    raw_paras = str(text_str or '').replace('\r\n', '\n').replace('\r', '\n').split('\n')
    lines = []

    def _make_line_entry(tok_list, l_ind, a_w, is_first_ln, is_last_ln, is_sub_para_start=False):
        line_text = (''.join(tok_list)).strip()
        shaped_line = shape_thai(line_text)
        return {
            'tokens': tok_list,
            'text': line_text,
            'shaped_text': shaped_line,
            'indent': l_ind,
            'indent_mm': round(l_ind / mm, 3),
            'avail_w': a_w,
            'is_first': is_first_ln,
            'is_last': is_last_ln,
            'is_sub_para_start': is_sub_para_start,
            # เก็บ 1 รายการทั้งบรรทัด (ไม่แยกทีละตัวอักษร) เพื่อความเข้ากันได้ย้อนหลัง
            'positioned_words': [{
                'word': line_text,
                'shaped_word': shaped_line,
                'x_pt': l_ind,
                'x_mm': round(l_ind / mm, 3),
            }] if line_text else [],
        }

    sub_para_idx = 0
    for raw_para in raw_paras:
        para_clean = raw_para.strip()
        if not para_clean:
            continue
        tokens = tokenize_thai_text(para_clean)
        cur_line_toks = []
        line_indent = indent
        avail_w = content_w - line_indent
        para_line_idx = 0

        for tok in tokens:
            if not cur_line_toks and tok == ' ':
                continue
            test_line = (''.join(cur_line_toks + [tok])).strip()
            sw = c.stringWidth(shape_thai(test_line), font_name, font_size)
            if sw <= avail_w or not cur_line_toks:
                cur_line_toks.append(tok)
            else:
                lines.append(_make_line_entry(
                    cur_line_toks, line_indent, avail_w, para_line_idx == 0, False,
                    is_sub_para_start=(sub_para_idx > 0 and para_line_idx == 0),
                ))
                para_line_idx += 1
                cur_line_toks = [tok] if tok != ' ' else []
                line_indent = 0.0
                avail_w = content_w
        if cur_line_toks:
            lines.append(_make_line_entry(
                cur_line_toks, line_indent, avail_w, para_line_idx == 0, True,
                is_sub_para_start=(sub_para_idx > 0 and para_line_idx == 0),
            ))
        sub_para_idx += 1
    return lines


compute_thai_wrapped_lines = compute_thai_justified_lines


def get_announcement_html_lines(setting, passed_count):
    """
    คำนวณการตัดบรรทัดของย่อหน้าที่ 1 และย่อหน้าที่ 2 ด้วย ReportLab stringWidth จริง
    เพื่อให้มุมมองเอกสาร A4 บนหน้าเว็บ (HTML) ตัดคำและขึ้นบรรทัดใหม่ตรงกับไฟล์ PDF 100%
    """
    font_reg, _ = _ensure_thai_fonts()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, _ = A4
    margin_left = float(getattr(setting, 'announcement_margin_left_mm', 25.0) or 25.0) * mm
    margin_right = float(getattr(setting, 'announcement_margin_right_mm', 25.0) or 25.0) * mm
    content_w = width - margin_left - margin_right
    font_size_body = float(getattr(setting, 'announcement_font_size_body', 14.5) or 14.5)
    indent_mm = float(getattr(setting, 'announcement_indent_mm', 25.0) or 25.0)
    indent = indent_mm * mm

    p1_text = format_announcement_body_p1(setting)
    p2_text = format_announcement_body_p2(setting, passed_count)
    p1_lines = compute_thai_justified_lines(c, p1_text, content_w, font_reg, font_size_body, indent=indent)
    p2_lines = compute_thai_justified_lines(c, p2_text, content_w, font_reg, font_size_body, indent=indent)
    return p1_lines, p2_lines


def get_announcement_page1_metrics(setting):
    """
    คำนวณพิกัดจัดวางแผ่นที่ 1 ด้วย ReportLab stringWidth จริง:
    - ขอบซ้ายของ 'ประกาศ ณ วันที่...' ตรงกับขอบซ้ายของคำว่า 'ประกาศ...' ในหัวข้อบรรทัดที่ 1 ตามแบบสารบรรณ
    - ตำแหน่งกึ่งกลางบล็อกลายเซ็นผู้ลงนาม จัดกึ่งกลางตรงกับบรรทัด 'ประกาศ ณ วันที่...' โดยอัตโนมัติ (ไม่ล้นขอบซ้าย/ขวา)
    เพื่อให้ HTML, PDF และ Word ตรงกัน 1:1
    """
    font_reg, font_bold = _ensure_thai_fonts()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, _ = A4
    margin_left = float(getattr(setting, 'announcement_margin_left_mm', 25.0) or 25.0) * mm
    margin_right = float(getattr(setting, 'announcement_margin_right_mm', 25.0) or 25.0) * mm
    content_w = max(10.0 * mm, width - margin_left - margin_right)
    center_x = margin_left + (content_w / 2.0)

    fs_title1 = float(getattr(setting, 'announcement_font_size_title1', 17.0) or 17.0)
    font_size_body = float(getattr(setting, 'announcement_font_size_body', 14.5) or 14.5)
    fs_dir_name = float(getattr(setting, 'announcement_font_size_director_name', 14.5) or 14.5)
    fs_dir_pos = float(getattr(setting, 'announcement_font_size_director_pos', 13.0) or 13.0)

    # 1. ขอบซ้ายของ "ประกาศ ณ วันที่..." ตรงกับขอบซ้ายของ "ประกาศ..." ในหัวข้อบรรทัดแรก
    title1_text = to_arabic_num(f'ประกาศ{setting.announcement_school_name}')
    title1_w = c.stringWidth(shape_thai(title1_text), font_bold, fs_title1)
    date_text = format_announcement_date_text(setting)
    date_w = c.stringWidth(shape_thai(date_text), font_reg, font_size_body)

    date_offset_from_margin = max(0.0, (content_w - title1_w) / 2.0)
    if (date_offset_from_margin + date_w) > content_w:
        date_offset_from_margin = max(0.0, content_w - date_w)
    date_x = margin_left + date_offset_from_margin
    date_center_x = date_x + (date_w / 2.0)

    # 2. ตำแหน่งกึ่งกลางบล็อกลายเซ็นผู้ลงนาม จัดกึ่งกลางตรงกับบรรทัด "ประกาศ ณ วันที่..." โดยอัตโนมัติ (ไม่ล้นขอบซ้าย/ขวา)
    dir_name_text = to_arabic_num(f'({setting.announcement_director_name})')
    pos_lines = [to_arabic_num(l.strip()) for l in (setting.announcement_director_position or '').splitlines() if l.strip()]
    max_sig_w = c.stringWidth(shape_thai(dir_name_text), font_bold, fs_dir_name)
    for pl in pos_lines:
        max_sig_w = max(max_sig_w, c.stringWidth(shape_thai(pl), font_reg, fs_dir_pos))

    sig_box_w = max(85.0 * mm, min(content_w, max_sig_w))
    sig_x = date_center_x
    max_allowed_sig_x = (width - margin_right) - (max_sig_w / 2.0)
    min_allowed_sig_x = margin_left + (max_sig_w / 2.0)
    if max_allowed_sig_x >= min_allowed_sig_x:
        sig_x = min(sig_x, max_allowed_sig_x)
        sig_x = max(sig_x, min_allowed_sig_x)
    else:
        sig_x = center_x

    sig_left_from_margin = (sig_x - margin_left) - (sig_box_w / 2.0)

    return {
        'content_w': content_w,
        'center_x': center_x,
        'title1_w': title1_w,
        'date_x': date_x,
        'date_w': date_w,
        'date_center_x': date_center_x,
        'date_indent_mm': round(date_offset_from_margin / mm, 3),
        'sig_x': sig_x,
        'sig_box_w': sig_box_w,
        'sig_box_w_mm': round(sig_box_w / mm, 3),
        'sig_left_from_margin_mm': round(sig_left_from_margin / mm, 3),
    }


def draw_thai_justified_paragraph(c, text_str, start_y, margin_left, content_w, font_name, font_size, line_h, indent=25 * mm, para_gap=0.0):
    """วาดข้อความย่อหน้าภาษาไทยแบบตัดคำชิดซ้ายปกติ (ไม่ยืดช่องไฟระหว่างตัวอักษร) ตรงกับหน้าเว็บ 1:1"""
    lines = compute_thai_justified_lines(c, text_str, content_w, font_name, font_size, indent=indent)
    c.setFont(font_name, font_size)
    c.setFillColor(colors.black)
    cur_y = start_y
    for item in lines:
        if item.get('is_sub_para_start') and para_gap:
            cur_y -= para_gap
        if item['text']:
            c.drawString(margin_left + item['indent'], cur_y, item.get('shaped_text') or shape_thai(item['text']))
        cur_y -= line_h
    return cur_y


draw_thai_paragraph = draw_thai_justified_paragraph


def _fit_text_shrink(c, text, font_name, initial_size, max_w, min_size=6.0, pad=0.0):
    """
    ปรับขนาดตัวอักษรลดลงอัตโนมัติ (Dynamic Font Shrinking) เพื่อให้ข้อความอยู่บรรทัดเดียว
    ไม่ตัดคำและไม่ขึ้นบรรทัดใหม่ โดยเฉพาะชื่อโรงเรียนและชื่อนักเรียน
    """
    text = str(text or '')
    avail_w = max_w - pad
    if avail_w <= 0 or not text:
        return shape_thai(text), initial_size

    cur_size = float(initial_size)
    st = shape_thai(text)
    while cur_size >= min_size:
        sw = c.stringWidth(st, font_name, cur_size)
        if sw <= avail_w:
            return st, cur_size
        cur_size -= 0.25

    # หากลดขนาดจนถึง min_size แล้วยังเกินเล็กน้อย ให้ใช้ _fit_text ตัดท้าย
    return _fit_text(c, text, font_name, min_size, avail_w), min_size


def format_default_announcement_date():
    d_now = timezone.localdate()
    y_th = d_now.year + (543 if d_now.year < 2400 else 0)
    return f'ประกาศ ณ วันที่ {d_now.day} {THAI_MONTHS[d_now.month]} พ.ศ. {y_th}'


def format_announcement_date_text(setting):
    date_text = to_arabic_num((getattr(setting, 'announcement_date_text', '') or '').strip())
    if not date_text:
        return format_default_announcement_date()
    if not date_text.startswith('ประกาศ ณ'):
        return f'ประกาศ ณ {date_text}'
    return date_text


def format_announcement_list_title(setting, year):
    custom = to_arabic_num((getattr(setting, 'announcement_list_title', '') or '').strip())
    if custom:
        return custom
    return f'ประกาศรายชื่อผู้ที่ผ่านการคัดเลือกเข้า สอวน.คอมพิวเตอร์ ค่ายที่ 2 ปีการศึกษา {year}'


def format_default_announcement_p1(setting):
    center_name = to_arabic_num(getattr(setting, 'announcement_center_name', '') or 'ศูนย์โรงเรียนเทพศิรินทร์-สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหาร ลาดกระบัง')
    return to_arabic_num(
        f'ตามที่ สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน (สพฐ.) ร่วมกับ มูลนิธิส่งเสริมโอลิมปิกวิชาการ'
        f'และพัฒนามาตรฐานวิทยาศาสตร์ศึกษา ในพระอุปถัมภ์สมเด็จพระพี่นางเธอ เจ้าฟ้ากัลยาณิวัฒนา กรมหลวงนราธิวาสราชนครินทร์ (สอวน.) '
        f'ได้ดำเนินการจัดตั้งศูนย์อบรม สอวน. สาขาวิชาคอมพิวเตอร์ ค่าย 1 {center_name} '
        f'สำหรับโรงเรียนในเขตพื้นที่กรุงเทพมหานคร และได้ดำเนินการจัดการอบรม ค่าย 1 เป็นที่เรียบร้อยแล้วนั้น'
    )


def format_announcement_body_p1(setting):
    custom = (getattr(setting, 'announcement_body_p1', '') or '').strip()
    if custom:
        return to_arabic_num(custom)
    return format_default_announcement_p1(setting)


def format_default_announcement_p2_template():
    return (
        'บัดนี้ คณะกรรมการได้ดำเนินการตัดสินผลการสอบคัดเลือกนักเรียนเพื่อเข้ารับการอบรม สอวน. '
        'สาขาวิชาคอมพิวเตอร์ ค่าย 2 เรียบร้อยแล้ว จึงขอประกาศรายชื่อนักเรียนที่ผ่านการคัดเลือกเข้ารับการอบรม สอวน. '
        'วิชาคอมพิวเตอร์ ค่าย 2 จำนวน {count} คน ตามรายชื่อแนบท้ายประกาศนี้'
    )


def format_default_announcement_p2(passed_count):
    return format_default_announcement_p2_template().replace('{count}', str(passed_count))


def format_announcement_body_p2(setting, passed_count):
    custom = (getattr(setting, 'announcement_body_p2', '') or '').strip()
    passed_str = str(passed_count)
    if custom:
        return to_arabic_num(custom.replace('{count}', passed_str))
    return format_default_announcement_p2(passed_count)


def _load_garuda_reader(setting):
    """โหลดตราครุฑแบบพื้นหลังโปร่งใสจาก ImageField หรือจาก static/img/garuda.png (.jpg fallback)"""
    if setting.logo_garuda:
        r = _load_logo_reader(setting.logo_garuda)
        if r:
            return r
    for rel_name in ('garuda.png', 'garuda.jpg'):
        static_garuda = os.path.join(settings.BASE_DIR, 'static', 'img', rel_name)
        if os.path.exists(static_garuda):
            try:
                from composn.logo_utils import remove_logo_background
                with Image.open(static_garuda) as raw_img:
                    raw_img.load()
                    rgba = remove_logo_background(raw_img)
                buf = io.BytesIO()
                rgba.save(buf, format='PNG')
                buf.seek(0)
                return ImageReader(buf)
            except Exception:
                pass
    return None


def draw_announcement_page(c, year, setting, passed_count, width, height, font_reg, font_bold):
    """วาดหน้าประกาศผลทางการ (แผ่นที่ 1: รูปแบบหนังสือราชการ ตราครุฑ ตาม Word/Web)"""
    margin_left = float(getattr(setting, 'announcement_margin_left_mm', 25.0) or 25.0) * mm
    margin_right = float(getattr(setting, 'announcement_margin_right_mm', 25.0) or 25.0) * mm
    margin_top = float(getattr(setting, 'announcement_margin_top_mm', 25.0) or 25.0) * mm
    content_w = width - margin_left - margin_right
    center_x = margin_left + (content_w / 2.0)

    fs_title1 = float(getattr(setting, 'announcement_font_size_title1', 17.0) or 17.0)
    fs_title2 = float(getattr(setting, 'announcement_font_size_title2', 15.0) or 15.0)
    fs_title3 = float(getattr(setting, 'announcement_font_size_title3', 14.0) or 14.0)
    font_size_body = float(getattr(setting, 'announcement_font_size_body', 14.5) or 14.5)
    fs_dir_name = float(getattr(setting, 'announcement_font_size_director_name', 14.5) or 14.5)
    fs_dir_pos = float(getattr(setting, 'announcement_font_size_director_pos', 13.0) or 13.0)

    # ระยะห่างระหว่างบรรทัดเดียวเท่ากันทั้งหน้าแผ่นที่ 1 (หัวข้อประกาศ, เนื้อหาย่อหน้า 1-2, วันที่, และตำแหน่งผู้อำนวยการ)
    line_spacing_mult = float(getattr(setting, 'announcement_line_spacing', 1.85) or 1.85)
    line_h = (font_size_body * 0.44 * line_spacing_mult) * mm
    _raw_para_sp = getattr(setting, 'announcement_paragraph_spacing_mm', None)
    para_spacing_mm = 4.0 if _raw_para_sp is None else max(0.0, float(_raw_para_sp))
    para_gap = para_spacing_mm * mm
    indent_mm = float(getattr(setting, 'announcement_indent_mm', 25.0) or 25.0)
    indent = indent_mm * mm
    garuda_h_mm = float(getattr(setting, 'announcement_garuda_height_mm', 30.0) or 30.0)

    # 1. ตราครุฑตรงกลางด้านบน (สูงมาตรฐานหนังสือราชการ 3.0 ซม. หรือตามที่ปรับแต่ง)
    garuda_rdr = _load_garuda_reader(setting)
    garuda_h = garuda_h_mm * mm
    garuda_w = (garuda_h_mm * (28.0 / 30.0)) * mm
    garuda_y = height - margin_top - garuda_h
    if garuda_rdr:
        _draw_logo_box(c, garuda_rdr, center_x - (garuda_w / 2.0), garuda_y, garuda_w, garuda_h)

    # 2. หัวข้อประกาศ (ใช้ระยะห่างบรรทัด line_h เดียวกันทั้งหมดทั้งหน้า และคำนวณ baseline ตรงกับ CSS line-box 1:1)
    metrics = get_announcement_page1_metrics(setting)
    title1_baseline_off = (0.5 * line_h) + (0.1935 * (fs_title1 * 25.4 / 72.0) * mm)
    title_y = garuda_y - (5.0 * mm) - title1_baseline_off
    c.setFont(font_bold, fs_title1)
    c.setFillColor(colors.black)
    c.drawCentredString(center_x, title_y, shape_thai(to_arabic_num(f'ประกาศ{setting.announcement_school_name}')))

    c.setFont(font_bold, fs_title2)
    c.drawCentredString(center_x, title_y - line_h, shape_thai(to_arabic_num(f'{setting.announcement_title} ค่าย 2 ปีการศึกษา {year}')))

    c.setFont(font_bold, fs_title3)
    c.drawCentredString(center_x, title_y - (2.0 * line_h), shape_thai(to_arabic_num(setting.announcement_center_name)))

    # เส้นคั่นแนวนอนตรงกลางยาว 50 มม. (ตามแบบหนังสือราชการ/สารบรรณ)
    divider_y = title_y - (3.0 * line_h)
    div_mid_y = divider_y + (0.1935 * (fs_title1 * 25.4 / 72.0) * mm)
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.75)
    c.line(center_x - (25.0 * mm), div_mid_y, center_x + (25.0 * mm), div_mid_y)

    # 3. ย่อหน้าที่ 1: ตามที่... (ชิดซ้ายปกติ ไม่ยืดช่องไฟ ระยะห่างทุกบรรทัด = line_h และใช้ฟอนต์ font_size_body จริง)
    p1_y = divider_y - line_h
    text_p1 = format_announcement_body_p1(setting)
    cur_y = draw_thai_justified_paragraph(
        c, text_p1, p1_y, margin_left, content_w, font_reg, font_size_body, line_h, indent=indent, para_gap=para_gap
    )

    # 4. ย่อหน้าที่ 2: บัดนี้... (เว้นระยะห่างระหว่างย่อหน้า para_gap)
    cur_y -= para_gap
    text_p2 = format_announcement_body_p2(setting, passed_count)
    cur_y = draw_thai_justified_paragraph(
        c, text_p2, cur_y, margin_left, content_w, font_reg, font_size_body, line_h, indent=indent, para_gap=para_gap
    )

    # 5. วันที่ในประกาศ (เว้น 1 1/2 บรรทัดจากบรรทัดสุดท้ายของเนื้อหา และจัดขอบซ้ายตรงกับคำว่า "ประกาศ..." ด้านบน)
    cur_y -= (0.5 * line_h)
    date_text = format_announcement_date_text(setting)
    c.setFont(font_reg, font_size_body)
    c.setFillColor(colors.black)
    c.drawString(metrics['date_x'], cur_y, shape_thai(date_text))

    # 6. ลายเซ็นผู้อำนวยการ (เว้น 3 บรรทัด = 3.0 * line_h จากบรรทัดวันที่) และตำแหน่งผู้อำนวยการ (ใช้ระยะ line_h เดียวกันทุกบรรทัด)
    sig_x = metrics['sig_x']
    sig_y = cur_y - (3.0 * line_h)
    c.setFont(font_bold, fs_dir_name)
    c.drawCentredString(sig_x, sig_y, shape_thai(to_arabic_num(f'({setting.announcement_director_name})')))

    pos_lines = [to_arabic_num(l.strip()) for l in (setting.announcement_director_position or '').splitlines() if l.strip()]
    c.setFont(font_reg, fs_dir_pos)
    for idx, pl in enumerate(pos_lines, 1):
        c.drawCentredString(sig_x, sig_y - (idx * line_h), shape_thai(pl))


def draw_attached_students_page(c, rows, year, setting, width, height, font_reg, font_bold):
    """วาดหน้ารายชื่อแนบท้ายประกาศ (แผ่นที่ 2: บังคับจัดให้อยู่ในหน้าเดียว เพื่อให้เอกสารประกาศผลมี 2 หน้าพอดี)"""
    margin_left = float(getattr(setting, 'announcement_margin_left_mm', 25.0) or 25.0) * mm
    margin_right = float(getattr(setting, 'announcement_margin_right_mm', 25.0) or 25.0) * mm
    margin_top = float(getattr(setting, 'announcement_margin_top_mm', 25.0) or 25.0) * mm
    margin_bottom = float(getattr(setting, 'announcement_margin_bottom_mm', 25.0) or 25.0) * mm
    content_w = width - margin_left - margin_right
    center_x = margin_left + (content_w / 2.0)
    fs_list1 = float(getattr(setting, 'announcement_font_size_list_title1', 14.0) or 14.0)
    fs_list2 = float(getattr(setting, 'announcement_font_size_list_title2', 12.0) or 12.0)

    # 1. ตราสัญลักษณ์ 4 รูปด้านบน จัดกึ่งกลางพอเหมาะเหมือนหน้าเว็บ
    logo_w, logo_h = 16 * mm, 16 * mm
    logo_y = height - margin_top - logo_h
    readers = []
    for field in (setting.logo_school, setting.logo_obec, setting.logo_posn, setting.logo_center):
        r = _load_logo_reader(field)
        if r:
            readers.append(r)

    if readers:
        spacing = 24 * mm
        start_x = center_x - ((len(readers) - 1) * spacing) / 2.0 - (logo_w / 2.0)
        for idx, rdr in enumerate(readers):
            _draw_logo_box(c, rdr, start_x + idx * spacing, logo_y, logo_w, logo_h)

    # 2. หัวข้อรายชื่อ
    head_y = logo_y - 7 * mm
    c.setFont(font_bold, fs_list1)
    c.setFillColor(colors.black)
    c.drawCentredString(center_x, head_y, shape_thai(format_announcement_list_title(setting, year)))
    c.setFont(font_reg, fs_list2)
    c.drawCentredString(center_x, head_y - 5.5 * mm, shape_thai(to_arabic_num(setting.announcement_center_name)))

    # 3. คำนวณขนาดคอลัมน์ตารางให้เต็มความกว้าง content_w พอดีตามสัดส่วนที่ปรับ
    raw_col_w = [
        float(getattr(setting, 'announcement_col_w_no', 14.0) or 14.0),
        float(getattr(setting, 'announcement_col_w_camp2', 24.0) or 24.0),
        float(getattr(setting, 'announcement_col_w_rights', 24.0) or 24.0),
        float(getattr(setting, 'announcement_col_w_name', 50.0) or 50.0),
        float(getattr(setting, 'announcement_col_w_school', 54.0) or 54.0),
        float(getattr(setting, 'announcement_col_w_grade', 14.0) or 14.0),
    ]
    raw_total = sum(raw_col_w) or 176.0
    col_w = [(w / raw_total) * content_w for w in raw_col_w]
    table_w = content_w
    start_tx = margin_left
    table_top_y = head_y - 13 * mm

    font_size_table = float(getattr(setting, 'announcement_font_size_table', 10.0) or 10.0)
    row_h = max(5.8 * mm, (font_size_table * 0.58) * mm)
    thead_h = max(7.0 * mm, (font_size_table * 0.70) * mm)
    pad_x = 1.5 * mm

    # คำนวณปรับขนาดความสูงแถวอัตโนมัติเพื่อให้รายชื่อทั้งหมดและหมายเหตุอยู่ในแผ่นที่ 2 (เอกสารรวม 2 หน้าพอดี)
    note_lines = [to_arabic_num(l.strip()) for l in (getattr(setting, 'announcement_note', '') or '').splitlines() if l.strip()]
    note_reserve_h = ((6.5 + len(note_lines) * 4.5) * mm) if note_lines else 0.0
    safe_bottom_y = min(margin_bottom, 16 * mm)
    avail_table_h = table_top_y - safe_bottom_y - note_reserve_h
    n_rows = max(len(rows), 1)
    if (thead_h + n_rows * row_h) > avail_table_h and avail_table_h > 40 * mm:
        scale = avail_table_h / float(thead_h + n_rows * row_h)
        row_h = max(3.6 * mm, row_h * scale)
        thead_h = max(5.0 * mm, thead_h * scale)
        font_size_table = max(7.0, min(font_size_table, (row_h / mm) / 0.54))

    # Header ตาราง (หัวข้อบรรทัดเดียว สไตล์ Word Table Grid ตามคำขอ)
    c.setFillColor(colors.HexColor('#f0f4f8'))
    c.rect(start_tx, table_top_y - thead_h, table_w, thead_h, fill=1, stroke=1)
    c.setFillColor(colors.black)

    headers = ['ลำดับ', 'เลขประจำตัวค่าย 2', 'เลขประจำตัวค่าย 1', 'ชื่อ นามสกุล', 'โรงเรียน', 'ระดับ']
    cur_x = start_tx
    for idx, (h, w) in enumerate(zip(headers, col_w)):
        mid_cx = cur_x + (w / 2.0)
        shaped_h, fitted_h_size = _fit_text_shrink(c, h, font_bold, font_size_table, w, min_size=6.8, pad=2 * pad_x)
        c.setFont(font_bold, fitted_h_size)
        y_head = table_top_y - thead_h + (thead_h - fitted_h_size * 0.35 * mm) / 2.0 + 0.3 * mm
        c.drawCentredString(mid_cx, y_head, shaped_h)
        cur_x += w

    # เส้นแบ่งคอลัมน์ Header
    c.setStrokeColor(colors.HexColor('#333333'))
    c.setLineWidth(0.5)
    cur_x = start_tx
    for w in col_w:
        c.line(cur_x, table_top_y, cur_x, table_top_y - thead_h)
        cur_x += w
    c.line(start_tx + table_w, table_top_y, start_tx + table_w, table_top_y - thead_h)

    # แถวข้อมูลนักเรียน (ข้อความทุกช่องอยู่บรรทัดเดียว ปรับขนาดย่อลงอัตโนมัติ โดยเฉพาะโรงเรียน)
    cur_y = table_top_y - thead_h
    c.setFont(font_reg, font_size_table)

    for idx, r in enumerate(rows, 1):
        # สลับสีพื้นหลังแถวคู่
        if idx % 2 == 0:
            c.setFillColor(colors.HexColor('#fafafa'))
            c.rect(start_tx, cur_y - row_h, table_w, row_h, fill=1, stroke=0)
        c.setFillColor(colors.black)

        # ข้อมูลแต่ละคอลัมน์
        km_code = r.get('camp2_code') or r['user'].camp2_code or (f'{str(year)[-2:]}KM{idx:02d}' if r.get('passed', True) else '-')
        orig_code = r.get('rights_code') or r['user'].camp1_code
        name = r['user'].get_full_name()
        school = r['user'].school_name or ''
        grade = r.get('grade') or r['user'].current_level or ''

        col_vals = [
            (str(idx), 'center', col_w[0], False),
            (km_code, 'center', col_w[1], True),
            (orig_code, 'center', col_w[2], False),
            (name, 'left', col_w[3], True),
            (school, 'left', col_w[4], False),
            (grade, 'center', col_w[5], False),
        ]

        cx = start_tx
        for val, align, cw, is_b in col_vals:
            fn = font_bold if is_b else font_reg
            # ลดขนาดตัวอักษรลงอัตโนมัติเพื่อให้ข้อความทุกช่องอยู่ในบรรทัดเดียวเสมอ
            shaped_val, fitted_size = _fit_text_shrink(c, val, fn, font_size_table, cw, min_size=5.5, pad=2 * pad_x)
            c.setFont(fn, fitted_size)
            y_text = cur_y - row_h + (row_h - fitted_size * 0.35 * mm) / 2.0 + 0.3 * mm
            if align == 'center':
                c.drawCentredString(cx + cw / 2.0, y_text, shaped_val)
            else:
                c.drawString(cx + pad_x, y_text, shaped_val)
            cx += cw

        # เส้นตารางแนวนอนและแนวตั้ง
        c.setStrokeColor(colors.HexColor('#333333'))
        c.setLineWidth(0.5)
        c.rect(start_tx, cur_y - row_h, table_w, row_h, fill=0, stroke=1)
        cur_x = start_tx
        for w in col_w:
            c.line(cur_x, cur_y, cur_x, cur_y - row_h)
            cur_x += w
        c.line(start_tx + table_w, cur_y, start_tx + table_w, cur_y - row_h)

        cur_y -= row_h

    # 4. วาดหมายเหตุท้ายตาราง (ถ้ามี) เหมือนหน้าเว็บและ Word
    if note_lines:
        cur_y -= 4.0 * mm
        c.setStrokeColor(colors.HexColor('#444444'))
        c.setLineWidth(0.5)
        c.line(start_tx, cur_y + 1.5 * mm, start_tx + table_w, cur_y + 1.5 * mm)
        cur_y -= 3.5 * mm
        c.setFont(font_bold, 10.5)
        note_prefix = shape_thai('หมายเหตุ: ')
        c.drawString(start_tx, cur_y, note_prefix)
        prefix_w = c.stringWidth(note_prefix, font_bold, 10.5)
        c.setFont(font_reg, 10.0)
        for idx_l, line in enumerate(note_lines):
            x_pos = start_tx + (prefix_w if idx_l == 0 else 0)
            c.drawString(x_pos, cur_y, shape_thai(line))
            cur_y -= 4.5 * mm


def build_announcement_pdf(rows, year, setting):
    """สร้าง PDF ประกาศผลการคัดเลือก (แผ่นที่ 1: ประกาศทางการ + แผ่นที่ 2: รายชื่อแนบท้าย)"""
    font_reg, font_bold = _ensure_thai_fonts()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f'ประกาศผลการคัดเลือกเข้าค่าย 2 ปีการศึกษา {year}')

    # แผ่นที่ 1: ประกาศราชการ (ตราครุฑ)
    draw_announcement_page(c, year, setting, len(rows), A4[0], A4[1], font_reg, font_bold)
    c.showPage()

    # แผ่นที่ 2: รายชื่อแนบท้ายประกาศ (4 โลโก้ + ตารางตามไฟล์ Excel/Word)
    draw_attached_students_page(c, rows, year, setting, A4[0], A4[1], font_reg, font_bold)
    c.showPage()

    c.save()
    buf.seek(0)
    return buf.getvalue()
