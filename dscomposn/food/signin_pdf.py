"""
สร้างไฟล์ PDF แบบลงชื่อเข้าอบรม / วันสอบท้ายค่าย (A4 แนวนอน · 50–52 คน/แผ่น)
- รองรับฟอนต์ไทย THSarabunNew พร้อมจัดตำแหน่งวรรณยุกต์/สระไทย (PUA shaping) ให้สวยงามคมชัด
- ตราสัญลักษณ์ซ้ายมือ = ตราโรงเรียนเทพศิรินทร์, ขวามือ = ตรา สอวน.
- รองรับหลายแผ่นในไฟล์เดียว (ช่วงวันที่, ภาคเช้า + ภาคบ่าย, แยกแผ่นนักเรียนปัจจุบัน/นักเรียนใช้สิทธิ์ในวันสอบ)
"""
import io
import os

from django.conf import settings
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

_FONT_REG = None
_FONT_BOLD = None
_USE_THAI_PUA = False

# ชุดอักษรไทยสำหรับจัดตำแหน่งสระ/วรรณยุกต์ (มาตรฐาน Thai Windows/SIPA PUA ใน THSarabunNew)
_UP_TAIL_CONSONANTS = set('ปฝฟฬ')
_LOW_TAIL_CONSONANTS = set('ฎฏ')
_BASE_REMOVE_MAP = {'ฐ': '\uf700', 'ญ': '\uf70f'}
_UPPER_VOWELS = set('\u0e31\u0e34\u0e35\u0e36\u0e37\u0e47\u0e4d')
_LOWER_VOWELS = set('\u0e38\u0e39\u0e3a')
_TONES = {'\u0e48': 0, '\u0e49': 1, '\u0e4a': 2, '\u0e4b': 3, '\u0e4c': 4}
_UPPER_VOWEL_LEFT = {
    '\u0e31': '\uf710',
    '\u0e34': '\uf701',
    '\u0e35': '\uf702',
    '\u0e36': '\uf703',
    '\u0e37': '\uf704',
    '\u0e47': '\uf712',
}
_LOWER_VOWEL_LOW = {
    '\u0e38': '\uf718',
    '\u0e39': '\uf719',
    '\u0e3a': '\uf71a',
}


def _ensure_thai_fonts():
    """ลงทะเบียนฟอนต์ไทย (THSarabunNew เป็นหลัก พร้อม fallback อัตโนมัติ)"""
    global _FONT_REG, _FONT_BOLD, _USE_THAI_PUA
    if _FONT_REG and _FONT_BOLD:
        return _FONT_REG, _FONT_BOLD

    base_dir = str(getattr(settings, 'BASE_DIR', ''))
    local_app = os.environ.get('LOCALAPPDATA', '')
    candidates = [
        (
            'THSarabunNew',
            os.path.join(base_dir, 'static', 'fonts', 'THSarabunNew.ttf'),
            os.path.join(base_dir, 'static', 'fonts', 'THSarabunNew-Bold.ttf'),
            True,
        ),
        (
            'THSarabunNewLocal',
            os.path.join(local_app, 'Microsoft', 'Windows', 'Fonts', 'THSarabunNew.ttf'),
            os.path.join(local_app, 'Microsoft', 'Windows', 'Fonts', 'THSarabunNew Bold.ttf'),
            True,
        ),
        (
            'TahomaThai',
            r'C:\Windows\Fonts\tahoma.ttf',
            r'C:\Windows\Fonts\tahomabd.ttf',
            False,
        ),
        (
            'LeelawadeeThai',
            r'C:\Windows\Fonts\leelawad.ttf',
            r'C:\Windows\Fonts\leelawdb.ttf',
            False,
        ),
        (
            'GarudaThai',
            '/usr/share/fonts/truetype/tlwg/Garuda.ttf',
            '/usr/share/fonts/truetype/tlwg/Garuda-Bold.ttf',
            False,
        ),
        (
            'LomaThai',
            '/usr/share/fonts/truetype/tlwg/Loma.ttf',
            '/usr/share/fonts/truetype/tlwg/Loma-Bold.ttf',
            False,
        ),
    ]

    for fam, reg_path, bold_path, pua_hint in candidates:
        if reg_path and os.path.exists(reg_path):
            reg_name = fam
            bold_name = f'{fam}-Bold'
            if reg_name not in pdfmetrics.getRegisteredFontNames():
                tt_reg = TTFont(reg_name, reg_path)
                pdfmetrics.registerFont(tt_reg)
                if pua_hint:
                    cmap = tt_reg.face.charToGlyph
                    _USE_THAI_PUA = all(cp in cmap for cp in range(0xF700, 0xF71B))
            if bold_path and os.path.exists(bold_path):
                if bold_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(bold_name, bold_path))
            else:
                bold_name = reg_name
            _FONT_REG, _FONT_BOLD = reg_name, bold_name
            return _FONT_REG, _FONT_BOLD

    _FONT_REG, _FONT_BOLD = 'Helvetica', 'Helvetica-Bold'
    _USE_THAI_PUA = False
    return _FONT_REG, _FONT_BOLD


def shape_thai(text):
    """ปรับรหัสสระ/วรรณยุกต์ไทยให้ตรงกับ PUA ของ THSarabunNew เพื่อไม่ให้วรรณยุกต์ลอยหรือซ้อนทับ"""
    if not text:
        return ''
    s = str(text)
    if not _USE_THAI_PUA:
        return s

    chars = list(s)
    out = []
    n = len(chars)
    i = 0
    while i < n:
        ch = chars[i]
        # ตรวจสอบพยัญชนะฐานที่มีสระล่างตามมา (ฐ, ญ -> ตัดเชิงล่าง)
        if ch in _BASE_REMOVE_MAP and i + 1 < n and chars[i + 1] in _LOWER_VOWELS:
            out.append(_BASE_REMOVE_MAP[ch])
            i += 1
            continue

        # หาพยัญชนะฐานก่อนหน้า (ข้ามสระบน/ล่าง)
        prev_cons = ''
        has_upper = False
        j = len(out) - 1
        while j >= 0:
            c = out[j]
            if c in _UPPER_VOWELS or c in _UPPER_VOWEL_LEFT.values():
                has_upper = True
                j -= 1
                continue
            if c in _LOWER_VOWELS or c in _LOWER_VOWEL_LOW.values():
                j -= 1
                continue
            prev_cons = c
            break

        if ch in _LOWER_VOWEL_LOW and prev_cons in _LOW_TAIL_CONSONANTS:
            out.append(_LOWER_VOWEL_LOW[ch])
        elif ch in _UPPER_VOWEL_LEFT and prev_cons in _UP_TAIL_CONSONANTS:
            out.append(_UPPER_VOWEL_LEFT[ch])
        elif ch in _TONES:
            idx = _TONES[ch]
            next_is_sara_am = (i + 1 < n and chars[i + 1] == '\u0e33')
            if has_upper or next_is_sara_am:
                if prev_cons in _UP_TAIL_CONSONANTS:
                    out.append(chr(0xF705 + idx))  # สูง + เลื่อนซ้าย
                else:
                    out.append(ch)                 # สูงปกติ (เหนือสระบน/สระอำ)
            else:
                if prev_cons in _UP_TAIL_CONSONANTS:
                    out.append(chr(0xF713 + idx))  # ต่ำ + เลื่อนซ้าย (หลบหาง ป ฝ ฟ ฬ)
                else:
                    out.append(chr(0xF70A + idx))  # ต่ำปกติ (แนบพยัญชนะ ไม่มีสระบน)
        else:
            out.append(ch)
        i += 1
    return ''.join(out)


def _load_logo_reader(image_field):
    """โหลด ImageField เป็น ReportLab ImageReader พร้อมรักษาพื้นหลังโปร่งใส (RGBA PNG)"""
    if not image_field:
        return None
    try:
        path = image_field.path
    except Exception:
        return None
    if not path or not os.path.exists(path):
        return None
    try:
        from composn.logo_utils import remove_logo_background
        with Image.open(path) as raw_img:
            raw_img.load()
            rgba = remove_logo_background(raw_img)
        buf = io.BytesIO()
        rgba.save(buf, format='PNG')
        buf.seek(0)
        return ImageReader(buf)
    except Exception:
        return None


def _fit_text(c, text, font_name, font_size, max_width):
    """ย่อหรือตัดข้อความให้พอดีความกว้างคอลัมน์ไม่ให้ล้นเส้นตาราง"""
    shaped = shape_thai(text)
    if not shaped:
        return ''
    if c.stringWidth(shaped, font_name, font_size) <= max_width:
        return shaped
    raw = str(text)
    while len(raw) > 1:
        raw = raw[:-1]
        cand = shape_thai(raw.rstrip() + '…')
        if c.stringWidth(cand, font_name, font_size) <= max_width:
            return cand
    return shape_thai(raw)


def _draw_logo_box(c, reader, x, y, box_w, box_h):
    if not reader:
        return
    try:
        iw, ih = reader.getSize()
        if iw <= 0 or ih <= 0:
            return
        ratio = min(box_w / float(iw), box_h / float(ih))
        dw, dh = iw * ratio, ih * ratio
        dx = x + (box_w - dw) / 2.0
        dy = y + (box_h - dh) / 2.0
        c.drawImage(reader, dx, dy, width=dw, height=dh, preserveAspectRatio=True, mask='auto')
    except Exception:
        pass


def build_signin_pdf(sheets, setting, camp=1):
    """
    สร้าง PDF ใบลงชื่อเข้าอบรม / วันสอบท้ายค่าย (A4 แนวนอน 1 แผ่นต่อ 1 รายการใน sheets)
    คืนค่า bytes ของไฟล์ PDF
    """
    font_reg, font_bold = _ensure_thai_fonts()
    page_w, page_h = landscape(A4)  # 297mm x 210mm
    margin_x = 6.0 * mm
    margin_top = 5.5 * mm
    margin_bottom = 5.5 * mm
    usable_w = page_w - 2 * margin_x  # 285mm

    logo_left = _load_logo_reader(getattr(setting, 'logo_school', None))
    logo_right = _load_logo_reader(getattr(setting, 'logo_posn', None)) or _load_logo_reader(
        getattr(setting, 'logo_center', None)
    )

    # ความกว้างคอลัมน์ 6 ช่องฝั่งซ้าย + 6 ช่องฝั่งขวา (รวม 142.5mm x 2 = 285mm)
    half_cols = [9.0 * mm, 18.5 * mm, 44.5 * mm, 11.5 * mm, 37.5 * mm, 21.5 * mm]
    col_widths = half_cols + half_cols
    col_xs = [margin_x]
    for w in col_widths:
        col_xs.append(col_xs[-1] + w)
    split_x = col_xs[6]  # เส้นแบ่งกึ่งกลางฝั่งซ้าย-ขวา

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setTitle('แบบลงชื่อเข้าอบรม / วันสอบท้ายค่าย')

    total_sheets = len(sheets) if sheets else 0
    for s_idx, sheet in enumerate(sheets, 1):
        if s_idx > 1:
            c.showPage()

        # ---------- ส่วนหัวเอกสาร (Header) ----------
        top_y = page_h - margin_top
        header_h = 24.5 * mm
        logo_box = 17.5 * mm
        logo_y = top_y - header_h + (header_h - logo_box) / 2.0 + 1.5 * mm
        _draw_logo_box(c, logo_left, margin_x + 1.0 * mm, logo_y, logo_box, logo_box)
        _draw_logo_box(c, logo_right, page_w - margin_x - logo_box - 1.0 * mm, logo_y, logo_box, logo_box)

        cx = page_w / 2.0
        c.setFillColor(colors.black)
        l1 = shape_thai('ศูนย์ส่งเสริมโอลิมปิกวิชาการและพัฒนามาตรฐานวิทยาศาสตร์ศึกษา')
        l2 = shape_thai('ในพระอุปถัมภ์สมเด็จพระเจ้าพี่นางเธอ เจ้าฟ้ากัลยาณิวัฒนา กรมหลวงนราธิวาสราชนครินทร์ (สอวน.)')
        l3 = shape_thai(
            f'สาขาวิชาคอมพิวเตอร์ ค่าย {camp} ศูนย์โรงเรียนเทพศิรินทร์   {sheet.get("date_thai_header", "")}'
        )
        title_l4 = sheet.get('title_line4') or (
            f'{"แบบลงชื่อเข้าสอบ" if sheet.get("day_mode") == "exam" else "แบบลงชื่อเข้าอบรม"} '
            f'({sheet.get("period_label", "")})'
        )
        if sheet.get('total_pages', 1) > 1:
            title_l4 += f' — หน้า {sheet["page_no"]}/{sheet["total_pages"]}'
        l4 = shape_thai(title_l4)

        c.setFont(font_bold, 15.0)
        c.drawCentredString(cx, top_y - 4.8 * mm, l1)
        c.setFont(font_bold, 13.0)
        c.drawCentredString(cx, top_y - 10.0 * mm, l2)
        c.setFont(font_bold, 13.8)
        c.drawCentredString(cx, top_y - 15.4 * mm, l3)
        c.setFont(font_bold, 13.8)
        l4_y = top_y - 21.0 * mm
        c.drawCentredString(cx, l4_y, l4)
        l4_w = c.stringWidth(l4, font_bold, 13.8)
        c.setLineWidth(0.9)
        c.line(cx - l4_w / 2.0 - 2 * mm, l4_y - 1.2 * mm, cx + l4_w / 2.0 + 2 * mm, l4_y - 1.2 * mm)

        # ---------- ส่วนตาราง 12 คอลัมน์ (50–52 คน/แผ่น) ----------
        table_top = top_y - header_h - 1.0 * mm
        thead_h = 9.5 * mm
        footer_reserve = 4.0 * mm
        paired = sheet.get('paired_rows') or []
        n_rows = max(len(paired), 1)
        avail_body_h = (table_top - thead_h) - (margin_bottom + footer_reserve)
        row_h = min(6.5 * mm, avail_body_h / float(max(n_rows, 25)))
        table_bottom = table_top - thead_h - (row_h * n_rows)

        # พื้นหลังหัวตาราง
        c.setFillColor(colors.HexColor('#eef1f4'))
        c.rect(margin_x, table_top - thead_h, usable_w, thead_h, fill=1, stroke=0)
        c.setFillColor(colors.black)

        # ข้อความหัวตาราง (ซ้าย + ขวา)
        headers_2line = [
            ('ลำดับ', ''),
            ('เลขประจำตัว', 'ค่าย'),
            ('ชื่อ-นามสกุล', ''),
            ('ระดับ', 'ชั้น'),
            ('โรงเรียน', ''),
            ('ลงชื่อ', ''),
        ]
        c.setFont(font_bold, 11.5)
        for side_offset in (0, 6):
            for idx_col, (h1, h2) in enumerate(headers_2line):
                ci = side_offset + idx_col
                mid_x = (col_xs[ci] + col_xs[ci + 1]) / 2.0
                if h2:
                    c.drawCentredString(mid_x, table_top - 4.1 * mm, shape_thai(h1))
                    c.drawCentredString(mid_x, table_top - 7.8 * mm, shape_thai(h2))
                else:
                    c.drawCentredString(mid_x, table_top - 6.0 * mm, shape_thai(h1))

        # วาดแถวข้อมูล
        if paired:
            pad_x = 1.4 * mm
            for r_idx, pair in enumerate(paired):
                y_top_row = table_top - thead_h - (r_idx * row_h)
                y_base = y_top_row - row_h + (row_h - 3.0 * mm) / 2.0 + 0.3 * mm

                # สลับสีแถวอ่อนมากเพื่อให้อ่านไล่บรรทัดง่าย
                if r_idx % 2 == 1:
                    c.setFillColor(colors.HexColor('#fafbfc'))
                    c.rect(margin_x, y_top_row - row_h, usable_w, row_h, fill=1, stroke=0)
                    c.setFillColor(colors.black)

                for side_offset, item in ((0, pair.get('left')), (6, pair.get('right'))):
                    if not item:
                        continue
                    # 0: ลำดับ
                    c.setFont(font_reg, 11.5)
                    c.drawCentredString((col_xs[side_offset + 0] + col_xs[side_offset + 1]) / 2.0,
                                        y_base, str(item.get('no', '')))
                    # 1: เลขประจำตัวค่าย
                    c.setFont(font_bold, 11.2)
                    code_txt = _fit_text(c, item.get('code', ''), font_bold, 11.2, col_widths[side_offset + 1] - 2 * pad_x)
                    c.drawCentredString((col_xs[side_offset + 1] + col_xs[side_offset + 2]) / 2.0, y_base, code_txt)
                    # 2: ชื่อ-นามสกุล
                    c.setFont(font_reg, 11.3)
                    name_txt = _fit_text(c, item.get('full_name', ''), font_reg, 11.3, col_widths[side_offset + 2] - 2 * pad_x)
                    c.drawString(col_xs[side_offset + 2] + pad_x, y_base, name_txt)
                    # 3: ระดับชั้น
                    c.setFont(font_reg, 11.0)
                    grade_txt = _fit_text(c, item.get('grade', ''), font_reg, 11.0, col_widths[side_offset + 3] - 1.2 * mm)
                    c.drawCentredString((col_xs[side_offset + 3] + col_xs[side_offset + 4]) / 2.0, y_base, grade_txt)
                    # 4: โรงเรียน
                    c.setFont(font_reg, 10.3)
                    school_txt = _fit_text(c, item.get('school', ''), font_reg, 10.3, col_widths[side_offset + 4] - 2 * pad_x)
                    c.drawString(col_xs[side_offset + 4] + pad_x, y_base, school_txt)
                    # 5: ลงชื่อ (เว้นว่างให้เซ็น)
        else:
            c.setFont(font_reg, 12.0)
            c.setFillColor(colors.HexColor('#6c757d'))
            c.drawCentredString(cx, table_top - thead_h - row_h / 2.0 - 1 * mm,
                                shape_thai('ไม่พบรายชื่อนักเรียนในกลุ่มนี้'))
            c.setFillColor(colors.black)

        # เส้นกรอบตารางทั้งหมด
        c.setStrokeColor(colors.HexColor('#333333'))
        c.setLineWidth(0.55)
        # เส้นแนวนอน
        c.line(margin_x, table_top, margin_x + usable_w, table_top)
        c.line(margin_x, table_top - thead_h, margin_x + usable_w, table_top - thead_h)
        for r_idx in range(1, n_rows + 1):
            y_line = table_top - thead_h - (r_idx * row_h)
            c.line(margin_x, y_line, margin_x + usable_w, y_line)
        # เส้นแนวตั้ง
        for ci, x_pos in enumerate(col_xs):
            if ci == 6:
                continue  # วาดเส้นหนาแยกต่างหาก
            c.line(x_pos, table_top, x_pos, table_bottom)
        # เส้นแบ่งกึ่งกลางซ้าย-ขวา (หนาชัดเจน)
        c.setStrokeColor(colors.black)
        c.setLineWidth(1.4)
        c.line(split_x, table_top, split_x, table_bottom)

        # ท้ายกระดาษ
        c.setFont(font_reg, 9.5)
        c.setFillColor(colors.HexColor('#555555'))
        footer_left = shape_thai(
            f'กลุ่มรายชื่อ: {sheet.get("group_label", "นักเรียนปัจจุบัน")} · จำนวน {sheet.get("count", 0)} คน'
        )
        footer_right = shape_thai(f'แผ่นที่ {s_idx} / {total_sheets}')
        c.drawString(margin_x, margin_bottom, footer_left)
        c.drawRightString(page_w - margin_x, margin_bottom, footer_right)

    c.save()
    return buf.getvalue()
