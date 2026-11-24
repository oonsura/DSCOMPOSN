"""
ส่งออกผังที่นั่งสอบเป็น PDF (THSarabunNew ตัวหนาชัดเจน ขนาดกล่องที่นั่งเท่ากันทุกหน้า) และ Excel (.xlsx)
- PDF: หน้าผังที่นั่งสอบใกล้เคียงหน้าเว็บ (แนวนอนสำหรับห้องคอมฯ / แนวตั้งสำหรับห้องประชุม) ขนาดที่นั่งอ้างอิงจากห้องที่มีแถวยาวที่สุดให้เท่ากันทุกแผ่น
- Excel: ชีตรายชื่อรวม + ชีตแผนผังที่นั่งสอบแต่ละห้องพร้อมสีและเส้นตาราง
"""
import io
import re

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from composn.templatetags.composn_tags import format_date_thai_long
from . import signin_pdf as _signin_pdf
from .exam_seating_service import build_room_layout_data
from .models import ExamRoom, ExamSeat, ExamSeatAssignment
from .signin_pdf import _ensure_thai_fonts, _fit_text, shape_thai

_CHART_FONT_REG = None
_CHART_FONT_BOLD = None
_CHART_FONT_IS_PUA = False
_CHART_FONT_SCALE = 1.0

_THAI_UP_TAIL = set('ปฝฟฬ')
_THAI_LOW_TAIL = set('ฎฏ')
_THAI_UPPER_VOWELS = set('\u0e31\u0e34\u0e35\u0e36\u0e37\u0e47\u0e4d')
_THAI_LOWER_VOWELS = set('\u0e38\u0e39\u0e3a')
_THAI_TONES = set('\u0e48\u0e49\u0e4a\u0e4b\u0e4c')
_THAI_COMBINING = _THAI_UPPER_VOWELS | _THAI_LOWER_VOWELS | _THAI_TONES


def _ensure_chart_fonts():
    """ใช้ฟอนต์ THSarabunNew (ปกติ/หนา) เป็นหลักสำหรับผังที่นั่งสอบ PDF"""
    global _CHART_FONT_REG, _CHART_FONT_BOLD, _CHART_FONT_IS_PUA, _CHART_FONT_SCALE
    if _CHART_FONT_REG and _CHART_FONT_BOLD:
        return _CHART_FONT_REG, _CHART_FONT_BOLD, _CHART_FONT_IS_PUA, _CHART_FONT_SCALE

    reg_name, bold_name = _ensure_thai_fonts()
    _CHART_FONT_REG = reg_name
    _CHART_FONT_BOLD = bold_name
    _CHART_FONT_IS_PUA = bool(getattr(_signin_pdf, '_USE_THAI_PUA', False))
    _CHART_FONT_SCALE = 1.0
    return _CHART_FONT_REG, _CHART_FONT_BOLD, _CHART_FONT_IS_PUA, _CHART_FONT_SCALE


def _measure_chart_text(c, text, font_name, sz, is_pua):
    s = shape_thai(str(text)) if is_pua else str(text)
    if not s:
        return 0.0
    return c.stringWidth(s, font_name, sz)


def _fit_chart_text(c, text, font_name, base_size, max_w, is_pua, scale=1.0, min_size=6.5):
    raw = str(text or '')
    sz = float(base_size) * float(scale)
    min_sz = float(min_size) * float(scale)
    if not raw:
        return '', sz
    if not max_w or max_w <= 0:
        return (shape_thai(raw) if is_pua else raw), sz

    while sz > min_sz and _measure_chart_text(c, raw, font_name, sz, is_pua) > max_w:
        sz = max(min_sz, sz - 0.25)

    if _measure_chart_text(c, raw, font_name, sz, is_pua) <= max_w:
        return (shape_thai(raw) if is_pua else raw), sz

    trimmed = raw
    while trimmed and _measure_chart_text(c, f'{trimmed}…', font_name, sz, is_pua) > max_w:
        trimmed = trimmed[:-1]
        while trimmed and trimmed[-1] in _THAI_COMBINING:
            trimmed = trimmed[:-1]
    candidate = f'{trimmed}…' if trimmed else '…'
    return (shape_thai(candidate) if is_pua else candidate), sz


def _draw_thai_positioned(c, text, x, y, font_name, sz):
    """วาดข้อความภาษาไทยพร้อมปรับตำแหน่งสระบน-ล่างและวรรณยุกต์สำหรับฟอนต์ที่ไม่ใช้ PUA"""
    cursor_x = x
    last_base = ''
    has_upper_vowel = False
    n = len(text)
    c.setFont(font_name, sz)
    for idx, ch in enumerate(text):
        next_ch = text[idx + 1] if idx + 1 < n else ''
        dx = 0.0
        dy = 0.0
        if ch in _THAI_UPPER_VOWELS:
            if last_base in _THAI_UP_TAIL:
                dx = -0.14 * sz
            has_upper_vowel = True
        elif ch in _THAI_TONES:
            if last_base in _THAI_UP_TAIL:
                dx = -0.14 * sz
            if has_upper_vowel or next_ch == '\u0e33':
                dy = 0.25 * sz
        elif ch in _THAI_LOWER_VOWELS:
            if last_base in _THAI_LOW_TAIL:
                dy = -0.18 * sz
        else:
            last_base = ch
            has_upper_vowel = False

        c.drawString(cursor_x + dx, y + dy, ch)
        cursor_x += c.stringWidth(ch, font_name, sz)


def _draw_chart_text(
    c,
    text,
    x,
    cy,
    font_name,
    font_size,
    is_pua,
    scale=1.0,
    align='center',
    max_w=None,
    color=colors.black,
    min_size=6.5,
    stroke_ratio=0.0,
):
    fitted, used_sz = _fit_chart_text(
        c, text, font_name, font_size, max_w, is_pua, scale=scale, min_size=min_size,
    )
    if not fitted:
        return 0.0
    w = c.stringWidth(fitted, font_name, used_sz)
    if align == 'center':
        start_x = x - w / 2.0
    elif align == 'right':
        start_x = x - w
    else:
        start_x = x
    baseline_y = cy - used_sz * (0.23 if is_pua else 0.34)

    c.saveState()
    c.setFillColor(color)
    if is_pua:
        if stroke_ratio > 0:
            c.setStrokeColor(color)
            c.setLineWidth(used_sz * stroke_ratio)
            c.setLineJoin(1)
            t = c.beginText(start_x, baseline_y)
            t.setFont(font_name, used_sz)
            t.setFillColor(color)
            t.setTextRenderMode(2)
            t.textOut(fitted)
            c.drawText(t)
        else:
            c.setFont(font_name, used_sz)
            c.drawString(start_x, baseline_y, fitted)
    else:
        _draw_thai_positioned(c, fitted, start_x, baseline_y, font_name, used_sz)
    c.restoreState()
    return w


def _sanitize_sheet_title(title, used_titles):
    cleaned = re.sub(r'[:\\/?*\[\]]', ' ', str(title or 'Room')).strip() or 'Room'
    base = cleaned[:31]
    candidate = base
    idx = 2
    while candidate in used_titles:
        suffix = f' ({idx})'
        candidate = base[:31 - len(suffix)] + suffix
        idx += 1
    used_titles.add(candidate)
    return candidate


def _fit_text_and_size(c, text, font_name, font_size, max_width, min_size=6.5):
    """ปรับขนาดตัวอักษรให้พอดีความกว้าง ก่อนตัดข้อความด้วย …"""
    shaped = shape_thai(str(text))
    if not shaped:
        return '', font_size
    sz = float(font_size)
    while sz > min_size and c.stringWidth(shaped, font_name, sz) > max_width:
        sz -= 0.5
    if c.stringWidth(shaped, font_name, sz) <= max_width:
        return shaped, sz
    return _fit_text(c, str(text), font_name, sz, max_width), sz


def _draw_centered_text(c, text, cx, y, font_name, font_size, max_w=None, color=colors.black):
    c.saveState()
    c.setFillColor(color)
    if max_w:
        fitted, used_size = _fit_text_and_size(c, str(text), font_name, font_size, max_w)
    else:
        fitted, used_size = shape_thai(str(text)), font_size
    c.setFont(font_name, used_size)
    c.drawCentredString(cx, y - used_size * 0.32, fitted)
    c.restoreState()


def _calc_page_geometry(room, layout, has_proctors):
    """คำนวณขอบเขตพื้นที่ผังที่นั่งบนหน้ากระดาษ A4 ของห้องหนึ่งๆ"""
    is_wide = room.has_computer or layout['max_col'] > 6
    pagesize = landscape(A4) if is_wide else A4
    pw, ph = pagesize

    margin_x = (9.0 * mm) if is_wide else (10.0 * mm)
    margin_y = 9.0 * mm
    usable_w = pw - 2 * margin_x
    top_y = ph - margin_y
    chart_top = top_y - 14.5 * mm

    proctor_bar_h = 11.0 * mm if has_proctors else 0.0
    proctor_gap = 4.0 * mm if has_proctors else 0.0
    chart_bottom = margin_y + proctor_bar_h + proctor_gap
    chart_h = max(30.0 * mm, chart_top - chart_bottom)

    front_pos = room.front_position or ExamRoom.FRONT_TOP
    front_thick = 10.5 * mm
    front_bar_w = 11.0 * mm
    front_gap = 3.5 * mm

    grid_x = margin_x
    grid_y = chart_bottom
    grid_w = usable_w
    grid_h = chart_h

    if front_pos == ExamRoom.FRONT_TOP:
        grid_h -= (front_thick + front_gap)
    elif front_pos == ExamRoom.FRONT_BOTTOM:
        grid_y += (front_thick + front_gap)
        grid_h -= (front_thick + front_gap)
    elif front_pos == ExamRoom.FRONT_LEFT:
        grid_x += (front_bar_w + front_gap)
        grid_w -= (front_bar_w + front_gap)
    elif front_pos == ExamRoom.FRONT_RIGHT:
        grid_w -= (front_bar_w + front_gap)

    return {
        'is_wide': is_wide,
        'pagesize': pagesize,
        'pw': pw,
        'ph': ph,
        'margin_x': margin_x,
        'margin_y': margin_y,
        'usable_w': usable_w,
        'top_y': top_y,
        'chart_top': chart_top,
        'proctor_bar_h': proctor_bar_h,
        'proctor_gap': proctor_gap,
        'chart_bottom': chart_bottom,
        'chart_h': chart_h,
        'front_pos': front_pos,
        'front_thick': front_thick,
        'front_bar_w': front_bar_w,
        'front_gap': front_gap,
        'grid_x': grid_x,
        'grid_y': grid_y,
        'grid_w': max(20.0 * mm, grid_w),
        'grid_h': max(20.0 * mm, grid_h),
    }


def _compute_uniform_seat_dimensions(room_entries):
    """คำนวณขนาดกล่องที่นั่งให้เท่ากันทุกหน้า โดยเริ่มพิจารณาจากห้องที่มีแถวยาวที่สุด"""
    w_fits = []
    h_fits = []
    for entry in room_entries:
        room = entry['plan_room'].room
        layout = entry['layout']
        geom = entry['geom']
        blocks = layout.get('blocks') or []
        max_col = layout.get('max_col') or 0
        total_rows = sum(len(b['rows']) for b in blocks)
        num_blocks = len(blocks)
        if not blocks or max_col <= 0 or total_rows <= 0:
            continue

        if room.has_computer:
            w_fit = geom['grid_w'] / max_col
            min_block_gap = 6.0 * mm if num_blocks <= 4 else 4.0 * mm
            h_fit = (geom['grid_h'] - min_block_gap * max(0, num_blocks - 1)) / total_rows
        else:
            row_label_w = 16.0 * mm
            label_gap = 2.5 * mm
            min_col_gap = 4.0 * mm
            w_fit = (geom['grid_w'] - row_label_w - label_gap - min_col_gap * max(0, max_col - 1)) / max_col
            min_row_gap = 2.4 * mm
            h_fit = (geom['grid_h'] - min_row_gap * max(0, total_rows - 1)) / total_rows

        w_fits.append(w_fit)
        h_fits.append(h_fit)

    if not w_fits:
        cell_w = 18.0 * mm
        seat_unit_h = 15.5 * mm
    else:
        cell_w = min(36.0 * mm, min(w_fits))
        seat_unit_h = min(cell_w * 1.15, 22.0 * mm, min(h_fits))

    return {
        'cell_w': cell_w,
        'seat_unit_h': seat_unit_h,
        'num_h': seat_unit_h * 0.36,
        'code_h': seat_unit_h * 0.64,
    }


def _draw_seat_box(
    c,
    cx,
    row_top,
    cell_w,
    num_h,
    code_h,
    cell,
    font_reg,
    font_bold,
    is_pua,
    scale,
):
    """วาดกล่องที่นั่ง 1 ช่องด้วยฟอนต์ THSarabunNew ตัวหนาขนาดใหญ่เต็มช่อง"""
    seat = cell['seat']
    status = cell['status']
    total_h = num_h + code_h

    if seat.label_position == ExamSeat.LABEL_ABOVE:
        code_y = row_top - code_h
        num_y = row_top - code_h - num_h
        divider_y = num_y + num_h
    else:
        num_y = row_top - num_h
        code_y = row_top - num_h - code_h
        divider_y = code_y + code_h

    c.saveState()
    # 1. พื้นหลังส่วนข้อมูล
    if status == ExamSeatAssignment.STATUS_BLOCKED:
        c.setFillColor(colors.HexColor('#eeeeee'))
        c.rect(cx, code_y, cell_w, code_h, stroke=0, fill=1)
        c.saveState()
        p = c.beginPath()
        p.rect(cx, code_y, cell_w, code_h)
        c.clipPath(p, stroke=0, fill=0)
        c.setStrokeColor(colors.HexColor('#d8dce0'))
        c.setLineWidth(0.4)
        step = 2.8 * mm
        offset = -code_h
        while offset < cell_w + code_h:
            c.line(cx + offset, code_y, cx + offset + code_h, code_y + code_h)
            offset += step
        c.restoreState()
    else:
        c.setFillColor(colors.white)
        c.rect(cx, code_y, cell_w, code_h, stroke=0, fill=1)

    # 2. แถบเลขเครื่อง/ที่นั่งสีเหลืองอ่อน
    c.setFillColor(colors.HexColor('#fff2cc'))
    c.rect(cx, num_y, cell_w, num_h, stroke=0, fill=1)

    # 3. เส้นคั่นกลางและกรอบนอกคมชัด
    c.setStrokeColor(colors.HexColor('#555555'))
    c.setLineWidth(0.6)
    c.line(cx, divider_y, cx + cell_w, divider_y)

    c.setStrokeColor(colors.HexColor('#333333'))
    c.setLineWidth(0.85)
    c.rect(cx, row_top - total_h, cell_w, total_h, stroke=1, fill=0)

    # 4. เลขเครื่อง/ที่นั่ง
    eff_num_pt = min(15.5, num_h * 1.05)
    _draw_chart_text(
        c,
        seat.seat_label,
        cx + cell_w / 2.0,
        num_y + num_h / 2.0,
        font_bold,
        eff_num_pt,
        is_pua,
        scale=scale,
        align='center',
        max_w=cell_w - 2.0,
        color=colors.HexColor('#111111'),
        min_size=9.5,
        stroke_ratio=0.026,
    )

    # 5. รหัสผู้เข้าสอบ / สำรอง / งดใช้
    if status == ExamSeatAssignment.STATUS_ASSIGNED and cell['display_code']:
        code_color = colors.HexColor('#1b5e20') if cell.get('is_rights') else colors.HexColor('#0000cc')
        eff_code_pt = min(19.5, code_h * 0.72)
        _draw_chart_text(
            c,
            cell['display_code'],
            cx + cell_w / 2.0,
            code_y + code_h / 2.0,
            font_bold,
            eff_code_pt,
            is_pua,
            scale=scale,
            align='center',
            max_w=cell_w - 3.2,
            color=code_color,
            min_size=10.0,
            stroke_ratio=0.032,
        )
    elif status == ExamSeatAssignment.STATUS_RESERVE:
        eff_res_pt = min(18.0, code_h * 0.68)
        _draw_chart_text(
            c,
            'สำรอง',
            cx + cell_w / 2.0,
            code_y + code_h / 2.0,
            font_bold,
            eff_res_pt,
            is_pua,
            scale=scale,
            align='center',
            max_w=cell_w - 3.0,
            color=colors.HexColor('#cc0000'),
            min_size=9.5,
            stroke_ratio=0.030,
        )
    elif status == ExamSeatAssignment.STATUS_BLOCKED:
        eff_blk_pt = min(14.0, code_h * 0.55)
        _draw_chart_text(
            c,
            'งดใช้',
            cx + cell_w / 2.0,
            code_y + code_h / 2.0,
            font_bold,
            eff_blk_pt,
            is_pua,
            scale=scale,
            align='center',
            max_w=cell_w - 3.0,
            color=colors.HexColor('#6c757d'),
            min_size=8.5,
        )
    c.restoreState()


def _draw_front_bar(c, x, y, w, h, label, font_bold, is_pua, scale, vertical_rotate=None):
    c.saveState()
    c.setFillColor(colors.HexColor('#d9ead3'))
    c.setStrokeColor(colors.HexColor('#93c47d'))
    c.setLineWidth(0.8)
    c.roundRect(x, y, w, h, 1.5 * mm, stroke=1, fill=1)
    if vertical_rotate is None:
        _draw_chart_text(
            c,
            label,
            x + w / 2.0,
            y + h / 2.0,
            font_bold,
            21.0,
            is_pua,
            scale=scale,
            align='center',
            max_w=w - 6 * mm,
            color=colors.HexColor('#1b5e20'),
            stroke_ratio=0.028,
        )
    else:
        c.translate(x + w / 2.0, y + h / 2.0)
        c.rotate(vertical_rotate)
        _draw_chart_text(
            c,
            label,
            0.0,
            0.0,
            font_bold,
            21.0,
            is_pua,
            scale=scale,
            align='center',
            max_w=h - 6 * mm,
            color=colors.HexColor('#1b5e20'),
            stroke_ratio=0.028,
        )
    c.restoreState()


def _draw_room_chart_page(
    c,
    plan,
    plan_room,
    layout,
    examinee_map,
    proctors=None,
    geom=None,
    uniform_dims=None,
):
    room = plan_room.room
    font_reg, font_bold, is_pua, scale = _ensure_chart_fonts()

    if proctors is None:
        proctors = list(plan_room.proctors.all().order_by('sort_order', 'id'))
    if geom is None:
        geom = _calc_page_geometry(room, layout, bool(proctors))
    if uniform_dims is None:
        uniform_dims = _compute_uniform_seat_dimensions([
            {'plan_room': plan_room, 'layout': layout, 'proctors': proctors, 'geom': geom}
        ])

    c.setPageSize(geom['pagesize'])
    pw = geom['pw']
    margin_x = geom['margin_x']
    margin_y = geom['margin_y']
    usable_w = geom['usable_w']
    top_y = geom['top_y']
    chart_top = geom['chart_top']
    proctor_bar_h = geom['proctor_bar_h']
    chart_bottom = geom['chart_bottom']
    chart_h = geom['chart_h']
    front_pos = geom['front_pos']
    front_thick = geom['front_thick']
    front_bar_w = geom['front_bar_w']
    grid_x = geom['grid_x']
    grid_y = geom['grid_y']
    grid_w = geom['grid_w']
    grid_h = geom['grid_h']

    # 1. ส่วนหัว: ชื่อห้อง (ซ้าย) + หัวข้อผัง (ขวาหนา) และวันที่สอบ (ขวาปกติสีเทา)
    head_cy = top_y - 4.0 * mm
    date_txt = format_date_thai_long(plan.training_date.date) if plan.training_date else ''
    sub_txt = plan_room.effective_subtitle

    right_date_w = 0.0
    if date_txt:
        date_str = f'({date_txt})'
        right_date_w = _draw_chart_text(
            c,
            date_str,
            pw - margin_x,
            head_cy,
            font_reg,
            16.0,
            is_pua,
            scale=scale,
            align='right',
            color=colors.HexColor('#555555'),
        )
    sub_right_x = (pw - margin_x - right_date_w - 1.5 * mm) if right_date_w else (pw - margin_x)
    sub_w = _draw_chart_text(
        c,
        sub_txt,
        sub_right_x,
        head_cy,
        font_bold,
        19.0,
        is_pua,
        scale=scale,
        align='right',
        max_w=usable_w * 0.48,
        color=colors.HexColor('#111111'),
        stroke_ratio=0.025,
    )

    title_pt = 22.0 if geom['is_wide'] else 21.0
    title_max_w = max(40 * mm, (sub_right_x - sub_w) - margin_x - 6 * mm)
    _draw_chart_text(
        c,
        room.name,
        margin_x,
        head_cy,
        font_bold,
        title_pt,
        is_pua,
        scale=scale,
        align='left',
        max_w=title_max_w,
        color=colors.HexColor('#111111'),
        stroke_ratio=0.025,
    )

    # คำอธิบายสี (Legend) เรียงจากซ้ายไปขวาเหมือนบนหน้าเว็บ
    legend_cy = top_y - 9.8 * mm
    sq_size = 2.2 * mm
    lx = pw - margin_x
    legend_items = (
        ('#0000cc', 'สีน้ำเงิน', ' = นักเรียนปัจจุบัน'),
        ('#1b5e20', 'สีเขียวเข้ม', ' = นักเรียนใช้สิทธิ์'),
        ('#cc0000', 'สีแดง', ' = สำรอง'),
    )
    for hex_c, bold_lbl, rest_lbl in reversed(legend_items):
        w_rest = _draw_chart_text(
            c,
            rest_lbl,
            lx,
            legend_cy,
            font_reg,
            13.5,
            is_pua,
            scale=scale,
            align='right',
            color=colors.HexColor('#222222'),
        )
        lx -= w_rest
        w_bold = _draw_chart_text(
            c,
            bold_lbl,
            lx,
            legend_cy,
            font_bold,
            13.5,
            is_pua,
            scale=scale,
            align='right',
            color=colors.HexColor(hex_c),
            stroke_ratio=0.022,
        )
        lx -= (w_bold + 1.0 * mm + sq_size)
        c.saveState()
        c.setFillColor(colors.HexColor(hex_c))
        c.rect(lx, legend_cy - sq_size * 0.42, sq_size, sq_size, stroke=0, fill=1)
        c.restoreState()
        lx -= 4.5 * mm

    # 2. แถบวิทยากรคุมสอบด้านล่างสุด
    if proctors:
        proctor_text = '    '.join(p.display_name for p in proctors)
        c.saveState()
        c.setFillColor(colors.HexColor('#fce5cd'))
        c.setStrokeColor(colors.HexColor('#e6b8af'))
        c.setLineWidth(0.8)
        c.roundRect(margin_x, margin_y, usable_w, proctor_bar_h, 1.8 * mm, stroke=1, fill=1)
        c.restoreState()
        _draw_chart_text(
            c,
            proctor_text,
            margin_x + usable_w / 2.0,
            margin_y + proctor_bar_h / 2.0,
            font_bold,
            19.0,
            is_pua,
            scale=scale,
            align='center',
            max_w=usable_w - 8 * mm,
            color=colors.HexColor('#111111'),
            stroke_ratio=0.025,
        )

    front_label = room.front_label or ('หน้าเวที' if not room.has_computer else 'หน้าห้อง')
    blocks = layout['blocks']
    max_col = layout['max_col']
    if not blocks or max_col <= 0:
        if front_pos == ExamRoom.FRONT_TOP:
            _draw_front_bar(c, margin_x, chart_top - front_thick, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_BOTTOM:
            _draw_front_bar(c, margin_x, chart_bottom, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_LEFT:
            _draw_front_bar(c, margin_x, chart_bottom, front_bar_w, chart_h, front_label, font_bold, is_pua, scale, vertical_rotate=90)
        elif front_pos == ExamRoom.FRONT_RIGHT:
            _draw_front_bar(c, margin_x + usable_w - front_bar_w, chart_bottom, front_bar_w, chart_h, front_label, font_bold, is_pua, scale, vertical_rotate=270)
        _draw_chart_text(
            c,
            'ยังไม่มีข้อมูลที่นั่งในห้องสอบนี้',
            grid_x + grid_w / 2.0,
            grid_y + grid_h / 2.0,
            font_bold,
            18.0,
            is_pua,
            scale=scale,
            align='center',
            color=colors.grey,
        )
        return

    # 3. วาดผังที่นั่งด้วยขนาดกล่องเท่ากันทุกแผ่น + แถบหน้าห้อง/หน้าเวที
    cell_w = uniform_dims['cell_w']
    seat_unit_h = uniform_dims['seat_unit_h']
    num_h = uniform_dims['num_h']
    code_h = uniform_dims['code_h']
    total_rows = sum(len(b['rows']) for b in blocks)
    num_blocks = len(blocks)

    if room.has_computer:
        if num_blocks > 1:
            raw_gap = (grid_h - total_rows * seat_unit_h) / (num_blocks + 1)
            max_possible_gap = max(0.0, (grid_h - total_rows * seat_unit_h) / (num_blocks - 1))
            block_gap = min(20.0 * mm, max(6.0 * mm, raw_gap), max_possible_gap)
        else:
            block_gap = 0.0

        used_w = max_col * cell_w
        if front_pos == ExamRoom.FRONT_LEFT:
            row_start_x = grid_x
        elif front_pos == ExamRoom.FRONT_RIGHT:
            row_start_x = grid_x + max(0.0, grid_w - used_w)
        else:
            row_start_x = grid_x + max(0.0, (grid_w - used_w) / 2.0)

        actual_blocks_h = total_rows * seat_unit_h + max(0, num_blocks - 1) * block_gap
        blocks_top = grid_y + (grid_h + actual_blocks_h) / 2.0
        blocks_bottom = blocks_top - actual_blocks_h

        # แถบหน้าห้อง / หน้าเวที
        if front_pos == ExamRoom.FRONT_TOP:
            _draw_front_bar(c, margin_x, chart_top - front_thick, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_BOTTOM:
            _draw_front_bar(c, margin_x, chart_bottom, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_LEFT:
            _draw_front_bar(c, margin_x, blocks_bottom, front_bar_w, actual_blocks_h, front_label, font_bold, is_pua, scale, vertical_rotate=90)
        elif front_pos == ExamRoom.FRONT_RIGHT:
            _draw_front_bar(c, margin_x + usable_w - front_bar_w, blocks_bottom, front_bar_w, actual_blocks_h, front_label, font_bold, is_pua, scale, vertical_rotate=270)

        cur_top = blocks_top
        for b_idx, blk in enumerate(blocks):
            if b_idx > 0:
                cur_top -= block_gap
            for r in blk['rows']:
                for c_idx, cell in enumerate(r['cells']):
                    if cell is None:
                        continue
                    _draw_seat_box(
                        c,
                        row_start_x + c_idx * cell_w,
                        cur_top,
                        cell_w,
                        num_h,
                        code_h,
                        cell,
                        font_reg,
                        font_bold,
                        is_pua,
                        scale,
                    )
                cur_top -= seat_unit_h
    else:
        # ห้องประชุม / โต๊ะสอบ
        row_label_w = 16.0 * mm
        label_gap = 2.5 * mm
        desk_area_w = grid_w - row_label_w - label_gap
        if max_col > 1:
            raw_cgap = (desk_area_w - max_col * cell_w) / max(1, max_col)
            max_cgap = max(0.0, (desk_area_w - max_col * cell_w) / (max_col - 1))
            col_gap = min(14.0 * mm, max(4.0 * mm, raw_cgap), max_cgap)
        else:
            col_gap = 0.0

        used_desks_w = max_col * cell_w + max(0, max_col - 1) * col_gap
        total_group_w = row_label_w + label_gap + used_desks_w
        group_start_x = grid_x + max(0.0, (grid_w - total_group_w) / 2.0)
        row_label_cx = group_start_x + row_label_w / 2.0
        desk_start_x = group_start_x + row_label_w + label_gap

        if total_rows > 1:
            raw_rgap = (grid_h - total_rows * seat_unit_h) / (total_rows - 1)
            row_gap = min(4.0 * mm, max(2.4 * mm, raw_rgap), max(0.0, raw_rgap))
        else:
            row_gap = 0.0

        actual_rows_h = total_rows * seat_unit_h + max(0, total_rows - 1) * row_gap
        rows_top = grid_y + (grid_h + actual_rows_h) / 2.0
        rows_bottom = rows_top - actual_rows_h

        # แถบหน้าเวที / หน้าห้อง
        if front_pos == ExamRoom.FRONT_TOP:
            _draw_front_bar(c, margin_x, chart_top - front_thick, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_BOTTOM:
            _draw_front_bar(c, margin_x, chart_bottom, usable_w, front_thick, front_label, font_bold, is_pua, scale)
        elif front_pos == ExamRoom.FRONT_LEFT:
            _draw_front_bar(c, margin_x, rows_bottom, front_bar_w, actual_rows_h, front_label, font_bold, is_pua, scale, vertical_rotate=90)
        elif front_pos == ExamRoom.FRONT_RIGHT:
            _draw_front_bar(c, margin_x + usable_w - front_bar_w, rows_bottom, front_bar_w, actual_rows_h, front_label, font_bold, is_pua, scale, vertical_rotate=270)

        cur_top = rows_top
        for blk in blocks:
            for r in blk['rows']:
                row_cy = cur_top - seat_unit_h / 2.0
                # ป้ายแถว 2 บรรทัด: "แถวที่" + เลขแถวสีน้ำเงิน
                _draw_chart_text(
                    c,
                    'แถวที่',
                    row_label_cx,
                    row_cy + seat_unit_h * 0.17,
                    font_bold,
                    12.5,
                    is_pua,
                    scale=scale,
                    align='center',
                    max_w=row_label_w - 2.0,
                    color=colors.HexColor('#222222'),
                    stroke_ratio=0.015,
                )
                _draw_chart_text(
                    c,
                    str(r['row_index']),
                    row_label_cx,
                    row_cy - seat_unit_h * 0.18,
                    font_bold,
                    17.5,
                    is_pua,
                    scale=scale,
                    align='center',
                    max_w=row_label_w - 2.0,
                    color=colors.HexColor('#0000cc'),
                    stroke_ratio=0.028,
                )

                for c_idx, cell in enumerate(r['cells']):
                    if cell is None:
                        continue
                    _draw_seat_box(
                        c,
                        desk_start_x + c_idx * (cell_w + col_gap),
                        cur_top,
                        cell_w,
                        num_h,
                        code_h,
                        cell,
                        font_reg,
                        font_bold,
                        is_pua,
                        scale,
                    )
                cur_top -= (seat_unit_h + row_gap)


def _draw_room_roster_pages(c, plan, plan_room, examinee_map):
    """วาดหน้ารายชื่อผู้เข้าสอบประจำห้อง (เรียงตามเลขที่นั่ง/เครื่อง)"""
    font_reg, font_bold = _ensure_thai_fonts()
    assignments = list(
        ExamSeatAssignment.objects.filter(
            plan_room=plan_room,
            status=ExamSeatAssignment.STATUS_ASSIGNED,
            user__isnull=False,
        ).select_related('seat', 'user').order_by('seat__seat_number', 'seat__seat_label')
    )
    if not assignments:
        return

    pagesize = A4
    pw, ph = pagesize
    margin_x = 14 * mm
    margin_y = 14 * mm
    usable_w = pw - 2 * margin_x
    col_widths = [14 * mm, 24 * mm, 28 * mm, 52 * mm, 46 * mm, 18 * mm]
    headers = ['ลำดับ', 'เลขเครื่อง/โต๊ะ', 'รหัสผู้เข้าสอบ', 'ชื่อ - นามสกุล', 'โรงเรียน', 'กลุ่ม']
    row_h = 6.8 * mm
    rows_per_page = 34

    date_txt = format_date_thai_long(plan.training_date.date) if plan.training_date else ''
    proctors = '  '.join(p.display_name for p in plan_room.proctors.all().order_by('sort_order', 'id'))

    for chunk_start in range(0, len(assignments), rows_per_page):
        c.showPage()
        c.setPageSize(pagesize)
        chunk = assignments[chunk_start:chunk_start + rows_per_page]
        top_y = ph - margin_y

        c.setFillColor(colors.black)
        c.setFont(font_bold, 15)
        c.drawString(margin_x, top_y - 4 * mm, shape_thai(f'รายชื่อผู้เข้าสอบ: {plan_room.room.name}'))
        c.setFont(font_reg, 12)
        c.drawRightString(pw - margin_x, top_y - 4 * mm, shape_thai(f'{plan_room.effective_subtitle} · {date_txt}'))
        if proctors:
            c.drawString(margin_x, top_y - 10 * mm, shape_thai(f'กรรมการคุมสอบ: {proctors}'))

        table_top = top_y - (15 * mm if proctors else 10 * mm)
        # Header
        cx = margin_x
        c.setFillColor(colors.HexColor('#e8eef7'))
        c.setStrokeColor(colors.HexColor('#666666'))
        c.setLineWidth(0.5)
        c.rect(margin_x, table_top - row_h, sum(col_widths), row_h, stroke=1, fill=1)
        for w, h_txt in zip(col_widths, headers):
            c.rect(cx, table_top - row_h, w, row_h, stroke=1, fill=0)
            _draw_centered_text(c, h_txt, cx + w / 2.0, table_top - row_h / 2.0, font_bold, 11, max_w=w - 2)
            cx += w

        cur_y = table_top - row_h
        for i, asgn in enumerate(chunk, start=chunk_start + 1):
            cur_y -= row_h
            info = examinee_map.get(asgn.user_id, {})
            code = info.get('display_code') or asgn.user.camp1_code
            full_name = info.get('full_name') or asgn.user.get_full_name()
            school = info.get('school_name') or (asgn.user.school_name or '-')
            grp = info.get('group_label') or 'ปัจจุบัน'

            vals = [str(i), asgn.seat.seat_label, code, full_name, school, grp]
            cx = margin_x
            c.setFillColor(colors.HexColor('#fafafa') if i % 2 == 0 else colors.white)
            c.rect(margin_x, cur_y, sum(col_widths), row_h, stroke=0, fill=1)
            for idx_col, (w, val) in enumerate(zip(col_widths, vals)):
                c.setStrokeColor(colors.HexColor('#888888'))
                c.rect(cx, cur_y, w, row_h, stroke=1, fill=0)
                if idx_col in (3, 4):
                    c.saveState()
                    c.setFillColor(colors.black)
                    fitted, used_sz = _fit_text_and_size(c, val, font_reg, 10.5, w - 3 * mm)
                    c.setFont(font_reg, used_sz)
                    c.drawString(cx + 1.5 * mm, cur_y + row_h / 2.0 - used_sz * 0.32, fitted)
                    c.restoreState()
                else:
                    fn = font_bold if idx_col in (1, 2) else font_reg
                    col_color = colors.HexColor('#0000cc') if idx_col == 2 else colors.black
                    _draw_centered_text(c, val, cx + w / 2.0, cur_y + row_h / 2.0, fn, 10.5, max_w=w - 2, color=col_color)
                cx += w


def build_exam_seating_pdf(plan, plan_rooms, examinee_map):
    """สร้างไฟล์ PDF ผังที่นั่งสอบรวมทุกห้องที่เปิดใช้ (ห้องละ 1 แผ่น A4 ขนาดกล่องที่นั่งเท่ากันทุกแผ่น)"""
    font_reg, _, is_pua, scale = _ensure_chart_fonts()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A4))

    code_by_user_id = {uid: info['display_code'] for uid, info in examinee_map.items()}
    rights_user_ids = {uid for uid, info in examinee_map.items() if info.get('is_rights')}

    room_entries = []
    for pr in plan_rooms:
        asgn_by_seat = {
            a.seat_id: a
            for a in ExamSeatAssignment.objects.filter(plan_room=pr).select_related('user', 'seat')
        }
        layout = build_room_layout_data(
            pr.room,
            assignment_by_seat_id=asgn_by_seat,
            code_by_user_id=code_by_user_id,
            rights_user_ids=rights_user_ids,
        )
        proctors = list(pr.proctors.all().order_by('sort_order', 'id'))
        geom = _calc_page_geometry(pr.room, layout, bool(proctors))
        room_entries.append({
            'plan_room': pr,
            'layout': layout,
            'proctors': proctors,
            'geom': geom,
        })

    if not room_entries:
        _draw_chart_text(
            c,
            'ไม่มีห้องสอบที่เปิดแสดงผล',
            30 * mm,
            150 * mm,
            font_reg,
            18.0,
            is_pua,
            scale=scale,
            align='left',
            color=colors.black,
        )
    else:
        uniform_dims = _compute_uniform_seat_dimensions(room_entries)
        for idx, entry in enumerate(room_entries):
            if idx > 0:
                c.showPage()
            _draw_room_chart_page(
                c,
                plan,
                entry['plan_room'],
                entry['layout'],
                examinee_map,
                proctors=entry['proctors'],
                geom=entry['geom'],
                uniform_dims=uniform_dims,
            )

    c.save()
    buf.seek(0)
    return buf.getvalue()


def build_exam_seating_xlsx(plan, plan_rooms, examinee_map):
    """สร้างไฟล์ Excel ผังที่นั่งสอบ (ชีตรายชื่อรวม + ชีตผังที่นั่งแยกตามห้อง ห้องละ 1 แผ่น A4)"""
    wb = Workbook()
    ws_list = wb.active
    used_titles = set()
    ws_list.title = _sanitize_sheet_title('รายชื่อจัดที่นั่งสอบ', used_titles)

    thin = Side(style='thin', color='666666')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_header = PatternFill('solid', fgColor='1A237E')
    fill_yellow = PatternFill('solid', fgColor='FFF2CC')
    fill_green = PatternFill('solid', fgColor='D9EAD3')
    fill_peach = PatternFill('solid', fgColor='FCE5CD')
    fill_blocked = PatternFill('solid', fgColor='EEEEEE')

    font_header = Font(name='TH Sarabun New', size=14, bold=True, color='FFFFFF')
    font_title = Font(name='TH Sarabun New', size=16, bold=True, color='000000')
    font_seat_no = Font(name='TH Sarabun New', size=12, bold=False, color='000000')
    font_code_current = Font(name='TH Sarabun New', size=14, bold=True, color='0000CC')
    font_code_rights = Font(name='TH Sarabun New', size=14, bold=True, color='1B5E20')
    font_reserve = Font(name='TH Sarabun New', size=13, bold=True, color='CC0000')
    font_normal = Font(name='TH Sarabun New', size=12, color='000000')

    date_txt = format_date_thai_long(plan.training_date.date) if plan.training_date else ''
    ws_list.append([f'รายชื่อจัดที่นั่งสอบ ปีการศึกษา {plan.academic_year} ({date_txt}) - {plan.effective_exam_type_display}'])
    ws_list.cell(1, 1).font = font_title
    ws_list.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)

    headers = ['ลำดับ', 'ห้องสอบ', 'เลขเครื่อง/ที่นั่ง', 'รหัสผู้เข้าสอบ', 'ชื่อ - นามสกุล', 'โรงเรียน', 'กลุ่ม', 'กรรมการคุมสอบ']
    ws_list.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws_list.cell(2, col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_all

    row_num = 1
    for pr in plan_rooms:
        proctors_txt = ', '.join(p.display_name for p in pr.proctors.all().order_by('sort_order', 'id'))
        asgns = list(
            ExamSeatAssignment.objects.filter(
                plan_room=pr, status=ExamSeatAssignment.STATUS_ASSIGNED, user__isnull=False,
            ).select_related('seat', 'user').order_by('seat__seat_number', 'seat__seat_label')
        )
        for asgn in asgns:
            info = examinee_map.get(asgn.user_id, {})
            code = info.get('display_code') or asgn.user.camp1_code
            full_name = info.get('full_name') or asgn.user.get_full_name()
            school = info.get('school_name') or (asgn.user.school_name or '')
            grp = info.get('group_label') or 'ปัจจุบัน'
            is_r = bool(info.get('is_rights'))
            ws_list.append([row_num, pr.room.name, asgn.seat.seat_label, code, full_name, school, grp, proctors_txt])
            cur_r = ws_list.max_row
            for c_i in range(1, len(headers) + 1):
                c_obj = ws_list.cell(cur_r, c_i)
                if c_i == 4:
                    c_obj.font = font_code_rights if is_r else font_code_current
                else:
                    c_obj.font = font_normal
                c_obj.border = border_all
                c_obj.alignment = Alignment(
                    horizontal='center' if c_i in (1, 3, 4, 7) else 'left',
                    vertical='center',
                )
            row_num += 1

    col_widths_list = [8, 30, 16, 16, 26, 28, 12, 32]
    for i, w in enumerate(col_widths_list, start=1):
        ws_list.column_dimensions[get_column_letter(i)].width = w
    ws_list.page_setup.paperSize = ws_list.PAPERSIZE_A4
    ws_list.page_setup.orientation = ws_list.ORIENTATION_LANDSCAPE
    ws_list.sheet_properties.pageSetUpPr.fitToPage = True
    ws_list.page_setup.fitToWidth = 1
    ws_list.page_setup.fitToHeight = 0

    # ชีตผังห้องสอบแต่ละห้อง (ทุกห้องที่เปิดใช้ ห้องละ 1 แผ่น A4)
    code_by_user_id = {uid: info['display_code'] for uid, info in examinee_map.items()}
    rights_user_ids = {uid for uid, info in examinee_map.items() if info.get('is_rights')}

    for pr in plan_rooms:
        room = pr.room
        ws = wb.create_sheet(title=_sanitize_sheet_title(room.name, used_titles))
        asgn_by_seat = {
            a.seat_id: a
            for a in ExamSeatAssignment.objects.filter(plan_room=pr).select_related('user', 'seat')
        }
        layout = build_room_layout_data(
            room,
            assignment_by_seat_id=asgn_by_seat,
            code_by_user_id=code_by_user_id,
            rights_user_ids=rights_user_ids,
        )
        max_col = max(1, layout['max_col'])
        front_pos = room.front_position or ExamRoom.FRONT_TOP
        front_label = room.front_label or ('หน้าเวที' if not room.has_computer else 'หน้าห้อง')

        col_offset = 1 if (not room.has_computer or front_pos == ExamRoom.FRONT_LEFT) else 0
        total_grid_cols = col_offset + max_col + (1 if front_pos == ExamRoom.FRONT_RIGHT else 0)
        total_grid_cols = max(total_grid_cols, 4)

        # หัวเรื่อง + คำอธิบายสี
        ws.cell(1, 1, room.name).font = font_title
        sub_cell = ws.cell(1, max(2, total_grid_cols - 2), f'{pr.effective_subtitle} ({date_txt})' if date_txt else pr.effective_subtitle)
        sub_cell.font = font_title

        ws.cell(2, 1, 'สีน้ำเงิน = นักเรียนปัจจุบัน').font = Font(name='TH Sarabun New', size=11, bold=True, color='0000CC')
        ws.cell(2, 3, 'สีเขียวเข้ม = นักเรียนใช้สิทธิ์').font = Font(name='TH Sarabun New', size=11, bold=True, color='1B5E20')

        cur_r = 3
        if front_pos == ExamRoom.FRONT_TOP:
            ws.merge_cells(start_row=cur_r, start_column=col_offset + 1, end_row=cur_r, end_column=col_offset + max_col)
            fc = ws.cell(cur_r, col_offset + 1, front_label)
            fc.fill = fill_green
            fc.font = font_title
            fc.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[cur_r].height = 28
            cur_r += 2

        grid_start_row = cur_r
        for b_idx, blk in enumerate(layout['blocks']):
            if b_idx > 0 and room.has_computer:
                cur_r += 1
            for r in blk['rows']:
                is_above = (r['label_position'] == ExamSeat.LABEL_ABOVE)
                row_code_idx = cur_r if is_above else (cur_r + 1)
                row_num_idx = (cur_r + 1) if is_above else cur_r
                ws.row_dimensions[row_num_idx].height = 20
                ws.row_dimensions[row_code_idx].height = 24

                if not room.has_computer:
                    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r + 1, end_column=1)
                    rc = ws.cell(cur_r, 1, f'แถวที่ {r["row_index"]}')
                    rc.font = Font(name='TH Sarabun New', size=12, bold=True, color='0000CC')
                    rc.alignment = Alignment(horizontal='center', vertical='center')

                for c_idx, cell in enumerate(r['cells']):
                    if cell is None:
                        continue
                    excel_col = col_offset + 1 + c_idx
                    c_num = ws.cell(row_num_idx, excel_col, cell['seat'].seat_label)
                    c_num.fill = fill_yellow
                    c_num.font = font_seat_no
                    c_num.alignment = Alignment(horizontal='center', vertical='center')
                    c_num.border = border_all

                    if cell['status'] == ExamSeatAssignment.STATUS_ASSIGNED and cell['display_code']:
                        c_val = ws.cell(row_code_idx, excel_col, cell['display_code'])
                        c_val.font = font_code_rights if cell.get('is_rights') else font_code_current
                    elif cell['status'] == ExamSeatAssignment.STATUS_RESERVE:
                        c_val = ws.cell(row_code_idx, excel_col, 'สำรอง')
                        c_val.font = font_reserve
                    elif cell['status'] == ExamSeatAssignment.STATUS_BLOCKED:
                        c_val = ws.cell(row_code_idx, excel_col, 'งดใช้')
                        c_val.fill = fill_blocked
                        c_val.font = font_normal
                    else:
                        c_val = ws.cell(row_code_idx, excel_col, '')
                    c_val.alignment = Alignment(horizontal='center', vertical='center')
                    c_val.border = border_all

                cur_r += 2
                if not room.has_computer:
                    cur_r += 1

        grid_end_row = max(grid_start_row, cur_r - 1)
        if front_pos == ExamRoom.FRONT_LEFT:
            ws.merge_cells(start_row=grid_start_row, start_column=1, end_row=grid_end_row, end_column=1)
            fc = ws.cell(grid_start_row, 1, front_label)
            fc.fill = fill_green
            fc.font = font_title
            fc.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
        elif front_pos == ExamRoom.FRONT_RIGHT:
            rcol = col_offset + max_col + 1
            ws.merge_cells(start_row=grid_start_row, start_column=rcol, end_row=grid_end_row, end_column=rcol)
            fc = ws.cell(grid_start_row, rcol, front_label)
            fc.fill = fill_green
            fc.font = font_title
            fc.alignment = Alignment(horizontal='center', vertical='center', text_rotation=180)
        elif front_pos == ExamRoom.FRONT_BOTTOM:
            cur_r += 1
            ws.merge_cells(start_row=cur_r, start_column=col_offset + 1, end_row=cur_r, end_column=col_offset + max_col)
            fc = ws.cell(cur_r, col_offset + 1, front_label)
            fc.fill = fill_green
            fc.font = font_title
            fc.alignment = Alignment(horizontal='center', vertical='center')
            cur_r += 1

        proctors = list(pr.proctors.all().order_by('sort_order', 'id'))
        if proctors:
            cur_r += 1
            ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=total_grid_cols)
            pc = ws.cell(cur_r, 1, '   '.join(p.display_name for p in proctors))
            pc.fill = fill_peach
            pc.font = font_title
            pc.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[cur_r].height = 28

        for c_i in range(1, total_grid_cols + 1):
            ws.column_dimensions[get_column_letter(c_i)].width = 11 if room.has_computer else 13
        if front_pos == ExamRoom.FRONT_LEFT or not room.has_computer:
            ws.column_dimensions['A'].width = 10

        # ตั้งค่าหน้ากระดาษให้พอดี A4 แผ่นเดียวต่อห้อง
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = (
            ws.ORIENTATION_LANDSCAPE if (room.has_computer or max_col > 6) else ws.ORIENTATION_PORTRAIT
        )
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.print_options.horizontalCentered = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
