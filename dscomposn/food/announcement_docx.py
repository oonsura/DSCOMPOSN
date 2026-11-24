"""
สร้างไฟล์ Word (.docx) เอกสารประกาศผลการคัดเลือก
- แผ่นที่ 1: ประกาศราชการ (ตราครุฑ, ข้อความ, ลายเซ็น)
- แผ่นที่ 2+: บัญชีรายชื่อแนบท้ายประกาศ (4 โลโก้ + ตาราง)
"""
import io
import os

from django.conf import settings

try:
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

from composn.templatetags.composn_tags import THAI_MONTHS
from food.announcement_pdf import (
    format_announcement_date_text,
    format_announcement_list_title,
    format_announcement_body_p1,
    format_announcement_body_p2,
    get_announcement_page1_metrics,
    to_arabic_num,
)

THAI_FONT = 'TH Sarabun New'


def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)


def _set_cell_border_all_none(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
        edge_el = OxmlElement(f'w:{edge}')
        edge_el.set(qn('w:val'), 'none')
        edge_el.set(qn('w:sz'), '0')
        edge_el.set(qn('w:color'), 'FFFFFF')
        tcBorders.append(edge_el)
    tcPr.append(tcBorders)


def _set_thai_font(run, size_pt, bold=False, color_rgb=None):
    run.font.name = THAI_FONT
    try:
        rPr = run._element.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), THAI_FONT)
        rFonts.set(qn('w:hAnsi'), THAI_FONT)
        rFonts.set(qn('w:eastAsia'), THAI_FONT)
        rFonts.set(qn('w:cs'), THAI_FONT)
        rPr.insert(0, rFonts)
    except Exception:
        pass
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)


def _add_para(doc_or_cell, text, size_pt=14, bold=False,
              align=WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0.0,
              space_before=0, space_after=3, line_spacing=None):
    """เพิ่ม paragraph พร้อม run"""
    if hasattr(doc_or_cell, 'add_paragraph'):
        para = doc_or_cell.add_paragraph()
    else:
        para = doc_or_cell.add_paragraph()
    para.alignment = align
    if indent_cm:
        para.paragraph_format.first_line_indent = Cm(indent_cm)
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    if line_spacing:
        para.paragraph_format.line_spacing = Pt(line_spacing)
    run = para.add_run(text)
    _set_thai_font(run, size_pt, bold=bold)
    return para, run


def _add_cell_text(cell, text, size_pt=11, bold=False,
                   align=WD_ALIGN_PARAGRAPH.CENTER):
    for para in cell.paragraphs:
        para.text = ''
    if cell.paragraphs:
        para = cell.paragraphs[0]
    else:
        para = cell.add_paragraph()
    para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(str(text) if text is not None else '')
    _set_thai_font(run, size_pt, bold=bold)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _load_image_path(field):
    try:
        if field:
            path = field.path
            if path and os.path.exists(path):
                return path
    except Exception:
        pass
    return None


def _load_garuda_path(setting):
    if _load_image_path(getattr(setting, 'logo_garuda', None)):
        return setting.logo_garuda.path
    for rel in ('garuda.png', 'garuda.jpg'):
        p = os.path.join(settings.BASE_DIR, 'static', 'img', rel)
        if os.path.exists(p):
            return p
    return None


def build_announcement_docx(rows, year, setting):
    """
    สร้าง Word (.docx) ประกาศผลการคัดเลือก
    คืนค่า bytes ของไฟล์ docx
    """
    if not HAS_DOCX:
        return b''

    doc = Document()

    # ตั้งขนาดหน้า A4 แนวตั้ง และระยะขอบตามการตั้งค่า (กั้นหน้า-กั้นหลัง)
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    top_m = float(getattr(setting, 'announcement_margin_top_mm', 25.0) or 25.0) / 10.0
    bottom_m = float(getattr(setting, 'announcement_margin_bottom_mm', 25.0) or 25.0) / 10.0
    left_m = float(getattr(setting, 'announcement_margin_left_mm', 25.0) or 25.0) / 10.0
    right_m = float(getattr(setting, 'announcement_margin_right_mm', 25.0) or 25.0) / 10.0
    section.top_margin = Cm(top_m)
    section.bottom_margin = Cm(bottom_m)
    section.left_margin = Cm(left_m)
    section.right_margin = Cm(right_m)

    fs_title1 = float(getattr(setting, 'announcement_font_size_title1', 17.0) or 17.0)
    fs_title2 = float(getattr(setting, 'announcement_font_size_title2', 15.0) or 15.0)
    fs_title3 = float(getattr(setting, 'announcement_font_size_title3', 14.0) or 14.0)
    font_body = float(getattr(setting, 'announcement_font_size_body', 14.5) or 14.5)
    fs_dir_name = float(getattr(setting, 'announcement_font_size_director_name', 14.5) or 14.5)
    fs_dir_pos = float(getattr(setting, 'announcement_font_size_director_pos', 13.0) or 13.0)
    fs_list1 = float(getattr(setting, 'announcement_font_size_list_title1', 14.0) or 14.0)
    fs_list2 = float(getattr(setting, 'announcement_font_size_list_title2', 12.0) or 12.0)
    font_table = float(getattr(setting, 'announcement_font_size_table', 10.0) or 10.0)
    line_spacing_mult = float(getattr(setting, 'announcement_line_spacing', 1.85) or 1.85)
    line_spacing_pt = font_body * 0.44 * line_spacing_mult * (72.0 / 25.4)
    _raw_para_sp = getattr(setting, 'announcement_paragraph_spacing_mm', None)
    para_spacing_mm = 4.0 if _raw_para_sp is None else max(0.0, float(_raw_para_sp))
    para_spacing_pt = para_spacing_mm * (72.0 / 25.4)
    indent_cm = float(getattr(setting, 'announcement_indent_mm', 25.0) or 25.0) / 10.0
    garuda_cm = float(getattr(setting, 'announcement_garuda_height_mm', 30.0) or 30.0) / 10.0

    passed_count = len(rows)

    # ============ แผ่นที่ 1: ประกาศราชการ ============

    # ตราครุฑ
    garuda_path = _load_garuda_path(setting)
    if garuda_path:
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run()
            r.add_picture(garuda_path, height=Cm(garuda_cm))
        except Exception:
            pass

    metrics = get_announcement_page1_metrics(setting)

    # หัวข้อประกาศ (ใช้ระยะห่างบรรทัด line_spacing_pt เดียวกันทั้งหมด)
    _add_para(doc, to_arabic_num(f'ประกาศ{setting.announcement_school_name}'), fs_title1, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=line_spacing_pt)
    _add_para(doc, to_arabic_num(f'{setting.announcement_title} ค่าย 2 ปีการศึกษา {year}'),
              fs_title2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=line_spacing_pt)
    _add_para(doc, to_arabic_num(setting.announcement_center_name), fs_title3, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=line_spacing_pt)

    # เส้นคั่นทึบยาว 5 ซม. จัดกึ่งกลางหน้ากระดาษ
    div_tbl = doc.add_table(rows=1, cols=1)
    div_tbl.style = 'Table Grid'
    div_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    div_cell = div_tbl.rows[0].cells[0]
    div_cell.width = Cm(5.0)
    div_tcPr = div_cell._tc.get_or_add_tcPr()
    div_borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'right', 'insideH', 'insideV'):
        edge_el = OxmlElement(f'w:{edge}')
        edge_el.set(qn('w:val'), 'none')
        edge_el.set(qn('w:sz'), '0')
        edge_el.set(qn('w:color'), 'FFFFFF')
        div_borders.append(edge_el)
    bot_el = OxmlElement('w:bottom')
    bot_el.set(qn('w:val'), 'single')
    bot_el.set(qn('w:sz'), '6')
    bot_el.set(qn('w:color'), '000000')
    div_borders.append(bot_el)
    div_tcPr.append(div_borders)
    if div_cell.paragraphs:
        dp = div_cell.paragraphs[0]
        dp.paragraph_format.space_before = Pt(0)
        dp.paragraph_format.space_after = Pt(0)
        dp.paragraph_format.line_spacing = Pt(line_spacing_pt * 0.5)

    # ย่อหน้าที่ 1 (ชิดซ้ายปกติ พร้อมย่อหน้า indent_cm ไม่ยืดตัวอักษร)
    raw_p1 = format_announcement_body_p1(setting)
    for para_p1 in raw_p1.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        if para_p1.strip():
            _add_para(doc, para_p1.strip(), font_body, align=WD_ALIGN_PARAGRAPH.LEFT,
                      indent_cm=indent_cm, space_before=0, space_after=para_spacing_pt, line_spacing=line_spacing_pt)

    # ย่อหน้าที่ 2 (ชิดซ้ายปกติ พร้อมย่อหน้า indent_cm; ย่อหน้าสุดท้ายเว้น 1 1/2 บรรทัดก่อนวันที่)
    raw_p2 = format_announcement_body_p2(setting, passed_count)
    p2_paras = [p.strip() for p in raw_p2.replace('\r\n', '\n').replace('\r', '\n').split('\n') if p.strip()]
    for idx_p2, para_p2 in enumerate(p2_paras):
        after_pt = (0.5 * line_spacing_pt) if (idx_p2 == len(p2_paras) - 1) else para_spacing_pt
        _add_para(doc, para_p2, font_body, align=WD_ALIGN_PARAGRAPH.LEFT,
                  indent_cm=indent_cm, space_before=0, space_after=after_pt, line_spacing=line_spacing_pt)

    # วันที่ในประกาศ (ขอบซ้ายตรงกับคำว่า "ประกาศ..." ในหัวข้อบรรทัดที่ 1)
    date_text = format_announcement_date_text(setting)
    _add_para(doc, date_text, font_body, align=WD_ALIGN_PARAGRAPH.LEFT,
              indent_cm=metrics['date_indent_mm'] / 10.0, space_before=0, space_after=0, line_spacing=line_spacing_pt)

    # เว้น 2 บรรทัดเต็มสำหรับลงลายมือชื่อ (รวมเป็น 3x ระยะบรรทัดจากบรรทัดวันที่)
    _add_para(doc, '', font_body, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=0, space_after=0, line_spacing=line_spacing_pt)
    _add_para(doc, '', font_body, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=0, space_after=0, line_spacing=line_spacing_pt)

    # ลายเซ็นผู้อำนวยการ (ตาราง 2 คอลัมน์จัดวางตามพิกัด sig_left_from_margin_mm และ sig_box_w_mm ตรงกับ PDF)
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    sig_table.autofit = False
    for cell in sig_table.rows[0].cells:
        _set_cell_border_all_none(cell)
    sig_table.rows[0].cells[0].width = Cm(max(0.5, metrics['sig_left_from_margin_mm'] / 10.0))
    sig_table.rows[0].cells[1].width = Cm(metrics['sig_box_w_mm'] / 10.0)

    sig_cell = sig_table.rows[0].cells[1]
    _add_cell_text(sig_cell,
                   to_arabic_num(f'({setting.announcement_director_name})'),
                   fs_dir_name, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    if sig_cell.paragraphs:
        sig_cell.paragraphs[0].paragraph_format.line_spacing = Pt(line_spacing_pt)
    for line in (setting.announcement_director_position or '').splitlines():
        line = to_arabic_num(line.strip())
        if line:
            p = sig_cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(line_spacing_pt)
            r = p.add_run(line)
            _set_thai_font(r, fs_dir_pos)

    # ============ แผ่นที่ 2: รายชื่อแนบท้ายประกาศ ============
    doc.add_page_break()

    # 4 โลโก้
    logo_fields = [
        getattr(setting, 'logo_school', None),
        getattr(setting, 'logo_obec', None),
        getattr(setting, 'logo_posn', None),
        getattr(setting, 'logo_center', None),
    ]
    logo_paths = [_load_image_path(f) for f in logo_fields]
    logo_paths = [p for p in logo_paths if p]

    if logo_paths:
        logo_tbl = doc.add_table(rows=1, cols=len(logo_paths))
        logo_tbl.style = 'Table Grid'
        logo_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        logo_w_cm = min(3.5, 14.0 / len(logo_paths))
        for ci, lpath in enumerate(logo_paths):
            cell = logo_tbl.rows[0].cells[ci]
            _set_cell_border_all_none(cell)
            cell.width = Cm(logo_w_cm)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                r = p.add_run()
                r.add_picture(lpath, height=Cm(1.8))
            except Exception:
                pass

    # หัวรายชื่อ
    list_title = format_announcement_list_title(setting, year)
    _add_para(doc, list_title, fs_list1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=4, space_after=2)
    _add_para(doc, to_arabic_num(setting.announcement_center_name), fs_list2, bold=False,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4)

    # ตารางรายชื่อ
    # คอลัมน์: ลำดับ | รหัสค่าย2 | รหัสค่าย1/สิทธิ์ | ชื่อ-นามสกุล | โรงเรียน | ระดับ
    tbl = doc.add_table(rows=1 + len(rows), cols=6)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    col_widths_cm = [
        float(getattr(setting, 'announcement_col_w_no', 14.0) or 14.0) / 10.0,
        float(getattr(setting, 'announcement_col_w_camp2', 24.0) or 24.0) / 10.0,
        float(getattr(setting, 'announcement_col_w_rights', 24.0) or 24.0) / 10.0,
        float(getattr(setting, 'announcement_col_w_name', 50.0) or 50.0) / 10.0,
        float(getattr(setting, 'announcement_col_w_school', 54.0) or 54.0) / 10.0,
        float(getattr(setting, 'announcement_col_w_grade', 14.0) or 14.0) / 10.0,
    ]
    for ci, cw in enumerate(col_widths_cm):
        for row in tbl.rows:
            row.cells[ci].width = Cm(cw)

    # หัวตาราง (บรรทัดเดียวตามคำขอ)
    head_texts = [
        'ลำดับ',
        'เลขประจำตัวค่าย 2',
        'เลขประจำตัวค่าย 1',
        'ชื่อ นามสกุล',
        'โรงเรียน',
        'ระดับ',
    ]
    head_row = tbl.rows[0]
    trPr0 = head_row._tr.get_or_add_trPr()
    trPr0.append(OxmlElement('w:cantSplit'))
    for ci, hd in enumerate(head_texts):
        cell = head_row.cells[ci]
        _set_cell_bg(cell, 'F0F4F8')
        _add_cell_text(cell, hd, font_table + 0.5, bold=True)

    # แถวข้อมูล (ปรับขนาดย่อลงอัตโนมัติเพื่อให้อยู่ในบรรทัดเดียวเสมอ โดยเฉพาะโรงเรียน)
    for idx, r in enumerate(rows, 1):
        row = tbl.rows[idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:cantSplit'))

        if idx % 2 == 0:
            for ci in range(6):
                _set_cell_bg(row.cells[ci], 'FAFAFA')

        school_str = r['user'].school_name or ''
        school_sz = font_table
        if len(school_str) > 42:
            school_sz = max(7.0, font_table - 2.5)
        elif len(school_str) > 26:
            school_sz = max(8.0, font_table - 1.5)

        name_str = r['user'].get_full_name()
        name_sz = font_table
        if len(name_str) > 36:
            name_sz = max(8.0, font_table - 1.5)

        vals = [
            (str(idx), font_table, False, WD_ALIGN_PARAGRAPH.CENTER),
            (str(r.get('camp2_code', '')), font_table, True, WD_ALIGN_PARAGRAPH.CENTER),
            (str(r.get('rights_code', '')), font_table, False, WD_ALIGN_PARAGRAPH.CENTER),
            (name_str, name_sz, True, WD_ALIGN_PARAGRAPH.LEFT),
            (school_str, school_sz, False, WD_ALIGN_PARAGRAPH.LEFT),
            (str(r.get('grade', '')), font_table, False, WD_ALIGN_PARAGRAPH.CENTER),
        ]
        for ci, (val, sz, bld, align) in enumerate(vals):
            _add_cell_text(row.cells[ci], val, sz, bold=bld, align=align)

    # หมายเหตุ
    if getattr(setting, 'announcement_note', ''):
        _add_para(doc, f'หมายเหตุ: {to_arabic_num(setting.announcement_note)}', 12,
                  align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=0)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
