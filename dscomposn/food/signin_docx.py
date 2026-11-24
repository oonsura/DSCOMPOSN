"""
สร้างไฟล์ Word (.docx) แบบลงชื่อเข้าอบรม / วันสอบท้ายค่าย
- รูปแบบตรงกับ PDF (A4 แนวนอน, ตาราง 2 ฝั่ง ซ้าย-ขวา)
- รองรับฟอนต์ไทย TH Sarabun New (สามารถปรับแต่งได้)
- 1 sheet = 1 หน้าในไฟล์ Word
"""
import io
import os

from django.conf import settings

try:
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import docx.oxml
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ฟอนต์ที่ต้องการ (ผู้ใช้สามารถแก้ไขได้)
THAI_FONT = 'TH Sarabun New'
THAI_FONT_FALLBACK = 'Cordia New'


def _set_cell_bg(cell, hex_color):
    """ตั้งสีพื้นหลัง cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)


def _set_cell_border(cell, **kwargs):
    """ตั้งเส้นขอบ cell (top, bottom, left, right, insideH, insideV)"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, props in kwargs.items():
        edge_el = OxmlElement(f'w:{edge}')
        for key, val in props.items():
            edge_el.set(qn(f'w:{key}'), val)
        tcBorders.append(edge_el)
    tcPr.append(tcBorders)


def _set_font(run, size_pt, bold=False):
    """ตั้งฟอนต์ไทย + ขนาด + ตัวหนา"""
    run.font.name = THAI_FONT
    run._element.rPr.rFonts.set(qn('w:eastAsia'), THAI_FONT)
    run._element.rPr.rFonts.set(qn('w:cs'), THAI_FONT)
    run.font.size = Pt(size_pt)
    run.font.bold = bold


def _para_font(para, text, size_pt, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    """เพิ่มข้อความใน paragraph พร้อมตั้งฟอนต์"""
    para.alignment = align
    run = para.add_run(text)
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
    return run


def _add_cell_text(cell, text, size_pt, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    """เพิ่มข้อความใน table cell"""
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ''
        para.text = ''
    if cell.paragraphs:
        para = cell.paragraphs[0]
    else:
        para = cell.add_paragraph()
    para.alignment = align
    run = para.add_run(str(text) if text is not None else '')
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
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _set_col_width(table, col_idx, width_cm):
    """ตั้งความกว้างคอลัมน์"""
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)


def build_signin_docx(sheets, setting, camp=1):
    """
    สร้าง Word (.docx) แบบลงชื่อ
    คืนค่า bytes ของไฟล์ docx
    """
    if not HAS_DOCX:
        return b''

    doc = Document()

    # ตั้งขนาดหน้ากระดาษ A4 แนวนอน
    section = doc.sections[0]
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(0.7)
    section.bottom_margin = Cm(0.7)
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.8)
    section.orientation = 1  # WD_ORIENT.LANDSCAPE

    # ความกว้างคอลัมน์ 2 ฝั่ง (ซ้าย + ขวา) รวม 28.1 cm
    # ลำดับ | รหัส | ชื่อ-นามสกุล | ชั้น | โรงเรียน | ลงชื่อ
    col_widths = [1.0, 2.1, 5.0, 1.3, 4.3, 0.95]  # half (cm)
    all_col_w = col_widths + col_widths  # 12 columns

    # โลโก้
    logo_school_path = None
    logo_posn_path = None
    try:
        if setting.logo_school:
            logo_school_path = setting.logo_school.path
    except Exception:
        pass
    try:
        logo_posn = getattr(setting, 'logo_posn', None) or getattr(setting, 'logo_center', None)
        if logo_posn:
            logo_posn_path = logo_posn.path
    except Exception:
        pass

    total_sheets = len(sheets) if sheets else 0
    for s_idx, sheet in enumerate(sheets):
        if s_idx > 0:
            # เพิ่มตัวแบ่งหน้า
            doc.add_page_break()

        # ============ ส่วนหัวเอกสาร ============
        # ใช้ตาราง 1 แถว 3 คอลัมน์: โลโก้ซ้าย | ข้อความ | โลโก้ขวา
        header_table = doc.add_table(rows=1, cols=3)
        header_table.style = 'Table Grid'
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # ซ่อนเส้นขอบหัวตาราง
        for cell in header_table.rows[0].cells:
            _set_cell_border(cell,
                top={'val': 'none', 'sz': '0', 'color': 'FFFFFF'},
                bottom={'val': 'none', 'sz': '0', 'color': 'FFFFFF'},
                left={'val': 'none', 'sz': '0', 'color': 'FFFFFF'},
                right={'val': 'none', 'sz': '0', 'color': 'FFFFFF'},
            )

        logo_cell_w = Cm(2.0)
        text_cell_w = Cm(25.1)
        logo_left_cell = header_table.rows[0].cells[0]
        logo_left_cell.width = logo_cell_w
        title_cell = header_table.rows[0].cells[1]
        title_cell.width = text_cell_w
        logo_right_cell = header_table.rows[0].cells[2]
        logo_right_cell.width = logo_cell_w

        # โลโก้ซ้าย
        try:
            if logo_school_path and os.path.exists(logo_school_path):
                p = logo_left_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run()
                r.add_picture(logo_school_path, width=Cm(1.8))
        except Exception:
            logo_left_cell.paragraphs[0].text = ''

        # ข้อความหัว
        p1 = title_cell.paragraphs[0]
        _para_font(p1, 'ศูนย์ส่งเสริมโอลิมปิกวิชาการและพัฒนามาตรฐานวิทยาศาสตร์ศึกษา', 14.5, bold=True)
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(1)

        p2 = title_cell.add_paragraph()
        _para_font(p2, 'ในพระอุปถัมภ์สมเด็จพระเจ้าพี่นางเธอ เจ้าฟ้ากัลยาณิวัฒนา กรมหลวงนราธิวาสราชนครินทร์ (สอวน.)', 12.5, bold=True)
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(1)

        camp_line = f'สาขาวิชาคอมพิวเตอร์ ค่าย {camp} ศูนย์โรงเรียนเทพศิรินทร์   {sheet.get("date_thai_header", "")}'
        p3 = title_cell.add_paragraph()
        _para_font(p3, camp_line, 13.5, bold=True)
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(1)

        title_l4 = sheet.get('title_line4') or (
            f'{"แบบลงชื่อเข้าสอบ" if sheet.get("day_mode") == "exam" else "แบบลงชื่อเข้าอบรม"} '
            f'({sheet.get("period_label", "")})'
        )
        if sheet.get('total_pages', 1) > 1:
            title_l4 += f' — หน้า {sheet["page_no"]}/{sheet["total_pages"]}'

        p4 = title_cell.add_paragraph()
        r4 = _para_font(p4, title_l4, 13.5, bold=True)
        r4.underline = True
        p4.paragraph_format.space_before = Pt(0)
        p4.paragraph_format.space_after = Pt(2)

        # โลโก้ขวา
        try:
            if logo_posn_path and os.path.exists(logo_posn_path):
                p = logo_right_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run()
                r.add_picture(logo_posn_path, width=Cm(1.8))
        except Exception:
            logo_right_cell.paragraphs[0].text = ''

        # ============ ตาราง 12 คอลัมน์ ============
        paired = sheet.get('paired_rows') or []
        n_rows = len(paired) if paired else 1

        tbl = doc.add_table(rows=1 + n_rows, cols=12)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        # ตั้งความกว้างคอลัมน์
        for col_idx, cw in enumerate(all_col_w):
            for row in tbl.rows:
                row.cells[col_idx].width = Cm(cw)

        # หัวตาราง
        head_labels = [
            ('ลำดับ', ''),
            ('เลขประจำตัว', 'ค่าย'),
            ('ชื่อ-นามสกุล', ''),
            ('ระดับ', 'ชั้น'),
            ('โรงเรียน', ''),
            ('ลงชื่อ', ''),
        ]
        head_row = tbl.rows[0]
        for side in range(2):
            for col_i, (h1, h2) in enumerate(head_labels):
                ci = side * 6 + col_i
                cell = head_row.cells[ci]
                _set_cell_bg(cell, 'EEF1F4')
                label = f'{h1}\n{h2}' if h2 else h1
                _add_cell_text(cell, label, 11.5, bold=True)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # แถวข้อมูล
        for r_idx, pair in enumerate(paired):
            row = tbl.rows[r_idx + 1]
            if r_idx % 2 == 1:
                for ci in range(12):
                    _set_cell_bg(row.cells[ci], 'FAFBFC')

            for side, item_key in ((0, 'left'), (1, 'right')):
                item = pair.get(item_key)
                vals = [
                    (str(item.get('no', '')) if item else '', 11.5, False, WD_ALIGN_PARAGRAPH.CENTER),
                    (str(item.get('code', '')) if item else '', 11.0, True, WD_ALIGN_PARAGRAPH.CENTER),
                    (str(item.get('full_name', '')) if item else '', 11.0, False, WD_ALIGN_PARAGRAPH.LEFT),
                    (str(item.get('grade', '')) if item else '', 11.0, False, WD_ALIGN_PARAGRAPH.CENTER),
                    (str(item.get('school', '')) if item else '', 10.0, False, WD_ALIGN_PARAGRAPH.LEFT),
                    ('', 11.0, False, WD_ALIGN_PARAGRAPH.CENTER),  # ลงชื่อ
                ]
                for col_i, (val, sz, bld, align) in enumerate(vals):
                    ci = side * 6 + col_i
                    _add_cell_text(row.cells[ci], val, sz, bold=bld, align=align)

        # หากไม่มีข้อมูล
        if not paired:
            row = tbl.rows[1]
            cell = row.cells[0]
            cell.merge(row.cells[11])
            _add_cell_text(cell, 'ไม่พบรายชื่อนักเรียนในกลุ่มนี้', 12, bold=False)

        # ท้ายหน้า
        footer_p = doc.add_paragraph()
        footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_fl = footer_p.add_run(
            f'กลุ่มรายชื่อ: {sheet.get("group_label", "นักเรียนปัจจุบัน")} · จำนวน {sheet.get("count", 0)} คน'
            f'{"":>50}แผ่นที่ {s_idx + 1} / {total_sheets}'
        )
        run_fl.font.name = THAI_FONT
        run_fl.font.size = Pt(9)
        try:
            rPr = run_fl._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), THAI_FONT)
            rFonts.set(qn('w:hAnsi'), THAI_FONT)
            rFonts.set(qn('w:eastAsia'), THAI_FONT)
            rFonts.set(qn('w:cs'), THAI_FONT)
            rPr.insert(0, rFonts)
        except Exception:
            pass

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
