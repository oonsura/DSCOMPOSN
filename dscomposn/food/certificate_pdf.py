"""
สร้างเกียรติบัตรออนไลน์ (Online Certificate Generator)
สำหรับนักเรียน, วิทยากร, และคณะกรรมการ โครงการ สอวน. วิชาคอมพิวเตอร์

กฎการรับเกียรติบัตร (นักเรียน):
1. นักเรียนปัจจุบัน: is_current_camp_student(user, year) == True
2. ขาดอบรมไม่เกินที่กำหนด (absence_count(user, year) <= setting.max_absences) (นับเฉพาะวันอบรม)
3. เข้าสอบครบทุกครั้ง (เข้าสอบทุกช่วงเวลาสอบใน TrainingDate event_type='exam')
4. นักเรียนใช้สิทธิ์ ไม่ต้องพิจารณาส่วนนี้ (ไม่ได้รับเกียรติบัตร)
5. กฎสำคัญ: **ทุกคนที่ได้รับเกียรติบัตร หรือไม่ได้รับเกียรติบัตร จะได้รับสิทธิ์สอบท้ายค่าย**

รูปแบบเกียรติบัตร (ดึงการตั้งค่าจาก CertificateSetting ในฐานข้อมูล):
- รองรับรูปพื้นหลังเต็มหน้า A4 แนวนอน (กรณีไม่มีรูปพื้นหลังจะวาดกรอบมาตรฐานเป็น Fallback)
- วางเฉพาะข้อความ 4 ส่วนด้วยฟอนต์ TH Sarabun New ตามตำแหน่งแกน Y, ขนาดตัวอักษร, ตัวหนา และสี:
  1. ชื่อ นามสกุล
  2. โรงเรียน จังหวัด
  3. ข้อความรายละเอียด (เช่น ได้เข้าร่วมเป็น ... ค่าย ๑ สอวน. วิชาคอมพิวเตอร์) รองรับหลายบรรทัด
  4. ให้ไว้ ณ วันที่ (ใช้ร่วมกันทุกกลุ่ม ตั้งค่าที่เดียว)
"""
import csv
import io
import os
import re
import zipfile
from datetime import date

from django.urls import reverse
from openpyxl import load_workbook
from PIL import Image
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

from composn.models import SystemSetting, User
from composn.templatetags.composn_tags import to_thai_num
from food.models import (
    Attendance, CertificateRecipient, CertificateRecord, CertificateSetting, TrainingDate,
    absence_count, current_year, is_current_camp_student,
)
from food.signin_pdf import (
    _draw_logo_box, _ensure_thai_fonts, _load_logo_reader, shape_thai,
)


def check_certificate_eligibility(user, year=None):
    """
    ตรวจสอบสิทธิ์การได้รับเกียรติบัตรออนไลน์ของนักเรียนในปีการศึกษาที่ระบุ
    คืนค่า dict รายละเอียดผลการตรวจสอบ
    """
    setting = SystemSetting.get()
    year = int(year) if year else current_year()

    # 1. สิทธิ์สอบท้ายค่าย: ทุกคนได้รับสิทธิ์สอบท้ายค่ายเสมอ
    final_exam_eligible = True

    # 2. ตรวจว่าเป็นนักเรียนปัจจุบันหรือไม่ (นักเรียนใช้สิทธิ์ไม่พิจารณาส่วนนี้)
    is_current = (user.role == User.ROLE_MEMBER) and is_current_camp_student(user, year)
    if not is_current:
        return {
            'is_eligible': False,
            'is_current': False,
            'absences': 0,
            'max_absences': setting.max_absences,
            'absence_passed': False,
            'total_exam_slots': 0,
            'attended_exam_slots': 0,
            'exams_passed': False,
            'final_exam_eligible': final_exam_eligible,
            'reason': 'เป็นนักเรียนใช้สิทธิ์ (เกียรติบัตรสำหรับนักเรียนปัจจุบันเท่านั้น)',
        }

    # 3. ตรวจจำนวนครั้งที่ขาดอบรม (นับเฉพาะวันอบรม)
    absences = absence_count(user, year)
    max_abs = setting.max_absences
    absence_passed = (absences <= max_abs)

    # 4. ตรวจการเข้าสอบครบทุกครั้ง
    exam_tds = list(TrainingDate.objects.filter(academic_year=year, event_type=TrainingDate.TYPE_EXAM).order_by('date'))
    total_exam_slots = 0
    attended_exam_slots = 0
    missing_exam_details = []

    for td in exam_tds:
        if getattr(td, 'exam_period', TrainingDate.EXAM_BOTH) == TrainingDate.EXAM_MORNING:
            periods = [(Attendance.PERIOD_MORNING, 'ภาคเช้า')]
        elif getattr(td, 'exam_period', TrainingDate.EXAM_BOTH) == TrainingDate.EXAM_AFTERNOON:
            periods = [(Attendance.PERIOD_AFTERNOON, 'ภาคบ่าย')]
        else:
            periods = [(Attendance.PERIOD_MORNING, 'ภาคเช้า'), (Attendance.PERIOD_AFTERNOON, 'ภาคบ่าย')]

        for p_code, p_lbl in periods:
            total_exam_slots += 1
            has_attended = Attendance.objects.filter(
                user=user, date=td.date, session_period=p_code, status=Attendance.STATUS_PRESENT
            ).exists()
            if has_attended:
                attended_exam_slots += 1
            else:
                missing_exam_details.append(f'{td.date.strftime("%d/%m/%Y")} ({p_lbl})')

    exams_passed = (len(missing_exam_details) == 0)

    # สรุปผลความมีสิทธิ์
    is_eligible = absence_passed and exams_passed
    reasons = []
    if not absence_passed:
        reasons.append(f'ขาดอบรม {absences} ภาคเวลา (เกินเกณฑ์สูงสุด {max_abs} ภาคเวลา)')
    if not exams_passed:
        reasons.append(f'ขาดสอบ {len(missing_exam_details)} ครั้ง: ' + ', '.join(missing_exam_details[:3]))

    reason_str = 'ผ่านเกณฑ์ได้รับเกียรติบัตร (ขาดอบรมไม่เกินเกณฑ์และเข้าสอบครบ)' if is_eligible else ' · '.join(reasons)

    return {
        'is_eligible': is_eligible,
        'is_current': True,
        'absences': absences,
        'max_absences': max_abs,
        'absence_passed': absence_passed,
        'total_exam_slots': total_exam_slots,
        'attended_exam_slots': attended_exam_slots,
        'missing_exams': missing_exam_details,
        'exams_passed': exams_passed,
        'final_exam_eligible': final_exam_eligible,
        'reason': reason_str,
    }


def get_certificate_data_for_students(year=None):
    """ดึงข้อมูลนักเรียนทุกคนในปีการศึกษา พร้อมสถานะเกียรติบัตรและสิทธิ์สอบท้ายค่าย"""
    year = int(year) if year else current_year()
    from food.models import score_members_for_year
    members, _ = score_members_for_year(year)
    results = []
    for m in members:
        elig = check_certificate_eligibility(m, year)
        results.append({
            'user': m,
            'eligibility': elig,
            'is_eligible': elig['is_eligible'],
            'is_current': elig['is_current'],
            'final_exam_eligible': elig['final_exam_eligible'],
        })
    return results


# =============================================================================
# PDF Rendering Helpers
# =============================================================================
_HEX_RE = re.compile(r'^#[0-9a-fA-F]{6}$')


def _safe_hex_color(value, default='#000000'):
    """แปลงโค้ดสี Hex ให้เป็น ReportLab Color อย่างปลอดภัย"""
    val = str(value or '').strip()
    if not _HEX_RE.match(val):
        val = default
    return colors.HexColor(val)


def _load_background_reader(cert_setting):
    """โหลดรูปพื้นหลังเกียรติบัตรจาก ImageField เป็น ReportLab ImageReader (ย่อขนาดหากใหญ่เกิน A4 212dpi เพื่อคุมขนาดไฟล์ PDF)"""
    if not cert_setting or not cert_setting.background_image:
        return None
    try:
        path = cert_setting.background_image.path
    except Exception:
        return None
    if not path or not os.path.exists(path):
        return None
    try:
        img = Image.open(path)
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            rgba = img.convert('RGBA')
            bg = Image.new('RGBA', rgba.size, (255, 255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[3])
            img = bg.convert('RGB')
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        # ย่อรูปที่ความละเอียดเกินความจำเป็น (2480 px สำหรับ A4 แนวนอน) เพื่อความเร็วในการสร้าง Batch PDF
        max_w = 2480
        if img.width > max_w:
            new_h = max(1, int(round(img.height * (max_w / float(img.width)))))
            img = img.resize((max_w, new_h), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=92)
        buf.seek(0)
        return ImageReader(buf)
    except Exception:
        return None


def _draw_centered_fitted_line(c, text, center_x, y_top_mm, font_name, font_size, color_hex, max_width, page_height):
    """วาดข้อความ 1 บรรทัดกึ่งกลางแนวนอน ณ ตำแหน่งแกน Y (mm จากขอบบน) โดยย่อขนาดอัตโนมัติหากยาวเกินขอบ"""
    raw = str(text or '').strip()
    if not raw:
        return
    shaped = shape_thai(raw)
    if not shaped:
        return

    actual_size = max(10.0, float(font_size or 18.0))
    while actual_size > 10.0 and c.stringWidth(shaped, font_name, actual_size) > max_width:
        actual_size = max(10.0, actual_size - 0.5)

    y_center_pt = page_height - (float(y_top_mm) * mm)
    # capHeight ของ THSarabunNew ~ 476/1000 -> จัดให้จุดกึ่งกลางตัวอักษรตรงกับ y_center_pt
    baseline_y = y_center_pt - (0.24 * actual_size)

    c.setFont(font_name, actual_size)
    c.setFillColor(_safe_hex_color(color_hex))
    c.drawCentredString(center_x, baseline_y, shaped)


def _draw_fallback_chrome(c, entry, year, setting, font_reg, font_bold, width, height):
    """วาดกรอบและโลโก้มาตรฐานเมื่อยังไม่ได้อัปโหลดรูปพื้นหลัง (จัดวางเหนือ y=78mm และใต้ y=172mm เพื่อไม่ซ้อนข้อความ 4 ส่วน)"""
    margin = 12 * mm
    c.saveState()

    c.setStrokeColor(colors.HexColor('#b8860b'))
    c.setLineWidth(3.0)
    c.rect(margin, margin, width - 2 * margin, height - 2 * margin)

    inner_margin = margin + 3.5 * mm
    c.setStrokeColor(colors.HexColor('#0f4c5c'))
    c.setLineWidth(1.2)
    c.rect(inner_margin, inner_margin, width - 2 * inner_margin, height - 2 * inner_margin)

    corner_size = 8 * mm
    c.setStrokeColor(colors.HexColor('#b8860b'))
    c.setLineWidth(1.5)
    for cx, cy, dx, dy in [
        (inner_margin + 2 * mm, inner_margin + 2 * mm, 1, 1),
        (width - inner_margin - 2 * mm, inner_margin + 2 * mm, -1, 1),
        (inner_margin + 2 * mm, height - inner_margin - 2 * mm, 1, -1),
        (width - inner_margin - 2 * mm, height - inner_margin - 2 * mm, -1, -1),
    ]:
        c.line(cx, cy, cx + dx * corner_size, cy)
        c.line(cx, cy, cx, cy + dy * corner_size)
    c.restoreState()

    center_x = width / 2.0

    # โลโก้ 4 รูปด้านบน
    logo_w, logo_h = 20 * mm, 20 * mm
    logo_y = height - margin - 24 * mm
    readers = [r for r in (
        _load_logo_reader(field)
        for field in (setting.logo_school, setting.logo_obec, setting.logo_posn, setting.logo_center)
    ) if r]
    if readers:
        spacing = 34 * mm
        start_x = center_x - ((len(readers) - 1) * spacing) / 2.0 - (logo_w / 2.0)
        for idx, rdr in enumerate(readers):
            _draw_logo_box(c, rdr, start_x + idx * spacing, logo_y, logo_w, logo_h)

    # ลายเซ็นด้านล่าง (ใต้ y=172mm)
    sig_y = height - 180 * mm
    c.setFont(font_bold, 14)
    c.setFillColor(colors.HexColor('#111111'))
    c.drawCentredString(center_x, sig_y, shape_thai(f'({setting.announcement_director_name})'))

    pos_lines = [line.strip() for line in (setting.announcement_director_position or '').splitlines() if line.strip()]
    c.setFont(font_reg, 11.5)
    c.setFillColor(colors.HexColor('#444444'))
    for idx, pl in enumerate(pos_lines[:2]):
        c.drawCentredString(center_x, sig_y - 4.5 * mm - (idx * 4.0 * mm), shape_thai(pl))


def _draw_qr_and_ref_code(c, ref_code, cert_setting, verify_base_url, font_reg, width, height):
    """วาด QR Code และเลขที่อ้างอิงเกียรติบัตรสำหรับสแกนตรวจสอบว่าเป็นเอกสารจากศูนย์อบรมจริง"""
    ref_code = (ref_code or '').strip()
    if not getattr(cert_setting, 'show_qr_verify', True) or not ref_code:
        return

    verify_path = reverse('food:certificate_verify', args=[ref_code])
    verify_url = (verify_base_url.rstrip('/') + verify_path) if verify_base_url else verify_path

    qr_size_mm = max(12.0, min(50.0, float(getattr(cert_setting, 'qr_size', 24.0) or 24.0)))
    qr_size_pt = qr_size_mm * mm
    x_pt = max(6 * mm, min(width - qr_size_pt - 6 * mm, float(getattr(cert_setting, 'qr_x', 252.0) or 252.0) * mm))
    y_top_pt = height - (max(6.0, min(200.0, float(getattr(cert_setting, 'qr_y', 164.0) or 164.0))) * mm)
    y_bottom_pt = max(11 * mm, y_top_pt - qr_size_pt)

    pad = 1.5 * mm
    c.saveState()
    c.setFillColor(colors.white)
    c.setStrokeColor(colors.HexColor('#cbd5e1'))
    c.setLineWidth(0.5)
    c.rect(x_pt - pad, y_bottom_pt - pad, qr_size_pt + 2 * pad, qr_size_pt + 2 * pad, fill=1, stroke=1)
    c.restoreState()

    qr_widget = qr.QrCodeWidget(verify_url)
    bounds = qr_widget.getBounds()
    w_qr = max(1.0, bounds[2] - bounds[0])
    h_qr = max(1.0, bounds[3] - bounds[1])
    drawing = Drawing(qr_size_pt, qr_size_pt, transform=[qr_size_pt / w_qr, 0, 0, qr_size_pt / h_qr, 0, 0])
    drawing.add(qr_widget)
    renderPDF.draw(drawing, c, x_pt, y_bottom_pt)

    ref_fs = max(6.5, min(20.0, float(getattr(cert_setting, 'ref_code_font_size', 9.5) or 9.5)))
    ref_color = _safe_hex_color(getattr(cert_setting, 'ref_code_color', '#333333'), '#333333')
    label_cx = max(24 * mm, min(width - 24 * mm, x_pt + (qr_size_pt / 2.0)))

    c.setFillColor(ref_color)
    c.setFont(font_reg, max(6.5, ref_fs - 1.0))
    c.drawCentredString(label_cx, y_bottom_pt - 4.0 * mm, shape_thai('สแกนตรวจสอบเกียรติบัตร'))
    c.setFont(font_reg, ref_fs)
    c.drawCentredString(label_cx, y_bottom_pt - 7.8 * mm, shape_thai(f'เลขที่อ้างอิง: {ref_code}'))


def draw_certificate_entry(c, entry, year, setting, cert_setting, font_reg, font_bold, width, height,
                           bg_reader=None, verify_base_url=''):
    """วาดหน้าเกียรติบัตร 1 หน้า: พื้นหลัง + ข้อความส่วนหัว (1–2) + ข้อความหลัก (1–4) + QR & เลขที่อ้างอิง"""
    if bg_reader is not None:
        c.drawImage(bg_reader, 0, 0, width=width, height=height, preserveAspectRatio=False)
    else:
        _draw_fallback_chrome(c, entry, year, setting, font_reg, font_bold, width, height)

    center_x = width / 2.0
    max_width = width - 24 * mm

    # ข้อความส่วนหัว 1 (รองรับหลายบรรทัด): มูลนิธิ / ศูนย์อบรม
    h1_raw = cert_setting.render_text(getattr(cert_setting, 'header1_text', ''), year)
    h1_lines = [ln.strip() for ln in h1_raw.splitlines() if ln.strip()]
    if h1_lines:
        fh1 = font_bold if getattr(cert_setting, 'header1_bold', True) else font_reg
        h1_y = float(getattr(cert_setting, 'header1_y', 45.0) or 45.0)
        h1_sp = float(getattr(cert_setting, 'header1_line_spacing', 10.0) or 10.0)
        h1_fs = getattr(cert_setting, 'header1_font_size', 22.0) or 22.0
        h1_color = getattr(cert_setting, 'header1_color', '#0f4c5c') or '#0f4c5c'
        for idx, ln in enumerate(h1_lines):
            _draw_centered_fitted_line(
                c, ln, center_x, h1_y + (idx * h1_sp),
                fh1, h1_fs, h1_color, max_width, height,
            )

    # ข้อความส่วนหัว 2: ขอมอบเกียรติบัตรฉบับนี้ให้ไว้เพื่อแสดงว่า
    h2_text = cert_setting.render_text(getattr(cert_setting, 'header2_text', ''), year).strip()
    if h2_text:
        fh2 = font_bold if getattr(cert_setting, 'header2_bold', True) else font_reg
        _draw_centered_fitted_line(
            c, h2_text, center_x, getattr(cert_setting, 'header2_y', 71.0) or 71.0,
            fh2, getattr(cert_setting, 'header2_font_size', 24.0) or 24.0,
            getattr(cert_setting, 'header2_color', '#b8860b') or '#b8860b', max_width, height,
        )

    # ข้อ 1: ชื่อ นามสกุล
    f1 = font_bold if cert_setting.line1_bold else font_reg
    _draw_centered_fitted_line(
        c, entry.get('full_name', ''), center_x, cert_setting.line1_y,
        f1, cert_setting.line1_font_size, cert_setting.line1_color, max_width, height,
    )

    # ข้อ 2: โรงเรียน จังหวัด
    f2 = font_bold if cert_setting.line2_bold else font_reg
    _draw_centered_fitted_line(
        c, entry.get('school_province', ''), center_x, cert_setting.line2_y,
        f2, cert_setting.line2_font_size, cert_setting.line2_color, max_width, height,
    )

    # ข้อ 3: รายละเอียด (รองรับหลายบรรทัด)
    desc_raw = cert_setting.render_text(entry.get('description', ''), year)
    desc_lines = [ln.strip() for ln in desc_raw.splitlines()]
    while desc_lines and not desc_lines[-1]:
        desc_lines.pop()
    f3 = font_bold if cert_setting.line3_bold else font_reg
    spacing_mm = float(cert_setting.line3_line_spacing or 11.0)
    for idx, ln in enumerate(desc_lines):
        if not ln:
            continue
        _draw_centered_fitted_line(
            c, ln, center_x, float(cert_setting.line3_y) + (idx * spacing_mm),
            f3, cert_setting.line3_font_size, cert_setting.line3_color, max_width, height,
        )

    # ข้อ 4: ให้ไว้ ณ วันที่ (ใช้ร่วมกันทุกกลุ่ม ตั้งค่าที่เดียว)
    date_line = cert_setting.render_text(cert_setting.line4_text, year)
    f4 = font_bold if cert_setting.line4_bold else font_reg
    _draw_centered_fitted_line(
        c, date_line, center_x, cert_setting.line4_y,
        f4, cert_setting.line4_font_size, cert_setting.line4_color, max_width, height,
    )

    # QR Code และเลขที่อ้างอิงเกียรติบัตร
    _draw_qr_and_ref_code(c, entry.get('ref_code', ''), cert_setting, verify_base_url, font_reg, width, height)


def build_entries_certificate_pdf(entries, year, setting=None, cert_setting=None, verify_base_url=''):
    """สร้างเอกสาร PDF เกียรติบัตรจากรายการ entries (ใช้ร่วมกันทุกกลุ่ม)"""
    font_reg, font_bold = _ensure_thai_fonts()
    setting = setting or SystemSetting.get()
    cert_setting = cert_setting or CertificateSetting.get()
    bg_reader = _load_background_reader(cert_setting)

    buf = io.BytesIO()
    width, height = landscape(A4)
    c = canvas.Canvas(buf, pagesize=landscape(A4))
    c.setTitle(f'เกียรติบัตร สอวน.คอมพิวเตอร์ ค่าย 1 ปีการศึกษา {year}')

    count = 0
    for entry in entries:
        draw_certificate_entry(
            c, entry, year, setting, cert_setting, font_reg, font_bold, width, height,
            bg_reader=bg_reader, verify_base_url=verify_base_url,
        )
        c.showPage()
        count += 1

    if count == 0:
        c.drawString(100, 100, 'No certificate')
        c.showPage()

    c.save()
    buf.seek(0)
    return buf.getvalue()


def build_student_cert_entry(user, year, cert_setting=None):
    """สร้างข้อมูลเกียรติบัตรสำหรับนักเรียน 1 คน พร้อมออกเลขที่อ้างอิง"""
    cert_setting = cert_setting or CertificateSetting.get()
    rec = CertificateRecord.issue_for_student(user, year, cert_setting)
    return {
        'full_name': user.get_full_name(),
        'school_province': cert_setting.format_student_school(user.school_name),
        'description': cert_setting.student_line3_text,
        'code': getattr(user, 'camp1_code', ''),
        'ref_code': rec.ref_code,
    }


def build_recipient_cert_entry(rec, cert_setting=None):
    """สร้างข้อมูลเกียรติบัตรสำหรับวิทยากร/คณะกรรมการ 1 คน พร้อมออกเลขที่อ้างอิง"""
    cert_setting = cert_setting or CertificateSetting.get()
    cert_rec = CertificateRecord.issue_for_recipient(rec, cert_setting)
    return {
        'full_name': rec.full_name,
        'school_province': rec.school_province,
        'description': rec.resolved_description(cert_setting),
        'code': '',
        'ref_code': cert_rec.ref_code,
    }


def build_certificate_pdf(students_data, year, setting=None, cert_setting=None, verify_base_url=''):
    """สร้างเอกสาร PDF เกียรติบัตรออนไลน์สำหรับนักเรียน พร้อมออกเลขที่อ้างอิงและ QR Code ตรวจสอบ"""
    cert_setting = cert_setting or CertificateSetting.get()
    entries = [
        build_student_cert_entry(item['user'] if isinstance(item, dict) else item, year, cert_setting)
        for item in students_data
    ]
    return build_entries_certificate_pdf(
        entries, year, setting=setting, cert_setting=cert_setting, verify_base_url=verify_base_url,
    )


def build_recipients_certificate_pdf(recipients, year, setting=None, cert_setting=None, verify_base_url=''):
    """สร้างเอกสาร PDF เกียรติบัตรสำหรับกลุ่มวิทยากรและคณะกรรมการ พร้อมออกเลขที่อ้างอิงและ QR Code ตรวจสอบ"""
    cert_setting = cert_setting or CertificateSetting.get()
    entries = [build_recipient_cert_entry(rec, cert_setting) for rec in recipients]
    return build_entries_certificate_pdf(
        entries, year, setting=setting, cert_setting=cert_setting, verify_base_url=verify_base_url,
    )


def _safe_zip_filename(text, fallback='certificate'):
    """ปรับชื่อไฟล์ PDF ภายใน ZIP ให้ปลอดภัยกับ Windows/macOS และยังคงอักษรไทยครบถ้วน"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]+', ' ', str(text or ''))
    cleaned = re.sub(r'\s+', '_', cleaned).strip(' ._')
    if len(cleaned) > 120:
        cleaned = cleaned[:120].rstrip(' ._')
    return cleaned or fallback


def build_certificates_zip(items, year, setting=None, cert_setting=None, verify_base_url=''):
    """
    สร้างไฟล์ ZIP รวมเกียรติบัตร PDF แยกเป็นรายบุคคล (1 ไฟล์ PDF ต่อ 1 คน)
    items: list of {'entry': dict, 'folder': str (ไม่บังคับ), 'filename': str}
    """
    font_reg, font_bold = _ensure_thai_fonts()
    setting = setting or SystemSetting.get()
    cert_setting = cert_setting or CertificateSetting.get()
    bg_reader = _load_background_reader(cert_setting)

    zip_buf = io.BytesIO()
    width, height = landscape(A4)
    used_paths = set()

    with zipfile.ZipFile(zip_buf, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, item in enumerate(items, 1):
            entry = item['entry']
            folder = _safe_zip_filename(item.get('folder') or '', fallback='') if item.get('folder') else ''
            raw_stem = str(item.get('filename') or f'certificate_{idx}')
            if raw_stem.lower().endswith('.pdf'):
                raw_stem = raw_stem[:-4]
            stem = _safe_zip_filename(raw_stem, fallback=f'certificate_{idx}')

            arc_base = f'{folder}/{stem}' if folder else stem
            arcname = f'{arc_base}.pdf'
            dup_i = 2
            while arcname.lower() in used_paths:
                arcname = f'{arc_base}_{dup_i}.pdf'
                dup_i += 1
            used_paths.add(arcname.lower())

            pdf_buf = io.BytesIO()
            c = canvas.Canvas(pdf_buf, pagesize=landscape(A4))
            c.setTitle(f"{entry.get('full_name', '')} - เกียรติบัตร {year}")
            draw_certificate_entry(
                c, entry, year, setting, cert_setting, font_reg, font_bold, width, height,
                bg_reader=bg_reader, verify_base_url=verify_base_url,
            )
            c.showPage()
            c.save()
            zf.writestr(arcname, pdf_buf.getvalue())

    zip_buf.seek(0)
    return zip_buf.getvalue()


# =============================================================================
# Excel / CSV Upload Parser (วิทยากร & คณะกรรมการ)
# =============================================================================
_RECIPIENT_HEADER_ALIASES = {
    'ประเภท': 'recipient_type',
    'กลุ่ม': 'recipient_type',
    'บทบาท': 'recipient_type',
    'type': 'recipient_type',
    'recipient_type': 'recipient_type',
    'ชื่อ-นามสกุล': 'full_name',
    'ชื่อ - นามสกุล': 'full_name',
    'ชื่อ นามสกุล': 'full_name',
    'ชื่อนามสกุล': 'full_name',
    'full_name': 'full_name',
    'name': 'full_name',
    'คำนำหน้า': 'prefix',
    'ชื่อ': 'first_name',
    'นามสกุล': 'last_name',
    'โรงเรียน จังหวัด': 'school_province',
    'โรงเรียน/จังหวัด': 'school_province',
    'โรงเรียน/หน่วยงาน': 'school_province',
    'โรงเรียน': 'school_province',
    'หน่วยงาน': 'school_province',
    'สังกัด': 'school_province',
    'school_province': 'school_province',
    'school': 'school_province',
    'จังหวัด': 'province',
    'province': 'province',
    'รายละเอียด': 'description',
    'รายละเอียด (ข้อ 3)': 'description',
    'รายละเอียด (ข้อ 3 - เว้นว่างใช้ค่าเริ่มต้น)': 'description',
    'ข้อความ': 'description',
    'ข้อความที่ 3': 'description',
    'ได้เข้าร่วมเป็น': 'description',
    'หน้าที่': 'description',
    'ตำแหน่ง': 'description',
    'description': 'description',
}


def _normalize_recipient_type(raw_val, default_type):
    val = str(raw_val or '').strip().lower()
    if not val:
        return default_type
    if 'วิทยากร' in val or 'lecturer' in val:
        return CertificateRecipient.TYPE_LECTURER
    if 'กรรมการ' in val or 'committee' in val:
        return CertificateRecipient.TYPE_COMMITTEE
    return default_type


def _is_explicit_recipient_type(raw_val):
    val = str(raw_val or '').strip().lower()
    return any(k in val for k in ('วิทยากร', 'lecturer', 'กรรมการ', 'committee'))


def parse_recipients_upload(uploaded_file, default_type=CertificateRecipient.TYPE_LECTURER):
    """
    อ่านรายชื่อวิทยากร / คณะกรรมการ จากไฟล์ Excel (.xlsx) หรือ CSV (.csv)
    คืนค่า (rows_list, errors_list) โดยแต่ละแถวมี keys:
    {'recipient_type', 'full_name', 'school_province', 'description'}
    """
    name = (getattr(uploaded_file, 'name', '') or '').lower()
    raw_rows = []
    if name.endswith('.csv'):
        raw_bytes = uploaded_file.read()
        text = None
        for enc in ('utf-8-sig', 'cp874', 'tis-620', 'utf-8'):
            try:
                text = raw_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            return [], ['ไม่สามารถอ่านรหัสภาษาในไฟล์ CSV ได้ กรุณาบันทึกเป็น UTF-8 หรือ Excel (.xlsx)']
        reader = csv.reader(io.StringIO(text))
        for r in reader:
            raw_rows.append([str(c or '').strip() for c in r])
    elif name.endswith('.xlsx'):
        try:
            wb = load_workbook(uploaded_file, read_only=True, data_only=True)
            ws = wb.active
            for r in ws.iter_rows(values_only=True):
                raw_rows.append([str(c).strip() if c is not None else '' for c in r])
            wb.close()
        except Exception as exc:
            return [], [f'ไม่สามารถอ่านไฟล์ Excel ได้: {exc}']
    else:
        return [], ['รองรับเฉพาะไฟล์ .xlsx หรือ .csv เท่านั้น']

    # กรองแถวว่างออก
    raw_rows = [r for r in raw_rows if any(cell for cell in r)]
    if not raw_rows:
        return [], ['ไม่พบข้อมูลในไฟล์']

    # ตรวจแถวหัวตาราง
    header_map = {}
    first_row = raw_rows[0]
    for idx, cell in enumerate(first_row):
        key = _RECIPIENT_HEADER_ALIASES.get(cell.strip().lower())
        if not key:
            key = _RECIPIENT_HEADER_ALIASES.get(cell.strip())
        if key and key not in header_map:
            header_map[key] = idx

    has_header = bool(header_map)
    data_rows = raw_rows[1:] if has_header else raw_rows

    parsed, errors = [], []
    start_row_no = 2 if has_header else 1
    for offset, r in enumerate(data_rows):
        row_no = start_row_no + offset
        if not any(r):
            continue

        def _get(col_key):
            i = header_map.get(col_key)
            return r[i].strip() if (i is not None and i < len(r)) else ''

        if has_header:
            rtype = _normalize_recipient_type(_get('recipient_type'), default_type)
            full_name = _get('full_name')
            if not full_name:
                fn = _get('first_name')
                ln = _get('last_name')
                pref = _get('prefix')
                if fn or ln:
                    name_part = f'{pref}{fn}'.strip() if pref else fn
                    full_name = f'{name_part} {ln}'.strip()
            school = _get('school_province')
            prov = _get('province')
            if school and prov and prov not in school:
                school = f'{school} {prov}'
            elif not school and prov:
                school = prov
            desc = _get('description')
        else:
            # ไม่มีหัวตาราง: เช็กว่าคอลัมน์แรกเป็นประเภทหรือไม่
            if len(r) >= 2 and _is_explicit_recipient_type(r[0]):
                rtype = _normalize_recipient_type(r[0], default_type)
                full_name = r[1] if len(r) > 1 else ''
                school = r[2] if len(r) > 2 else ''
                desc = r[3] if len(r) > 3 else ''
            else:
                rtype = default_type
                full_name = r[0] if len(r) > 0 else ''
                school = r[1] if len(r) > 1 else ''
                desc = r[2] if len(r) > 2 else ''

        if not full_name:
            errors.append(f'แถวที่ {row_no}: ไม่มีชื่อ-นามสกุล')
            continue

        parsed.append({
            'recipient_type': rtype,
            'full_name': full_name[:200],
            'school_province': school[:255],
            'description': desc,
        })

    if not parsed and not errors:
        errors.append('ไม่พบรายชื่อที่นำเข้าได้ในไฟล์')
    return parsed, errors
