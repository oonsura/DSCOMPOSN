"""
Import / Export Excel และระบบ Backup
รองรับทั้ง:
1. ไฟล์รูปแบบมาตรฐานตามคู่มือ (มีแถวหัวตาราง camp1_code, first_name, ...)
2. ไฟล์รูปแบบระบบ COMPOSN เดิม (ใช้สิทธิ์ คอม.xlsx — 14 คอลัมน์ ไม่มีหัวตาราง:
   รหัส, ชื่อ, นามสกุล, โรงเรียน, วันเกิด, เดือนเกิด, ปีเกิด, ค่าย, ชั้นปัจจุบัน, +1..+5)
3. ไฟล์รูปแบบระบบ Food เดิม (ไม่มีหัวตาราง 4 คอลัมน์: รหัส, รหัสผ่าน, ชื่อ, นามสกุล)
4. การตรวจสอบข้อมูลล่วงหน้า (Preview) ก่อนยืนยันนำเข้า
"""
import datetime as dt
import re
from io import BytesIO

from django.conf import settings
from django.db import transaction
from django.http import HttpResponse
from django.utils import timezone
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import levels
from .models import CAMP_CODE_REGEX, ArchiveRecord, UseRecord, User, normalize_school_name

STUDENT_COLUMNS = [
    'camp1_code', 'full_name', 'nickname', 'school_name', 'birth_date',
    'academic_year', 'current_level', 'student_phone', 'parent_phone',
    'email', 'chronic_disease', 'food_allergy', 'is_muslim', 'note',
]
STUDENT_COLUMNS_TH = [
    'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'โรงเรียน', 'วันเกิด',
    'ปีการศึกษา', 'ระดับชั้นปัจจุบัน', 'เบอร์โทรนักเรียน', 'เบอร์โทรผู้ปกครอง',
    'อีเมล', 'โรคประจำตัว', 'แพ้อาหาร', 'ศาสนาอิสลาม', 'หมายเหตุ',
]
STUDENT_REQUIRED_COLUMNS = ['camp1_code', 'full_name']
STUDENT_IMPORT_ALLOWED_COLUMNS = set(STUDENT_COLUMNS) | {
    'prefix', 'first_name', 'last_name', 'camp_center',
    'year_plus_1', 'year_plus_2', 'year_plus_3', 'year_plus_4', 'year_plus_5', 'password',
}
STUDENT_FIELD_LABELS_TH = {
    **dict(zip(STUDENT_COLUMNS, STUDENT_COLUMNS_TH)),
    'camp2_code': 'รหัสค่าย 2',
    'prefix': 'คำนำหน้า',
    'first_name': 'ชื่อ',
    'last_name': 'นามสกุล',
    'camp_center': 'ศูนย์',
    'year_plus_1': 'ระดับชั้น +1 ปี',
    'year_plus_2': 'ระดับชั้น +2 ปี',
    'year_plus_3': 'ระดับชั้น +3 ปี',
    'year_plus_4': 'ระดับชั้น +4 ปี',
    'year_plus_5': 'ระดับชั้น +5 ปี',
    'password': 'รหัสผ่าน',
}
CAMP2_COLUMNS = ['camp1_code', 'camp2_code']
CAMP2_COLUMNS_TH = ['รหัสค่าย 1', 'รหัสค่าย 2']

# รองรับชื่อหัวตารางทั้งภาษาไทยและภาษาอังกฤษ (จับคู่ตามชื่อหัวคอลัมน์ ไม่ขึ้นกับลำดับคอลัมน์)
HEADER_ALIASES = {
    'camp1_code': 'camp1_code',
    'รหัส': 'camp1_code',
    'รหัสค่าย': 'camp1_code',
    'รหัสค่าย 1': 'camp1_code',
    'รหัสค่าย1': 'camp1_code',
    'รหัสประจำตัว': 'camp1_code',
    'รหัสประจำตัวค่าย 1': 'camp1_code',
    'รหัสประจำตัวค่าย1': 'camp1_code',
    'รหัสนักเรียน': 'camp1_code',
    'เลขประจำตัว': 'camp1_code',
    'เลขประจำตัวค่าย 1': 'camp1_code',
    'เลขประจำตัวค่าย1': 'camp1_code',
    'เลขประจำตัวค่าย': 'camp1_code',
    'username': 'camp1_code',
    'code': 'camp1_code',
    'camp2_code': 'camp2_code',
    'รหัสค่าย 2': 'camp2_code',
    'รหัสค่าย2': 'camp2_code',
    'เลขประจำตัวค่าย 2': 'camp2_code',
    'เลขประจำตัวค่าย2': 'camp2_code',
    'ชื่อ-นามสกุล': 'full_name',
    'ชื่อ - นามสกุล': 'full_name',
    'ชื่อ นามสกุล': 'full_name',
    'ชื่อนามสกุล': 'full_name',
    'ชื่อสกุล': 'full_name',
    'ชื่อ-สกุล': 'full_name',
    'ชื่อ - สกุล': 'full_name',
    'ชื่อ สกุล': 'full_name',
    'fullname': 'full_name',
    'full_name': 'full_name',
    'name': 'full_name',
    'คำนำหน้า': 'prefix',
    'คำนำหน้าชื่อ': 'prefix',
    'prefix': 'prefix',
    'title': 'prefix',
    'ชื่อ': 'first_name',
    'ชื่อจริง': 'first_name',
    'first_name': 'first_name',
    'firstname': 'first_name',
    'นามสกุล': 'last_name',
    'สกุล': 'last_name',
    'last_name': 'last_name',
    'lastname': 'last_name',
    'surname': 'last_name',
    'ชื่อเล่น': 'nickname',
    'nickname': 'nickname',
    'nick_name': 'nickname',
    'โรงเรียน': 'school_name',
    'ชื่อโรงเรียน': 'school_name',
    'สถานศึกษา': 'school_name',
    'school_name': 'school_name',
    'school': 'school_name',
    'วันเกิด': 'birth_date',
    'วัน/เดือน/ปีเกิด': 'birth_date',
    'วันเดือนปีเกิด': 'birth_date',
    'birth_date': 'birth_date',
    'birthdate': 'birth_date',
    'dob': 'birth_date',
    'birthday': 'birth_date',
    'ศูนย์': 'camp_center',
    'ศูนย์อบรม': 'camp_center',
    'camp_center': 'camp_center',
    'center': 'camp_center',
    'ปีการศึกษา': 'academic_year',
    'ปีการศึกษา (พ.ศ.)': 'academic_year',
    'ปี พ.ศ.': 'academic_year',
    'ปี': 'academic_year',
    'academic_year': 'academic_year',
    'year': 'academic_year',
    'ระดับชั้นปัจจุบัน': 'current_level',
    'ระดับชั้น (ปัจจุบัน)': 'current_level',
    'ชั้นปัจจุบัน': 'current_level',
    'ระดับชั้น': 'current_level',
    'ชั้น': 'current_level',
    'ชั้นปี': 'current_level',
    'current_level': 'current_level',
    'year_level': 'current_level',
    'grade': 'current_level',
    'grade_level': 'current_level',
    'level': 'current_level',
    'year_plus_1': 'year_plus_1',
    '+1': 'year_plus_1',
    'ระดับชั้น +1 ปี': 'year_plus_1',
    'ระดับชั้น +1': 'year_plus_1',
    'ปี +1': 'year_plus_1',
    'year_plus_2': 'year_plus_2',
    '+2': 'year_plus_2',
    'ระดับชั้น +2 ปี': 'year_plus_2',
    'ระดับชั้น +2': 'year_plus_2',
    'ปี +2': 'year_plus_2',
    'year_plus_3': 'year_plus_3',
    '+3': 'year_plus_3',
    'ระดับชั้น +3 ปี': 'year_plus_3',
    'ระดับชั้น +3': 'year_plus_3',
    'ปี +3': 'year_plus_3',
    'year_plus_4': 'year_plus_4',
    '+4': 'year_plus_4',
    'ระดับชั้น +4 ปี': 'year_plus_4',
    'ระดับชั้น +4': 'year_plus_4',
    'ปี +4': 'year_plus_4',
    'year_plus_5': 'year_plus_5',
    '+5': 'year_plus_5',
    'ระดับชั้น +5 ปี': 'year_plus_5',
    'ระดับชั้น +5': 'year_plus_5',
    'ปี +5': 'year_plus_5',
    'เบอร์โทรนักเรียน': 'student_phone',
    'เบอร์โทรศัพท์นักเรียน': 'student_phone',
    'เบอร์นักเรียน': 'student_phone',
    'เบอร์โทร': 'student_phone',
    'เบอร์โทรศัพท์': 'student_phone',
    'โทรศัพท์นักเรียน': 'student_phone',
    'โทรศัพท์': 'student_phone',
    'โทรนักเรียน': 'student_phone',
    'student_phone': 'student_phone',
    'phone': 'student_phone',
    'tel': 'student_phone',
    'mobile': 'student_phone',
    'เบอร์โทรผู้ปกครอง': 'parent_phone',
    'เบอร์โทรศัพท์ผู้ปกครอง': 'parent_phone',
    'เบอร์ผู้ปกครอง': 'parent_phone',
    'โทรศัพท์ผู้ปกครอง': 'parent_phone',
    'โทรผู้ปกครอง': 'parent_phone',
    'parent_phone': 'parent_phone',
    'อีเมล': 'email',
    'อีเมล์': 'email',
    'เมล': 'email',
    'email': 'email',
    'e-mail': 'email',
    'โรคประจำตัว': 'chronic_disease',
    'โรค': 'chronic_disease',
    'chronic_disease': 'chronic_disease',
    'disease': 'chronic_disease',
    'แพ้อาหาร': 'food_allergy',
    'อาหารที่แพ้': 'food_allergy',
    'การแพ้อาหาร': 'food_allergy',
    'food_allergy': 'food_allergy',
    'allergy': 'food_allergy',
    'ศาสนาอิสลาม': 'is_muslim',
    'อิสลาม': 'is_muslim',
    'อาหารฮาลาล': 'is_muslim',
    'ฮาลาล': 'is_muslim',
    'ศาสนา': 'is_muslim',
    'is_muslim': 'is_muslim',
    'muslim': 'is_muslim',
    'halal': 'is_muslim',
    'หมายเหตุ': 'note',
    'note': 'note',
    'notes': 'note',
    'remark': 'note',
    'remarks': 'note',
    'รหัสผ่าน': 'password',
    'password': 'password',
}
_COMPACT_HEADER_ALIASES = {
    re.sub(r'[\s_\-]+', '', k.lower()): v for k, v in HEADER_ALIASES.items()
}

HEADER_FILL = PatternFill('solid', fgColor='1F4E79')
HEADER_FONT = Font(bold=True, color='FFFFFF')


# =============================================================================
# Workbook helpers
# =============================================================================
def build_workbook(sheets):
    """sheets = [(title, headers, rows), ...]"""
    wb = Workbook()
    wb.remove(wb.active)
    for title, headers, rows in sheets:
        safe_title = re.sub(r'[:\\/?*\[\]]', '_', str(title or 'Sheet')).strip()[:31] or 'Sheet'
        ws = wb.create_sheet(title=safe_title)
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for row in rows:
            ws.append([_cell_value(v) for v in row])
        for i, h in enumerate(headers, start=1):
            width = max([len(str(h))] + [len(str(r[i - 1])) for r in rows[:300] if i - 1 < len(r)]) + 3
            ws.column_dimensions[get_column_letter(i)].width = min(max(width, 10), 50)
        ws.freeze_panes = 'A2'
    if not wb.worksheets:
        wb.create_sheet('Sheet1')
    return wb


def _cell_value(v):
    if isinstance(v, dt.datetime):
        if timezone.is_aware(v):
            v = timezone.localtime(v)
        return v.replace(tzinfo=None)
    return v


def workbook_response(wb, filename):
    buf = BytesIO()
    wb.save(buf)
    resp = HttpResponse(buf.getvalue(),
                        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    from urllib.parse import quote
    resp['Content-Disposition'] = f"attachment; filename*=UTF-8''{quote(filename)}"
    return resp


def template_workbook(columns, sample):
    rows = sample if (sample and isinstance(sample[0], (list, tuple))) else [sample]
    return build_workbook([('ตัวอย่างข้อมูล', columns, rows)])


def students_template_workbook(lang='th'):
    """สร้างไฟล์ Template Excel สำหรับนำเข้านักเรียน (ครบทุกข้อมูลที่ต้องการเก็บ ยกเว้นรหัสค่าย 2 พร้อมแผ่นคำอธิบาย)"""
    headers = STUDENT_COLUMNS_TH if lang == 'th' else STUDENT_COLUMNS
    sample_rows = [
        [
            '69DS01', 'ด.ช.สมชาย ใจดี', 'ต้น', 'โรงเรียนเทพศิรินทร์', '15/03/2555',
            2569, 'ม.1', '081-234-5678', '089-876-5432', 'somchai@example.com',
            '', 'กุ้ง', 'ไม่ใช่', '',
        ],
        [
            '69DS02', 'น.ส.สมหญิง รักเรียน', 'หญิง', 'โรงเรียนสาธิต สจล.', '20/08/2553',
            2569, 'ม.4', '082-345-6789', '086-543-2109', 'somying@example.com',
            'หอบหืด', '', 'ใช่', 'ทานอาหารฮาลาล',
        ],
    ]
    guide_headers = ['ลำดับคอลัมน์', 'ชื่อหัวตาราง (ภาษาไทย)', 'ชื่อหัวตาราง (ภาษาอังกฤษ)', 'คำอธิบายและรูปแบบข้อมูล', 'ตัวอย่าง']
    guide_rows = [
        ['A', 'รหัสค่าย 1', 'camp1_code', 'รหัสประจำตัวค่าย 1 (เลขปี 2 หลัก + DS หรือ KM + เลข 2 หลัก) ใช้เป็น Username และรหัสผ่านเริ่มต้น (บังคับ)', '69DS01'],
        ['B', 'ชื่อ-นามสกุล', 'full_name', 'ชื่อและนามสกุลในช่องเดียวกัน หรือแยกเป็นคอลัมน์ "ชื่อ" และ "นามสกุล" ก็ได้ (บังคับ)', 'ด.ช.สมชาย ใจดี'],
        ['C', 'ชื่อเล่น', 'nickname', 'ชื่อเล่นของนักเรียน (เว้นว่างได้)', 'ต้น'],
        ['D', 'โรงเรียน', 'school_name', 'ชื่อโรงเรียนปัจจุบัน (เว้นว่างได้ — หากไม่มีคำว่า "โรงเรียน" นำหน้า ระบบจะเติมให้อัตโนมัติ)', 'โรงเรียนเทพศิรินทร์'],
        ['E', 'วันเกิด', 'birth_date', 'วัน/เดือน/ปีเกิด (พ.ศ. หรือ ค.ศ.) เช่น 15/03/2555 (เว้นว่างได้)', '15/03/2555'],
        ['F', 'ปีการศึกษา', 'academic_year', 'ปีการศึกษา พ.ศ. 4 หลัก (หากเว้นว่าง ระบบจะคำนวณจาก 2 หลักแรกของรหัสค่าย 1)', '2569'],
        ['G', 'ระดับชั้นปัจจุบัน', 'current_level', 'รองรับ ป.1-ป.6, ม.1-ม.6, ปวช.1-ปวช.3 (ระบบคำนวณชั้นปี +1 ถึง +5 ให้อัตโนมัติ)', 'ม.1'],
        ['H', 'เบอร์โทรนักเรียน', 'student_phone', 'เบอร์โทรศัพท์ติดต่อนักเรียน (เว้นว่างได้)', '081-234-5678'],
        ['I', 'เบอร์โทรผู้ปกครอง', 'parent_phone', 'เบอร์โทรศัพท์ติดต่อผู้ปกครอง (เว้นว่างได้)', '089-876-5432'],
        ['J', 'อีเมล', 'email', 'อีเมลติดต่อ (เว้นว่างได้)', 'somchai@example.com'],
        ['K', 'โรคประจำตัว', 'chronic_disease', 'ระบุโรคประจำตัว หรือเว้นว่างหากไม่มี', 'หอบหืด'],
        ['L', 'แพ้อาหาร', 'food_allergy', 'ระบุอาหารที่แพ้ หรือเว้นว่างหากไม่มี', 'กุ้ง'],
        ['M', 'ศาสนาอิสลาม', 'is_muslim', 'ระบุ ใช่/ไม่ใช่ (หรือ 1/0, อิสลาม/ฮาลาล) เพื่อใช้จัดเตรียมอาหารฮาลาล', 'ใช่ / ไม่ใช่'],
        ['N', 'หมายเหตุ', 'note', 'ข้อมูลเพิ่มเติมหรือหมายเหตุอื่น ๆ (เว้นว่างได้)', 'ทานอาหารฮาลาล'],
        ['-', '-', '-', 'หมายเหตุ: คอลัมน์ในไฟล์ที่นำเข้าสามารถเรียงลำดับสลับกันได้ ระบบจะจับคู่ลงฐานข้อมูลตามชื่อหัวคอลัมน์ (ไม่รวมรหัสค่าย 2)', '-'],
    ]
    return build_workbook([
        ('รายชื่อนักเรียน', headers, sample_rows),
        ('คำอธิบายการกรอกข้อมูล', guide_headers, guide_rows),
    ])


def camp2_template_workbook(lang='th'):
    """สร้างไฟล์ Template Excel สำหรับนำเข้ารหัสค่าย 2"""
    headers = CAMP2_COLUMNS_TH if lang == 'th' else CAMP2_COLUMNS
    sample_rows = [
        ['69DS01', '69KM01'],
        ['69DS02', '69KM02'],
    ]
    guide_headers = ['ลำดับคอลัมน์', 'ชื่อหัวตาราง (ภาษาไทย)', 'ชื่อหัวตาราง (ภาษาอังกฤษ)', 'คำอธิบาย', 'ตัวอย่าง']
    guide_rows = [
        ['A', 'รหัสค่าย 1', 'camp1_code', 'รหัสค่าย 1 ของนักเรียนที่มีอยู่ในระบบแล้ว', '69DS01'],
        ['B', 'รหัสค่าย 2', 'camp2_code', 'รหัสค่าย 2 ที่ต้องการผูกให้ (เมื่อมีรหัสค่าย 2 แล้วจะใช้สิทธิ์ค่าย 2 เท่านั้น)', '69KM01'],
    ]
    return build_workbook([
        ('รหัสค่าย 2', headers, sample_rows),
        ('คำอธิบาย', guide_headers, guide_rows),
    ])


def rights_template_workbook(lang='th'):
    """สร้างไฟล์ Template Excel สำหรับนำเข้าผู้ใช้สิทธิ์ค่าย 1 ในปีการศึกษานั้น"""
    headers = ['รหัสค่าย 1'] if lang == 'th' else ['camp1_code']
    sample_rows = [
        ['67DS01'],
        ['68DS02'],
    ]
    guide_headers = ['ลำดับคอลัมน์', 'ชื่อหัวตาราง (ภาษาไทย)', 'ชื่อหัวตาราง (ภาษาอังกฤษ)', 'คำอธิบาย', 'ตัวอย่าง']
    guide_rows = [
        ['A', 'รหัสค่าย 1', 'camp1_code', 'รหัสค่าย 1 ของนักเรียนที่ต้องการนำเข้าเป็นผู้ใช้สิทธิ์ค่ายที่ 1 ในปีการศึกษาปัจจุบัน', '67DS01'],
    ]
    return build_workbook([
        ('ผู้ใช้สิทธิ์ค่าย 1', headers, sample_rows),
        ('คำอธิบาย', guide_headers, guide_rows),
    ])


def read_rights_codes(uploaded_file):
    """อ่านรหัสค่าย 1 จากไฟล์ Excel หรือ CSV (รองรับทั้งมีหัวตารางและไม่มีหัวตาราง)"""
    rows = _load_sheet_rows(uploaded_file)
    result = []
    for idx, row in enumerate(rows, start=1):
        if row is None or all(v in (None, '') for v in row):
            continue
        cell0 = _clean(row[0])
        if not cell0:
            continue
        # ข้ามแถวหัวตารางถ้าพบคำว่า รหัส / camp1_code
        if idx == 1 and (cell0.lower() in HEADER_ALIASES or 'รหัส' in cell0 or 'camp' in cell0.lower()):
            continue
        result.append((idx, cell0.upper().replace(' ', '')))
    if not result:
        raise ValueError('ไม่พบรหัสค่าย 1 ในไฟล์')
    return result


# =============================================================================
# Parsing helpers
# =============================================================================
def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _clean_optional_text(v, clear_no=False):
    s = _clean(v)
    if not s or s in ('-', '—', '–'):
        return ''
    if clear_no and s.lower() in ('ไม่มี', 'none', 'n/a', 'na'):
        return ''
    return s


def _clean_phone(v):
    s = _clean_optional_text(v)
    if not s:
        return ''
    digits = re.sub(r'[\s\-]+', '', s)
    if digits.isdigit() and ((len(digits) == 9 and digits[0] in '689') or (len(digits) == 8 and digits[0] == '2')):
        return '0' + s
    return s


def parse_bool_muslim(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    s = _clean(v).lower()
    if not s or s in ('0', 'false', 'no', 'n', 'ไม่ใช่', 'ไม่', '-', '—', '–', 'none', 'ไม่มี', 'พุทธ', 'คริสต์', 'ฮินดู', 'ซิกข์'):
        return False
    if s in ('1', 'true', 'yes', 'y', 'ใช่', '✓', '✔', 'x'):
        return True
    return any(k in s for k in ('อิสลาม', 'มุสลิม', 'ฮาลาล', 'islam', 'muslim', 'halal'))


def _resolve_header(raw_header):
    """จับคู่ชื่อหัวคอลัมน์จาก Excel/CSV (ภาษาไทย/อังกฤษ/มีวงเล็บ/เว้นวรรคต่างกัน) เป็นชื่อฟิลด์มาตรฐาน"""
    norm = ' '.join(_clean(raw_header).lower().split())
    if not norm:
        return None
    if norm in HEADER_ALIASES:
        return HEADER_ALIASES[norm]
    stripped = re.sub(r'\s*[\(\[].*?[\)\]]', '', norm).strip()
    if stripped in HEADER_ALIASES:
        return HEADER_ALIASES[stripped]
    compact = re.sub(r'[\s_\-]+', '', stripped)
    return _COMPACT_HEADER_ALIASES.get(compact)


def _int_or_none(v):
    s = _clean(v)
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def parse_date(value):
    """รองรับ date/datetime, 'dd/mm/yyyy', 'yyyy-mm-dd' และปี พ.ศ. (แปลงเป็น ค.ศ.)"""
    if value in (None, ''):
        return None
    if isinstance(value, dt.datetime):
        d = value.date()
    elif isinstance(value, dt.date):
        d = value
    else:
        s = str(value).strip()
        m = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$', s)
        m2 = re.match(r'^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', s)
        try:
            if m:
                d = dt.date(int(m.group(3)) - (543 if int(m.group(3)) > 2400 else 0),
                            int(m.group(2)), int(m.group(1)))
            elif m2:
                d = dt.date(int(m2.group(1)) - (543 if int(m2.group(1)) > 2400 else 0),
                            int(m2.group(2)), int(m2.group(3)))
            else:
                raise ValueError
        except ValueError:
            raise ValueError(f'รูปแบบวันเกิดไม่ถูกต้อง: {s}')
        return d
    if d.year > 2400:
        try:
            d = d.replace(year=d.year - 543)
        except ValueError:
            raise ValueError(f'รูปแบบวันเกิดไม่ถูกต้อง: {value}')
    return d


def _load_sheet_rows(uploaded_file):
    fname = getattr(uploaded_file, 'name', '') or ''
    if hasattr(uploaded_file, 'seek'):
        uploaded_file.seek(0)
    if fname.lower().endswith('.csv'):
        import csv
        raw = uploaded_file.read()
        if isinstance(raw, bytes):
            for enc in ('utf-8-sig', 'utf-8', 'cp874', 'tis-620'):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                text = raw.decode('utf-8', errors='replace')
        else:
            text = str(raw)
        rows = [tuple(r) for r in csv.reader(text.splitlines())]
        if not rows:
            raise ValueError('ไฟล์ว่างเปล่า')
        return rows
    try:
        wb = load_workbook(uploaded_file, read_only=True, data_only=True)
    except Exception:
        raise ValueError('ไม่สามารถอ่านไฟล์ได้ กรุณาใช้ไฟล์ Excel (.xlsx) หรือ CSV (.csv)')
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise ValueError('ไฟล์ว่างเปล่า')
    return rows


EMAIL_RE = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')


def _effective_required_columns(required_columns):
    if required_columns == STUDENT_COLUMNS:
        return STUDENT_REQUIRED_COLUMNS
    return required_columns


def _missing_required_columns(canon_set, required_columns):
    missing = []
    for c in _effective_required_columns(required_columns):
        if c == 'full_name':
            if 'full_name' not in canon_set and not ('first_name' in canon_set and 'last_name' in canon_set):
                missing.append('ชื่อ-นามสกุล')
        elif c != 'camp_center' and c not in canon_set:
            missing.append(STUDENT_FIELD_LABELS_TH.get(c, c))
    return missing


def read_rows_with_meta(uploaded_file, required_columns, allowed_columns=None):
    """อ่านไฟล์ Excel/CSV คืนค่า (rows_list, meta)
    - rows_list: [(row_number, dict), ...]
    - จับคู่คอลัมน์ตามชื่อหัวตาราง (ไม่ขึ้นกับลำดับคอลัมน์ในไฟล์)
    - รองรับไฟล์จากระบบ COMPOSN/Food เดิมที่ไม่มีแถวหัวตาราง"""
    rows = _load_sheet_rows(uploaded_file)
    non_empty = [
        (idx, row)
        for idx, row in enumerate(rows, start=1)
        if row is not None and any(v not in (None, '') for v in row)
    ]
    if not non_empty:
        raise ValueError('ไฟล์ว่างเปล่า')

    first_row_vals = [_clean(c) for c in non_empty[0][1]]
    first_cell = first_row_vals[0].upper() if first_row_vals else ''
    is_headerless_code = bool(re.match(CAMP_CODE_REGEX, first_cell))
    is_student_req = required_columns in (STUDENT_COLUMNS, STUDENT_REQUIRED_COLUMNS)

    if is_headerless_code and is_student_req:
        result = []
        present_set = []
        seen_present = set()
        for idx, row in non_empty:
            vals = list(row) + [None] * 14
            code = _clean(vals[0]).upper().replace(' ', '')
            col4 = _clean(vals[4]).upper()
            col5 = _clean(vals[5]).upper()
            col5_lv = levels.normalize_level(vals[5]) if vals[5] is not None else None
            col6_lv = levels.normalize_level(vals[6]) if vals[6] is not None else None
            col7_lv = levels.normalize_level(vals[7]) if vals[7] is not None else None
            col8_lv = levels.normalize_level(vals[8]) if vals[8] is not None else None

            # กรณี 1A: รูปแบบใหม่ ชื่อ-นามสกุลเป็นช่องเดียว (6-7 คอลัมน์):
            # 0:รหัส, 1:ชื่อ-นามสกุล, 2:โรงเรียน, 3:วันเกิด, 4:ปีการศึกษา(หรือศูนย์), 5:ระดับชั้น(หรือปีการศึกษา), 6:ระดับชั้น
            if col5_lv and not col7_lv and not col8_lv:
                m_yr = re.match(r'^(\d{2})', code)
                acad_yr = _int_or_none(vals[4]) or ((2500 + int(m_yr.group(1))) if m_yr else None)
                data = {
                    'camp1_code': code,
                    'full_name': vals[1],
                    'school_name': vals[2],
                    'birth_date': vals[3],
                    'camp_center': code[2:4] if len(code) >= 4 else '',
                    'academic_year': acad_yr,
                    'current_level': vals[5],
                }
            elif col4 in ('DS', 'KM') and col6_lv:
                m_yr = re.match(r'^(\d{2})', code)
                acad_yr = _int_or_none(vals[5]) or ((2500 + int(m_yr.group(1))) if m_yr else None)
                data = {
                    'camp1_code': code,
                    'full_name': vals[1],
                    'school_name': vals[2],
                    'birth_date': vals[3],
                    'camp_center': col4,
                    'academic_year': acad_yr,
                    'current_level': vals[6],
                }
            # กรณีที่ 1B: รูปแบบ 8 คอลัมน์โดยไม่มีหัวตาราง (แยกชื่อ/นามสกุล)
            elif col5 in ('DS', 'KM') or (col7_lv and not col8_lv):
                m_yr = re.match(r'^(\d{2})', code)
                acad_yr = _int_or_none(vals[6]) or ((2500 + int(m_yr.group(1))) if m_yr else None)
                data = {
                    'camp1_code': code,
                    'first_name': vals[1],
                    'last_name': vals[2],
                    'school_name': vals[3],
                    'birth_date': vals[4],
                    'camp_center': col5 if col5 in ('DS', 'KM') else (code[2:4] if len(code) >= 4 else ''),
                    'academic_year': acad_yr,
                    'current_level': vals[7],
                    'year_plus_1': vals[8],
                    'year_plus_2': vals[9],
                    'year_plus_3': vals[10],
                    'year_plus_4': vals[11],
                    'year_plus_5': vals[12],
                }
            # กรณีที่ 2: รูปแบบ COMPOSN เดิม (9-14 คอลัมน์ ไม่มีหัวตาราง)
            elif len(row) >= 9 or col8_lv:
                bd, bm, by = _int_or_none(vals[4]), _int_or_none(vals[5]), _int_or_none(vals[6])
                birth_val = f'{bd:02d}/{bm:02d}/{by}' if (bd and bm and by) else (vals[4] if vals[4] and not bm else None)
                m_yr = re.match(r'^(\d{2})', code)
                acad_yr = (2500 + int(m_yr.group(1))) if m_yr else None
                data = {
                    'camp1_code': code,
                    'first_name': vals[1],
                    'last_name': vals[2],
                    'school_name': vals[3],
                    'birth_date': birth_val,
                    'camp_center': code[2:4] if len(code) >= 4 else '',
                    'academic_year': acad_yr,
                    'current_level': vals[8],
                    'year_plus_1': vals[9],
                    'year_plus_2': vals[10],
                    'year_plus_3': vals[11],
                    'year_plus_4': vals[12],
                    'year_plus_5': vals[13],
                }
            else:
                # กรณีที่ 3: รูปแบบไม่มีหัวตารางแบบย่อ (2-5 คอลัมน์: รหัส, ชื่อ[-สกุล], [นามสกุล/โรงเรียน])
                m_yr = re.match(r'^(\d{2})', code)
                acad_yr = (2500 + int(m_yr.group(1))) if m_yr else None
                col1 = _clean(vals[1])
                if ' ' in col1:
                    data = {
                        'camp1_code': code,
                        'full_name': vals[1],
                        'school_name': vals[2] or '',
                        'birth_date': None,
                        'camp_center': code[2:4] if len(code) >= 4 else '',
                        'academic_year': acad_yr,
                        'current_level': '',
                    }
                else:
                    data = {
                        'camp1_code': code,
                        'first_name': vals[1],
                        'last_name': vals[2] or '',
                        'school_name': vals[3] or '',
                        'birth_date': None,
                        'camp_center': code[2:4] if len(code) >= 4 else '',
                        'academic_year': acad_yr,
                        'current_level': '',
                    }
            for k in data:
                if k not in seen_present:
                    seen_present.add(k)
                    present_set.append(k)
            result.append((idx, data))
        return result, {
            'has_header': False,
            'matched_columns': [],
            'ignored_columns': [],
            'present_fields': present_set,
        }

    if is_headerless_code and required_columns == CAMP2_COLUMNS and len(first_row_vals) >= 2:
        result = []
        for idx, row in non_empty:
            result.append((idx, {'camp1_code': row[0], 'camp2_code': row[1] if len(row) > 1 else None}))
        return result, {
            'has_header': False,
            'matched_columns': [],
            'ignored_columns': [],
            'present_fields': ['camp1_code', 'camp2_code'],
        }

    # ค้นหาแถวหัวตารางในช่วง 5 แถวแรกที่ไม่ว่าง (รองรับไฟล์ที่มีแถวหัวเรื่อง/ชื่อรายงานด้านบน)
    header_pos = None
    header_resolved = []
    for pos, (_, row) in enumerate(non_empty[:5]):
        resolved = [_resolve_header(c) for c in row]
        canon_set = {k for k in resolved if k and (allowed_columns is None or k in allowed_columns)}
        if not _missing_required_columns(canon_set, required_columns):
            header_pos = pos
            header_resolved = resolved
            break

    if header_pos is None:
        first_resolved = {_resolve_header(c) for c in non_empty[0][1]}
        first_resolved.discard(None)
        if allowed_columns is not None:
            first_resolved &= set(allowed_columns)
        missing = _missing_required_columns(first_resolved, required_columns)
        raise ValueError('ไฟล์ขาดคอลัมน์: ' + ', '.join(missing))

    header_row = non_empty[header_pos][1]
    col_mapping = {}
    matched_columns = []
    ignored_columns = []
    seen_canon = set()
    for i, raw_h in enumerate(header_row):
        orig_h = _clean(raw_h)
        canon = header_resolved[i] if i < len(header_resolved) else None
        if canon and (allowed_columns is None or canon in allowed_columns) and canon not in seen_canon:
            seen_canon.add(canon)
            col_mapping[i] = canon
            matched_columns.append({
                'col_letter': get_column_letter(i + 1),
                'header': orig_h,
                'field': canon,
                'label': STUDENT_FIELD_LABELS_TH.get(canon, canon),
            })
        elif orig_h:
            if canon == 'camp2_code' and allowed_columns is not None and 'camp2_code' not in allowed_columns:
                ignored_columns.append(f'{orig_h} (ข้าม: ใช้เมนู Import รหัสค่าย 2)')
            else:
                ignored_columns.append(orig_h)

    result = []
    for excel_rownum, row in non_empty[header_pos + 1:]:
        data = {canon: (row[i] if i < len(row) else None) for i, canon in col_mapping.items()}
        if all(v in (None, '') for v in data.values()):
            continue
        result.append((excel_rownum, data))

    return result, {
        'has_header': True,
        'matched_columns': matched_columns,
        'ignored_columns': ignored_columns,
        'present_fields': list(seen_canon),
    }


def read_rows(uploaded_file, required_columns):
    """อ่านไฟล์ Excel คืนค่า list ของ (row_number, dict)
    รองรับทั้งไฟล์ที่มีหัวตาราง (เรียงลำดับคอลัมน์ใดก็ได้) และไฟล์ที่ไม่มีหัวตาราง"""
    rows, _ = read_rows_with_meta(uploaded_file, required_columns)
    return rows


def _extract_names(d):
    """ดึง first_name, last_name จาก dict ข้อมูลแถว (รองรับทั้งคอลัมน์ full_name ช่องเดียว และแยก first_name/last_name/prefix)"""
    full = _clean(d.get('full_name'))
    if full:
        parts = full.split(None, 1)
        return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], '')
    first = _clean(d.get('first_name'))
    last = _clean(d.get('last_name'))
    pref = _clean(d.get('prefix'))
    if pref and first and not first.startswith(pref):
        first = f'{pref}{first}'
    elif pref and not first:
        first = pref
    return first, last


def _parse_student_row(rownum, d, seen, existing_camp2):
    """ตรวจสอบและแปลงข้อมูลนักเรียน 1 แถวให้อยู่ในรูปแบบมาตรฐานสำหรับ Preview และ Import"""
    code = _clean(d.get('camp1_code')).upper()
    fname, lname = _extract_names(d)
    nickname = _clean_optional_text(d.get('nickname'))[:60]
    school = normalize_school_name(_clean(d.get('school_name')))
    center = _clean(d.get('camp_center')).upper() or (code[2:4] if len(code) >= 4 else '')
    year_raw = _clean(d.get('academic_year'))
    level_raw = _clean(d.get('current_level'))
    student_phone = _clean_phone(d.get('student_phone'))[:30]
    parent_phone = _clean_phone(d.get('parent_phone'))[:30]
    email = _clean_optional_text(d.get('email'))[:150]
    chronic_disease = _clean_optional_text(d.get('chronic_disease'), clear_no=True)[:255]
    food_allergy = _clean_optional_text(d.get('food_allergy'), clear_no=True)[:255]
    is_muslim = parse_bool_muslim(d.get('is_muslim'))
    note = _clean(d.get('note'))

    row_errors = []
    warnings = []

    if not re.match(CAMP_CODE_REGEX, code):
        row_errors.append(f'รหัส "{code}" ไม่ตรงรูปแบบ (เช่น 69DS01)')
    if code in seen:
        warnings.append(f'ซ้ำกับแถว {seen[code]} (จะใช้แถวล่าสุด)')
    seen[code] = rownum

    if center and center not in ('DS', 'KM'):
        row_errors.append(f'ศูนย์ต้องเป็น DS หรือ KM (พบ "{center}")')

    year = None
    if year_raw:
        if not re.match(r'^\d{4}$', year_raw):
            row_errors.append(f'ปีการศึกษาต้องเป็นตัวเลข 4 หลัก (พบ "{year_raw}")')
        else:
            year = int(year_raw)
            if year < 2400:
                year += 543
    elif re.match(r'^\d{2}', code):
        year = 2500 + int(code[:2])

    level = levels.normalize_level(level_raw) if level_raw else ''
    if level_raw and level is None:
        row_errors.append(f'ระดับชั้น "{level_raw}" ไม่รู้จัก')

    birth = None
    try:
        birth = parse_date(d.get('birth_date'))
    except ValueError as e:
        row_errors.append(str(e))

    if email and not EMAIL_RE.match(email):
        row_errors.append(f'รูปแบบอีเมลไม่ถูกต้อง: "{email}"')

    if code in existing_camp2 and existing_camp2[code] != code:
        row_errors.append(f'ถูกใช้เป็นรหัสค่าย 2 ของ {existing_camp2[code]} แล้ว')

    return {
        'rownum': rownum,
        'camp1_code': code,
        'first_name': fname,
        'last_name': lname,
        'nickname': nickname,
        'school_name': school,
        'birth_date': birth.isoformat() if birth else '',
        'camp_center': center,
        'academic_year': year or '',
        'current_level': level or '',
        'year_plus_1': _clean(d.get('year_plus_1')),
        'year_plus_2': _clean(d.get('year_plus_2')),
        'year_plus_3': _clean(d.get('year_plus_3')),
        'year_plus_4': _clean(d.get('year_plus_4')),
        'year_plus_5': _clean(d.get('year_plus_5')),
        'student_phone': student_phone,
        'parent_phone': parent_phone,
        'email': email,
        'chronic_disease': chronic_disease,
        'food_allergy': food_allergy,
        'is_muslim': is_muslim,
        'note': note,
        'password': _clean(d.get('password')),
        'valid': len(row_errors) == 0,
        'errors': row_errors,
        'warnings': warnings,
    }


def _save_student_row(r, present_fields=None, recalc_levels=True):
    """บันทึก/อัปเดต User 1 แถวจากข้อมูลที่ผ่านการ parse แล้ว คืนค่า 'created' หรือ 'updated'"""
    code = r['camp1_code']
    birth = dt.date.fromisoformat(r['birth_date']) if r.get('birth_date') else None
    year = int(r['academic_year']) if r.get('academic_year') else None
    all_fields = {
        'first_name': r['first_name'],
        'last_name': r['last_name'],
        'nickname': r.get('nickname', ''),
        'school_name': normalize_school_name(r.get('school_name', '')),
        'birth_date': birth,
        'camp_center': r.get('camp_center', ''),
        'academic_year': year,
        'current_level': r.get('current_level', ''),
        'student_phone': r.get('student_phone', ''),
        'parent_phone': r.get('parent_phone', ''),
        'email': r.get('email', ''),
        'chronic_disease': r.get('chronic_disease', ''),
        'food_allergy': r.get('food_allergy', ''),
        'is_muslim': bool(r.get('is_muslim', False)),
        'note': r.get('note', ''),
    }
    user = User.objects.filter(camp1_code=code).first()
    if user:
        if user.role != User.ROLE_MEMBER:
            raise ValueError(f'{code} เป็นบัญชี {user.get_role_display()} ไม่สามารถทับได้')
        if present_fields is None:
            for k, v in all_fields.items():
                setattr(user, k, v)
        else:
            if 'full_name' in present_fields:
                user.first_name = all_fields['first_name']
                user.last_name = all_fields['last_name']
            else:
                if 'first_name' in present_fields or 'prefix' in present_fields:
                    user.first_name = all_fields['first_name']
                if 'last_name' in present_fields:
                    user.last_name = all_fields['last_name']
            for k in (
                'nickname', 'school_name', 'birth_date', 'academic_year', 'current_level',
                'student_phone', 'parent_phone', 'email', 'chronic_disease',
                'food_allergy', 'is_muslim', 'note',
            ):
                if k in present_fields:
                    setattr(user, k, all_fields[k])
            if 'camp_center' in present_fields or not user.camp_center:
                user.camp_center = all_fields['camp_center']

        if recalc_levels and user.current_level and (
            present_fields is None or 'current_level' in present_fields or not user.year_plus_1
        ):
            levels.apply_levels(user)
        elif not recalc_levels:
            for yp in ('year_plus_1', 'year_plus_2', 'year_plus_3', 'year_plus_4', 'year_plus_5'):
                if (present_fields is None or yp in present_fields) and r.get(yp):
                    setattr(user, yp, r[yp])
        if user.is_first_login and not user.check_password(code):
            user.set_password(code)
        user.save()
        return 'updated'

    user = User(camp1_code=code, role=User.ROLE_MEMBER, is_first_login=True, **all_fields)
    user.set_password(code)
    if recalc_levels and user.current_level:
        levels.apply_levels(user)
    else:
        for yp in ('year_plus_1', 'year_plus_2', 'year_plus_3', 'year_plus_4', 'year_plus_5'):
            if r.get(yp):
                setattr(user, yp, r[yp])
    user.save()
    return 'created'


# =============================================================================
# Import & Preview
# =============================================================================
def preview_students(uploaded_file):
    """ตรวจสอบข้อมูลนักเรียนในไฟล์ Excel ก่อนนำเข้าจริง คืนค่ารายการแถวพร้อมสถานะถูกต้อง/ข้อผิดพลาด"""
    rows, meta = read_rows_with_meta(
        uploaded_file, STUDENT_REQUIRED_COLUMNS, allowed_columns=STUDENT_IMPORT_ALLOWED_COLUMNS
    )
    preview_list = []
    valid_count = 0
    invalid_count = 0
    seen = {}
    existing_codes = set(User.objects.values_list('camp1_code', flat=True))
    existing_camp2 = dict(User.objects.exclude(camp2_code__isnull=True).exclude(camp2_code='')
                          .values_list('camp2_code', 'camp1_code'))

    for rownum, d in rows:
        parsed = _parse_student_row(rownum, d, seen, existing_camp2)
        parsed['is_update'] = parsed['camp1_code'] in existing_codes
        if parsed['valid']:
            valid_count += 1
        else:
            invalid_count += 1
        preview_list.append(parsed)

    return {
        'rows': preview_list,
        'total': len(rows),
        'valid_count': valid_count,
        'invalid_count': invalid_count,
        'matched_columns': meta['matched_columns'],
        'ignored_columns': meta['ignored_columns'],
        'present_fields': meta['present_fields'],
    }


def import_students_from_preview(pdata, recalc_levels=True):
    """นำเข้าข้อมูลนักเรียนจากผล Preview ที่เก็บใน Session"""
    rows = pdata.get('rows', [])
    present_fields = set(pdata.get('present_fields', [])) or None
    created, updated, errors = 0, 0, []
    with transaction.atomic():
        for r in rows:
            if not r.get('valid'):
                errors.append((r.get('rownum'), '; '.join(r.get('errors', []))))
                continue
            try:
                status = _save_student_row(r, present_fields=present_fields, recalc_levels=recalc_levels)
                if status == 'created':
                    created += 1
                else:
                    updated += 1
            except ValueError as e:
                errors.append((r.get('rownum'), str(e)))
    return {'created': created, 'updated': updated, 'errors': errors, 'total': len(rows)}


def import_students(uploaded_file, recalc_levels=True):
    """Import ข้อมูลนักเรียน: camp1_code ซ้ำ = ทับข้อมูลเดิม, ใหม่ = สร้าง User (รหัสผ่าน = camp1_code)"""
    rows, meta = read_rows_with_meta(
        uploaded_file, STUDENT_REQUIRED_COLUMNS, allowed_columns=STUDENT_IMPORT_ALLOWED_COLUMNS
    )
    present_fields = set(meta.get('present_fields', [])) or None
    existing_camp2 = dict(User.objects.exclude(camp2_code__isnull=True).exclude(camp2_code='')
                          .values_list('camp2_code', 'camp1_code'))
    created, updated, errors = 0, 0, []
    seen = {}
    with transaction.atomic():
        for rownum, d in rows:
            parsed = _parse_student_row(rownum, d, seen, existing_camp2)
            for w in parsed['warnings']:
                errors.append((rownum, w))
            if not parsed['valid']:
                errors.append((rownum, '; '.join(parsed['errors'])))
                continue
            try:
                status = _save_student_row(parsed, present_fields=present_fields, recalc_levels=recalc_levels)
                if status == 'created':
                    created += 1
                else:
                    updated += 1
            except ValueError as e:
                errors.append((rownum, str(e)))
    return {'created': created, 'updated': updated, 'errors': errors, 'total': len(rows)}


def import_camp2(uploaded_file):
    """Import รหัสค่าย 2: ค้นหา User จาก camp1_code แล้วอัปเดต camp2_code (รองรับคอลัมน์สลับลำดับ)"""
    from food.models import ScoreApproval, camp2_code_year
    rows, _ = read_rows_with_meta(
        uploaded_file, CAMP2_COLUMNS, allowed_columns={'camp1_code', 'camp2_code'}
    )
    updated, errors = 0, []
    with transaction.atomic():
        for rownum, d in rows:
            c1 = _clean(d.get('camp1_code')).upper()
            c2 = _clean(d.get('camp2_code')).upper()
            try:
                if not re.match(CAMP_CODE_REGEX, c1):
                    raise ValueError(f'camp1_code "{c1}" ไม่ตรงรูปแบบ')
                if not re.match(CAMP_CODE_REGEX, c2):
                    raise ValueError(f'camp2_code "{c2}" ไม่ตรงรูปแบบ (เช่น 69KM01)')
                new_yr = camp2_code_year(c2)
                if new_yr and ScoreApproval.is_year_approved(new_yr):
                    raise ValueError(f'ปีการศึกษา {new_yr} อนุมัติประกาศผลแล้ว ไม่สามารถแก้ไขรหัสค่าย 2 ({c2}) ได้ กรุณายกเลิกประกาศผลก่อน')
                user = User.objects.filter(camp1_code=c1, role=User.ROLE_MEMBER).first()
                if not user:
                    raise ValueError(f'ไม่พบผู้ใช้ camp1_code {c1}')
                old_yr = camp2_code_year(user.camp2_code)
                if old_yr and ScoreApproval.is_year_approved(old_yr) and user.camp2_code != c2:
                    raise ValueError(f'{c1} มีรหัสค่าย 2 ({user.camp2_code}) ของปีการศึกษา {old_yr} ที่อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนแปลงได้')
                if c2 == c1:
                    raise ValueError('camp2_code ต้องไม่ซ้ำกับ camp1_code')
                dup = User.objects.filter(camp2_code=c2).exclude(pk=user.pk).first() or \
                    User.objects.filter(camp1_code=c2).exclude(pk=user.pk).first()
                if dup:
                    raise ValueError(f'camp2_code {c2} ถูกใช้แล้วโดย {dup.camp1_code}')
                user.camp2_code = c2
                user.save(update_fields=['camp2_code', 'updated_at'])
                updated += 1
            except ValueError as e:
                errors.append((rownum, str(e)))
    return {'updated': updated, 'errors': errors, 'total': len(rows)}


# =============================================================================
# Export / Backup
# =============================================================================
USE_RECORD_HEADERS = ['ลำดับ', 'รหัสใช้สิทธิ์', 'สถานะ', 'รหัสค่าย 1', 'รหัสค่ายที่ใช้', 'ชื่อ-นามสกุล',
                      'โรงเรียน', 'ระดับชั้น', 'ปีการศึกษา', 'ค่าย', 'วันเวลายืนยัน']


def use_record_rows(qs):
    rows = []
    for i, r in enumerate(qs, start=1):
        rows.append([
            i, r.rights_code or '-', r.status_text, r.user.camp1_code, r.camp_code_snapshot,
            r.full_name_snapshot or r.user.get_full_name(), r.school_snapshot,
            r.grade_level, r.academic_year, r.camp_level, r.confirmed_at or '',
        ])
    return rows


def use_records_queryset(academic_year, camp_level):
    return UseRecord.objects.filter(academic_year=academic_year, camp_level=camp_level) \
        .select_related('user').order_by('-is_used', 'rights_code', 'user__camp1_code')


def backup_use_records(academic_year, camp_level):
    """Export USE_RECORD เป็น Excel ที่ backups/ใช้สิทธิ์ค่ายX_ปีการศึกษา25XX.xlsx"""
    backup_dir = settings.BACKUP_DIR
    backup_dir.mkdir(parents=True, exist_ok=True)
    qs = use_records_queryset(academic_year, camp_level)
    wb = build_workbook([(f'ค่าย{camp_level}_{academic_year}', USE_RECORD_HEADERS, use_record_rows(qs))])
    path = backup_dir / f'ใช้สิทธิ์ค่าย{camp_level}_ปีการศึกษา{academic_year}.xlsx'
    if path.exists():
        # เก็บไฟล์เดิมไว้ด้วย timestamp ไม่ให้ข้อมูลหาย (ใช้ replace ป้องกัน WinError 183 บน Windows)
        stamp = timezone.localtime().strftime('%Y%m%d_%H%M%S')
        target = backup_dir / f'ใช้สิทธิ์ค่าย{camp_level}_ปีการศึกษา{academic_year}_{stamp}.xlsx'
        idx = 1
        while target.exists():
            target = backup_dir / f'ใช้สิทธิ์ค่าย{camp_level}_ปีการศึกษา{academic_year}_{stamp}_{idx}.xlsx'
            idx += 1
        path.replace(target)
    wb.save(path)
    return path


def list_backups():
    backup_dir = settings.BACKUP_DIR
    backup_dir.mkdir(parents=True, exist_ok=True)
    files = []
    for p in sorted(backup_dir.glob('*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True):
        st = p.stat()
        files.append({'name': p.name, 'size': st.st_size,
                      'modified': timezone.make_aware(dt.datetime.fromtimestamp(st.st_mtime))})
    return files


def safe_backup_path(name):
    """ป้องกัน path traversal"""
    backup_dir = settings.BACKUP_DIR.resolve()
    p = (backup_dir / name).resolve()
    if p.parent != backup_dir or p.suffix.lower() != '.xlsx' or not p.exists():
        return None
    return p


def students_workbook(qs):
    headers = ['รหัสค่าย 1', 'รหัสค่าย 2', 'ชื่อ-นามสกุล', 'โรงเรียน', 'วันเกิด', 'ปีการศึกษา',
               'ระดับชั้นปัจจุบัน', '+1', '+2', '+3', '+4', '+5', 'ใช้สิทธิ์แล้ว (ครั้ง)', 'สถานะบัญชี', 'หมายเหตุ']
    rows = []
    for u in qs:
        rows.append([u.camp1_code, u.camp2_code or '', u.get_full_name(), u.school_name,
                     u.birth_date, u.academic_year, u.current_level, u.year_plus_1,
                     u.year_plus_2, u.year_plus_3, u.year_plus_4, u.year_plus_5,
                     getattr(u, 'used_count', None) if hasattr(u, 'used_count') else u.rights_used_count(),
                     'ใช้งาน' if u.is_active else 'ปิด', u.note])
    return build_workbook([('นักเรียน', headers, rows)])


def history_workbook(user):
    headers = ['ครั้งที่', 'ปีการศึกษา', 'ค่าย', 'รหัสใช้สิทธิ์', 'ใช้สิทธิ์เมื่อ']
    rows = [[i, a.academic_year, a.camp_level, a.rights_code, a.used_at]
            for i, a in enumerate(ArchiveRecord.objects.filter(user=user).order_by('used_at'), start=1)]
    year_headers = ['ปีการศึกษา', 'ค่าย', 'สถานะ', 'รหัสใช้สิทธิ์', 'ระดับชั้น', 'ยืนยันเมื่อ']
    year_rows = [[r.academic_year, r.camp_level, r.status_text, r.rights_code or '-', r.grade_level,
                  r.confirmed_at or ''] for r in UseRecord.objects.filter(user=user).order_by('academic_year')]
    info = [['รหัสค่าย 1', user.camp1_code], ['รหัสค่าย 2', user.camp2_code or '-'],
            ['ชื่อ-นามสกุล', user.get_full_name()], ['โรงเรียน', user.school_name],
            ['ใช้สิทธิ์แล้ว', f'{len(rows)} ครั้ง']]
    return build_workbook([('ข้อมูล', ['หัวข้อ', 'ค่า'], info),
                           ('ประวัติใช้สิทธิ์', headers, rows),
                           ('สถานะรายปี', year_headers, year_rows)])
