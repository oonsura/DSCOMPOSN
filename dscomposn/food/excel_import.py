"""
โมดูลสำหรับตรวจสอบความถูกต้อง (Preview) และนำเข้าคะแนนจากไฟล์ Excel (.xlsx):
1. ตรวจสอบคอลัมน์ใน Excel เทียบกับหัวข้อย่อยคะแนนในระบบ (ตรงหรือไม่ ขาดหัวข้อใด มีคอลัมน์ส่วนเกินใด)
   - รองรับไฟล์ที่ "มีแถวหัวตาราง" และ "ไม่มีแถวหัวตาราง" (เริ่มที่รายชื่อนักเรียนเลย)
   - จับคู่ชื่อคอลัมน์แบบตรงตัว (ไม่ใช้การค้นหาบางส่วน เพื่อป้องกัน "ข้อ 1" ไปจับคู่กับ "ข้อ 10")
2. จับคู่นักเรียนด้วย รหัสค่าย 1 (หรือรหัสค่าย 2) หรือ ชื่อ-นามสกุล (พร้อมตัดคำนำหน้าและช่องว่าง)
   - กรณีไม่พบในระบบปีนี้ แจ้งสาเหตุชัดเจน (เช่น อยู่อีกปีและยังไม่ใช้สิทธิ์ หรือแนะนำชื่อใกล้เคียงกรณีสะกดต่างกัน)
3. ตรวจสอบความถูกต้องของคะแนน (ตัวเลข, ไม่ติดลบ, ไม่เกินคะแนนเต็ม)
4. สรุปปัญหาและแสดงตาราง Preview ก่อนกดยืนยันนำเข้า
"""
import difflib
import re
from collections import Counter
from decimal import Decimal, InvalidOperation

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from composn.audit import log_action
from .services import FoodError, save_score, to_decimal

CAMP_CODE_RE = re.compile(r'^\d{2}(DS|KM)\d{2}$', re.IGNORECASE)
NUMBER_RE = re.compile(r'^-?\d+(\.\d+)?$')
THAI_TEXT_RE = re.compile(r'[ก-๙]{2,}')

# หัวคอลัมน์ที่ไม่ใช่คะแนน (ข้ามเมื่อจับคู่คอลัมน์ตามลำดับ)
NON_SCORE_EXACT = {'ที่', 'ลำดับ', 'ลำดับที่', 'no', 'no.', '#', 'เลขที่'}
NON_SCORE_CONTAINS = ['รวม', 'total', 'sum', 'หมายเหตุ', 'โรงเรียน', 'ระดับชั้น', 'ชั้น', 'คำนำหน้า',
                      'เพศ', 'ศูนย์', 'ร้อยละ', '%', 'เกรด', 'grade', 'อันดับ', 'ผลการ', 'remark', 'note']

NAME_PREFIXES = ['เด็กชาย', 'เด็กหญิง', 'ด.ช.', 'ด.ช', 'ด.ญ.', 'ด.ญ', 'นางสาว', 'น.ส.', 'น.ส', 'นาย', 'นาง',
                 'mr.', 'mrs.', 'miss', 'ms.']
BANNER_KEYWORDS = ['รายงาน', 'ตารางคะแนน', 'สรุปคะแนน', 'ผลการสอบ', 'ปีการศึกษา', 'ค่าย', 'ศูนย์อบรม', 'สอวน']


def norm_thai_name(s):
    """ตัดคำนำหน้าชื่อ และตัดช่องว่างทั้งหมดออก เพื่อให้เปรียบเทียบชื่อ-นามสกุลภาษาไทยได้อย่างแม่นยำ"""
    if not s:
        return ''
    s = str(s).strip()
    low = s.lower()
    for p in NAME_PREFIXES:
        if low.startswith(p):
            s = s[len(p):].strip()
            break
    # ตัดช่องว่าง จุด ขีด ออกทั้งหมด
    s = re.sub(r'[\s\.\-]+', '', s)
    return s.lower()


def _starts_with_person_prefix(s):
    low = str(s or '').strip().lower()
    return any(low.startswith(p) for p in NAME_PREFIXES)


def clean_col_header(h):
    """ลบส่วนวงเล็บคะแนนเต็มหรือข้อความขยายออก เช่น 'Lab 1 (เต็ม 10.00)' -> 'Lab 1'"""
    if not h:
        return ''
    s = str(h).strip()
    s = re.sub(r'\s*[\(\[].*?[\)\]]', '', s).strip()
    return s


def _norm_key(s):
    return re.sub(r'[\s_]+', '', str(s or '')).lower()


def _is_number(v):
    return bool(NUMBER_RE.match(str(v or '').strip().replace(',', '')))


def _col_label(col_i, header=''):
    letter = get_column_letter(col_i + 1)
    header = (header or '').strip()
    return f'คอลัมน์ {letter} "{header}"' if header else f'คอลัมน์ {letter}'


def _sub_col_keys(sc):
    """คำที่ใช้จับคู่หัวคอลัมน์กับหัวข้อย่อย (ต้องตรงตัวหลังตัดช่องว่าง)"""
    keys = {_norm_key(sc.name)}
    label = getattr(sc, 'display_label', '')
    if label:
        keys.add(_norm_key(label))
    code = getattr(sc, 'display_code', '')
    if code:
        keys.add(_norm_key(f'{code}{sc.name}'))
        keys.add(_norm_key(f'{code}.{sc.name}'))
    keys.discard('')
    return keys


def _is_non_score_header(h):
    hl = _norm_key(clean_col_header(h))
    if not hl:
        return False
    if hl in NON_SCORE_EXACT:
        return True
    return any(kw in hl for kw in NON_SCORE_CONTAINS)


def _looks_like_banner_row(cells):
    """ตรวจว่าเป็นแถวหัวเรื่อง/ชื่อรายงานด้านบนตาราง (มีข้อความเดียว เช่น 'คะแนนสอบทฤษฎี 1 ปี 2569')"""
    non_empty = [c for c in cells if c]
    if not non_empty or len(non_empty) > 2:
        return False
    if any(_is_number(c) or _starts_with_person_prefix(c) for c in non_empty):
        return False
    joined = ' '.join(non_empty)
    return any(kw in joined for kw in BANNER_KEYWORDS) or len(non_empty) == 1


def preview_score_excel(file_obj, allowed_members, sub_cols, exam_no=1, year=None):
    """
    ตรวจสอบไฟล์ Excel คะแนน คืนค่า dict (บันทึกลง Session ได้ ทุกค่าเป็น JSON):
      success, error_message, matched_cols, missing_sub_cols, ignored_cols, column_issues,
      rows, valid_count, invalid_count, skipped_count, warning_count, total_rows, has_problems,
      has_header, has_resolvable_unmatched, exam_no
    """
    try:
        wb = load_workbook(file_obj, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        all_rows = []
        for excel_rownum, r in enumerate(ws.iter_rows(values_only=True), start=1):
            if r and any(v not in (None, '') and str(v).strip() != '' for v in r):
                cells = []
                for v in r:
                    if v is None:
                        cells.append('')
                    elif isinstance(v, float) and v.is_integer():
                        cells.append(str(int(v)))
                    else:
                        cells.append(str(v).strip())
                all_rows.append((excel_rownum, cells))
        wb.close()
    except Exception as e:
        return {
            'success': False,
            'error_message': f'ไม่สามารถเปิดอ่านไฟล์ Excel ได้ ({str(e)}) กรุณาใช้ไฟล์ .xlsx',
        }

    if not all_rows:
        return {'success': False, 'error_message': 'ไฟล์ Excel ว่างเปล่า ไม่มีข้อมูล'}

    max_cols = max(len(c) for _, c in all_rows)
    all_rows = [(n, c + [''] * (max_cols - len(c))) for n, c in all_rows]

    # 1. สร้าง Map สำหรับค้นหานักเรียนในระบบปีการศึกษานี้
    yy_prefix = f'{int(year) % 100:02d}' if year else ''
    members_list = list(allowed_members)
    user_by_code = {}
    name_owners = {}
    for m in members_list:
        if m.camp1_code:
            user_by_code[m.camp1_code.strip().upper()] = m
        if getattr(m, 'camp2_code', None):
            user_by_code[m.camp2_code.strip().upper()] = m
        for n in {norm_thai_name(m.get_full_name()), norm_thai_name(f'{m.first_name}{m.last_name}')}:
            if n:
                name_owners.setdefault(n, set()).add(m.pk)
    member_by_pk = {m.pk: m for m in members_list}
    user_by_name = {n: member_by_pk[next(iter(pks))] for n, pks in name_owners.items() if len(pks) == 1}
    ambiguous_names = {n for n, pks in name_owners.items() if len(pks) > 1}
    all_norm_names = list(user_by_name.keys())

    # สร้าง Map นักเรียนทั้งหมดในฐานข้อมูล (เพื่อแจ้งสาเหตุได้ชัดเจนเมื่อไม่พบในปีนี้ เช่น ยังไม่ได้ใช้สิทธิ์/ปิดใช้งาน)
    from composn.models import User
    other_by_code = {}
    other_by_name = {}
    allowed_pks = set(member_by_pk.keys())
    for u in User.objects.filter(role=User.ROLE_MEMBER).exclude(pk__in=allowed_pks):
        if u.camp1_code:
            other_by_code[u.camp1_code.strip().upper()] = u
        if u.camp2_code:
            other_by_code[u.camp2_code.strip().upper()] = u
        for n in {norm_thai_name(u.get_full_name()), norm_thai_name(f'{u.first_name}{u.last_name}')}:
            if n and n not in other_by_name:
                other_by_name[n] = u

    def row_has_identity(cells):
        """แถวนี้เป็นข้อมูลนักเรียนหรือไม่ (มีรหัสค่าย หรือมีชื่อที่ตรงกับนักเรียนในระบบ)"""
        for v in cells:
            if CAMP_CODE_RE.match(v.strip()):
                return True
            n = norm_thai_name(v)
            if n and (n in user_by_name or n in ambiguous_names or n in other_by_name):
                return True
        # รองรับกรณีแยกคอลัมน์ชื่อ-นามสกุลติดกันโดยไม่มีหัวตาราง
        for i in range(len(cells) - 1):
            if cells[i] and cells[i + 1] and not _is_number(cells[i]) and not _is_number(cells[i + 1]):
                n2 = norm_thai_name(f'{cells[i]} {cells[i + 1]}')
                if n2 and (n2 in user_by_name or n2 in ambiguous_names or n2 in other_by_name):
                    return True
                if i + 2 < len(cells) and cells[i + 2] and not _is_number(cells[i + 2]):
                    n3 = norm_thai_name(f'{cells[i]} {cells[i + 1]} {cells[i + 2]}')
                    if n3 and (n3 in user_by_name or n3 in ambiguous_names or n3 in other_by_name):
                        return True
        return False

    all_sub_keys = set()
    for sc in sub_cols:
        all_sub_keys |= _sub_col_keys(sc)

    def header_score(cells):
        """ให้คะแนนความเป็นแถวหัวตาราง (ป้องกันแถวชื่อเรื่องที่มีคำว่า 'ชื่อ' เพียงเซลล์เดียว)"""
        if row_has_identity(cells):
            return 0
        non_empty = [v for v in cells if v]
        if not non_empty:
            return 0
        if len(non_empty) == 1 and max_cols > 1:
            return 0
        # ถ้าส่วนใหญ่เป็นตัวเลข (ยกเว้นกรณีหัวคอลัมน์เป็นคำว่า รหัส/ชื่อ นำหน้า) มักเป็นแถวข้อมูลไม่ใช่หัวตาราง
        score = 0
        for v in non_empty:
            vl = v.lower().strip()
            if _is_number(v):
                continue
            if any(kw in vl for kw in ['รหัส', 'code', 'ชื่อ', 'นามสกุล', 'name']):
                score += 3
            elif _norm_key(clean_col_header(v)) in all_sub_keys:
                score += 4
            elif _is_non_score_header(v):
                score += 1
        return score

    # 2. ค้นหาแถวหัวตาราง (Header) ในช่วง 5 แถวแรก หยุดทันทีเมื่อเจอแถวข้อมูลนักเรียน
    header_pos = None
    best_hscore = 0
    for pos, (_, cells) in enumerate(all_rows[:5]):
        if row_has_identity(cells):
            break
        hs = header_score(cells)
        if hs > best_hscore:
            best_hscore = hs
            header_pos = pos
    has_header = header_pos is not None and best_hscore >= 3
    header_cells = all_rows[header_pos][1] if has_header else [''] * max_cols
    data_rows = all_rows[header_pos + 1:] if has_header else all_rows
    # ถ้าไม่มีหัวตาราง ข้ามเฉพาะแถวหัวเรื่อง/ชื่อรายงานด้านบน (ไม่ข้ามแถวนักเรียนที่เว้นว่างคะแนน)
    if not has_header:
        while data_rows and not row_has_identity(data_rows[0][1]) and _looks_like_banner_row(data_rows[0][1]):
            data_rows = data_rows[1:]

    # 3. ระบุคอลัมน์รหัส / ชื่อ / นามสกุล
    code_col_idx = name_col_idx = lname_col_idx = None
    extra_name_cols = []
    if has_header:
        for col_i, h in enumerate(header_cells):
            hl = h.lower().strip()
            if not hl:
                continue
            if code_col_idx is None and ('รหัส' in hl or 'code' in hl or hl in ('id', 'student id')):
                code_col_idx = col_i
            elif ('นามสกุล' in hl and 'ชื่อ' not in hl) or 'surname' in hl or 'last name' in hl:
                if lname_col_idx is None:
                    lname_col_idx = col_i
            elif 'ชื่อ' in hl or 'name' in hl:
                if name_col_idx is None:
                    name_col_idx = col_i

    sample = [c for _, c in data_rows[:30]]
    if code_col_idx is None:
        counts = Counter(i for c in sample for i, v in enumerate(c) if CAMP_CODE_RE.match(v.strip()))
        if counts:
            code_col_idx = counts.most_common(1)[0][0]
    if name_col_idx is None:
        thai_cols = Counter(
            i for c in sample for i, v in enumerate(c)
            if i not in (code_col_idx, lname_col_idx) and THAI_TEXT_RE.search(v) and not _is_number(v)
            and not CAMP_CODE_RE.match(v.strip())
        )
        if thai_cols:
            # ลองเช็กว่ามีคอลัมน์เดี่ยวที่จับคู่ชื่อในระบบได้หรือไม่
            best_single = None
            best_single_hits = 0
            for ci in thai_cols:
                hits = sum(1 for c in sample if norm_thai_name(c[ci]) in user_by_name or norm_thai_name(c[ci]) in other_by_name)
                if hits > best_single_hits:
                    best_single_hits = hits
                    best_single = ci
            if best_single is not None and best_single_hits > 0:
                name_col_idx = best_single
            else:
                # ลองตรวจกรณีแยกคอลัมน์ ชื่อ + นามสกุล (หรือ คำนำหน้า + ชื่อ + นามสกุล) ติดกัน
                sorted_tcols = sorted(thai_cols.keys())
                best_pair = None
                best_pair_hits = 0
                for idx_t in range(len(sorted_tcols) - 1):
                    c1, c2 = sorted_tcols[idx_t], sorted_tcols[idx_t + 1]
                    if c2 == c1 + 1:
                        hits = sum(
                            1 for c in sample
                            if norm_thai_name(f'{c[c1]} {c[c2]}') in user_by_name
                            or norm_thai_name(f'{c[c1]} {c[c2]}') in other_by_name
                        )
                        if hits > best_pair_hits:
                            best_pair_hits = hits
                            best_pair = (c1, c2)
                if best_pair and best_pair_hits > 0:
                    name_col_idx, lname_col_idx = best_pair
                    # ถ้าคอลัมน์ก่อนหน้าเป็นคำนำหน้า ให้ข้ามไว้ใน identity_cols ด้วย
                    if best_pair[0] - 1 in thai_cols:
                        if all(_starts_with_person_prefix(c[best_pair[0] - 1]) or not c[best_pair[0] - 1] for c in sample[:10]):
                            extra_name_cols.append(best_pair[0] - 1)
                else:
                    name_col_idx = thai_cols.most_common(1)[0][0]

    identity_cols = {i for i in (code_col_idx, name_col_idx, lname_col_idx, *extra_name_cols) if i is not None}

    # 4. จับคู่คอลัมน์ Excel -> หัวข้อย่อย
    matched_cols = []
    used_col_indices = set(identity_cols)
    unmatched_sub_cols = list(sub_cols)

    # รอบที่ 1: ชื่อหัวคอลัมน์ตรงกับชื่อหัวข้อย่อยแบบตรงตัว
    if has_header:
        for col_i, h in enumerate(header_cells):
            if col_i in used_col_indices:
                continue
            norm_h = _norm_key(clean_col_header(h))
            if not norm_h:
                continue
            best = next((sc for sc in unmatched_sub_cols if norm_h in _sub_col_keys(sc)), None)
            if best:
                matched_cols.append({
                    'col_idx': col_i,
                    'col_letter': get_column_letter(col_i + 1),
                    'header': h,
                    'sub_col_pk': best.pk,
                    'sub_col_name': getattr(best, 'display_label', best.name),
                    'max_score': float(best.default_max_score),
                    'method': 'name',
                })
                used_col_indices.add(col_i)
                unmatched_sub_cols.remove(best)

    # รอบที่ 2: คอลัมน์คะแนนที่เหลือ จับคู่ตามลำดับ (ข้ามคอลัมน์ที่เป็นข้อความ/ลำดับที่/รวม)
    def col_is_text_only(col_i):
        vals = [c[col_i] for c in sample if col_i < len(c) and c[col_i] != '']
        if not vals:
            return False
        return sum(1 for v in vals if not _is_number(v)) > len(vals) / 2

    def col_is_running_number(col_i):
        """คอลัมน์ลำดับที่ 1,2,3,... ทางซ้ายของรหัส/ชื่อ ไม่ใช่คะแนน"""
        if not identity_cols or col_i >= min(identity_cols):
            return False
        vals = [c[col_i] for c in sample if col_i < len(c) and c[col_i] != '']
        if len(vals) < 2 or not all(_is_number(v) for v in vals):
            return False
        try:
            nums = [int(Decimal(v)) for v in vals]
        except (InvalidOperation, ValueError):
            return False
        return nums == list(range(nums[0], nums[0] + len(nums)))

    candidate_cols = [
        col_i for col_i in range(max_cols)
        if col_i not in used_col_indices
        and not _is_non_score_header(header_cells[col_i])
        and not col_is_text_only(col_i)
        and not col_is_running_number(col_i)
    ]
    # ตัดคอลัมน์ว่างท้ายตาราง
    while candidate_cols and all(c[candidate_cols[-1]] == '' for _, c in data_rows) \
            and not header_cells[candidate_cols[-1]]:
        candidate_cols.pop()

    if unmatched_sub_cols:
        for sc in list(unmatched_sub_cols):
            if not candidate_cols:
                break
            col_i = candidate_cols.pop(0)
            matched_cols.append({
                'col_idx': col_i,
                'col_letter': get_column_letter(col_i + 1),
                'header': header_cells[col_i] or _col_label(col_i),
                'sub_col_pk': sc.pk,
                'sub_col_name': getattr(sc, 'display_label', sc.name),
                'max_score': float(sc.default_max_score),
                'method': 'order',
            })
            used_col_indices.add(col_i)
            unmatched_sub_cols.remove(sc)
    extra_score_cols = [c for c in candidate_cols
                        if any(_is_number(r[c]) for _, r in data_rows)]

    sub_col_order = {sc.pk: i for i, sc in enumerate(sub_cols)}
    sub_col_by_pk = {sc.pk: sc for sc in sub_cols}
    matched_cols.sort(key=lambda x: sub_col_order.get(x['sub_col_pk'], 999))

    # รายงานปัญหาของคอลัมน์
    missing_sub_cols = [{'pk': sc.pk, 'name': getattr(sc, 'display_label', sc.name)} for sc in unmatched_sub_cols]
    column_issues = []
    if not has_header:
        column_issues.append('ไฟล์นี้ไม่มีแถวหัวตาราง ระบบจึงจับคู่คอลัมน์คะแนนกับหัวข้อย่อยตามลำดับจากซ้ายไปขวา '
                             'กรุณาตรวจสอบว่าลำดับคอลัมน์ตรงกับหัวข้อย่อยในระบบ')
    elif any(mc['method'] == 'order' for mc in matched_cols):
        names = ', '.join(f'{mc["col_letter"]} ➔ {mc["sub_col_name"]}' for mc in matched_cols if mc['method'] == 'order')
        column_issues.append(f'ชื่อหัวคอลัมน์ไม่ตรงกับหัวข้อย่อย ระบบจับคู่ตามลำดับ: {names}')
    if code_col_idx is None and name_col_idx is None:
        column_issues.append('ไม่พบคอลัมน์รหัสค่าย 1 หรือชื่อ-นามสกุล จึงไม่สามารถระบุตัวนักเรียนได้')
    if missing_sub_cols:
        column_issues.append(f'ไฟล์มีคอลัมน์คะแนนน้อยกว่าหัวข้อย่อย ไม่พบคอลัมน์สำหรับ: '
                             f'{", ".join(sc["name"] for sc in missing_sub_cols)}')
    if extra_score_cols:
        column_issues.append(
            f'ไฟล์มีคอลัมน์คะแนนมากกว่าหัวข้อย่อยในระบบ {len(extra_score_cols)} คอลัมน์ '
            f'({", ".join(get_column_letter(c + 1) for c in extra_score_cols)}) ซึ่งจะไม่ถูกนำเข้า '
            f'— อาจเลือกหัวข้อคะแนนไม่ตรงกับไฟล์ กรุณาตรวจสอบ')

    ignored_cols = [{'col_idx': i, 'header': _col_label(i, h)}
                    for i, h in enumerate(header_cells)
                    if i not in used_col_indices and (h or any(r[i] != '' for _, r in data_rows))]

    name_col_desc = _col_label(name_col_idx, header_cells[name_col_idx]) if name_col_idx is not None else ''
    if name_col_desc and lname_col_idx is not None:
        name_col_desc += ' + ' + _col_label(lname_col_idx, header_cells[lname_col_idx])
    identity_info = {
        'code_col': _col_label(code_col_idx, header_cells[code_col_idx]) if code_col_idx is not None else '',
        'name_col': name_col_desc,
    }

    # 5. ประมวลผลแถวนักเรียนแต่ละแถว
    preview_rows = []
    valid_count = invalid_count = skipped_count = warning_count = 0
    has_resolvable_unmatched = False
    seen_users = {}

    for row_idx, r in data_rows:
        raw_code = r[code_col_idx].strip() if code_col_idx is not None else ''
        raw_name = r[name_col_idx].strip() if name_col_idx is not None else ''
        if lname_col_idx is not None and r[lname_col_idx].strip():
            raw_name = f'{raw_name} {r[lname_col_idx].strip()}'.strip()

        matched_user = None
        match_by = None
        suggestions = []
        code_clean = raw_code.upper()
        norm_n = norm_thai_name(raw_name)
        if code_clean and code_clean in user_by_code:
            matched_user = user_by_code[code_clean]
            match_by = 'code'
        elif norm_n and norm_n in user_by_name:
            matched_user = user_by_name[norm_n]
            match_by = 'name'

        row_errors = []
        row_warnings = []

        if not matched_user:
            other_u = (other_by_code.get(code_clean) if code_clean else None) or (other_by_name.get(norm_n) if norm_n else None)
            if norm_n in ambiguous_names:
                row_errors.append(f'พบชื่อ "{raw_name}" ซ้ำกันมากกว่า 1 คนในระบบ กรุณาระบุรหัสค่าย 1 ในไฟล์')
            elif other_u is not None:
                if not other_u.is_active:
                    row_errors.append(f'ไม่พบนักเรียนที่ใช้งานได้ — พบ "{other_u.camp1_code} {other_u.get_full_name()}" แต่บัญชีถูกปิดใช้งาน')
                else:
                    row_errors.append(
                        f'ไม่พบนักเรียนในปีการศึกษา {year or "นี้"} — พบ "{other_u.camp1_code} {other_u.get_full_name()}" '
                        f'(ปี {other_u.academic_year}) แต่ยังไม่ได้ยืนยันใช้สิทธิ์ในปีการศึกษา {year or "นี้"}'
                    )
            else:
                close_keys = difflib.get_close_matches(norm_n, all_norm_names, n=3, cutoff=0.78) if norm_n else []
                for ck in close_keys:
                    cu = user_by_name[ck]
                    suggestions.append({
                        'pk': cu.pk,
                        'code': cu.camp1_code,
                        'name': cu.get_full_name(),
                    })
                if suggestions:
                    s_txt = ', '.join(f'{s["code"]} {s["name"]}' for s in suggestions)
                    row_errors.append(
                        f'ไม่พบนักเรียนจากรหัส "{raw_code or "-"}" หรือชื่อ "{raw_name or "-"}" ตรงกัน — '
                        f'ชื่อใกล้เคียงในระบบ: {s_txt} (เลือกจับคู่ด้านซ้ายได้)'
                    )
                else:
                    row_errors.append(f'ไม่พบนักเรียนจากรหัส "{raw_code or "-"}" หรือชื่อ "{raw_name or "-"}" ในระบบปีการศึกษานี้')
        else:
            if matched_user.pk in seen_users:
                row_warnings.append(f'พบรายชื่อซ้ำในไฟล์ (ซ้ำกับแถวที่ {seen_users[matched_user.pk]}) ข้อมูลแถวนี้จะแทนที่แถวก่อนหน้า')
            seen_users[matched_user.pk] = row_idx
            if match_by == 'name' and raw_code and raw_code.upper() != (matched_user.camp1_code or '').upper():
                row_warnings.append(f'จับคู่ด้วยชื่อ (รหัสในไฟล์ "{raw_code}" ไม่ตรงกับรหัสในระบบ "{matched_user.camp1_code}")')

        row_scores = {}
        row_scores_data = {}
        score_errors = []
        has_any_score = False
        for mc in matched_cols:
            sc = sub_col_by_pk[mc['sub_col_pk']]
            cell_val = r[mc['col_idx']].strip()
            sc_info = {
                'sub_col_pk': sc.pk,
                'sub_col_name': mc['sub_col_name'],
                'max_score': float(sc.default_max_score),
                'raw_value': cell_val,
                'value': None,
                'status': 'blank',
                'error': None,
            }
            if cell_val != '':
                has_any_score = True
                try:
                    sc_val = to_decimal(cell_val)
                    if sc_val > sc.default_max_score:
                        sc_info['status'] = 'exceeded'
                        sc_info['error'] = f'คะแนน ({sc_val}) เกินคะแนนเต็ม ({sc.default_max_score})'
                        score_errors.append(f'{mc["sub_col_name"]}: คะแนน {sc_val} เกินเต็ม {sc.default_max_score}')
                    else:
                        sc_info['status'] = 'valid'
                        sc_info['value'] = str(sc_val)
                        row_scores_data[str(sc.pk)] = str(sc_val)
                except FoodError as fe:
                    sc_info['status'] = 'error'
                    sc_info['error'] = str(fe)
                    score_errors.append(f'{mc["sub_col_name"]}: {str(fe)}')
            row_scores[str(sc.pk)] = sc_info

        row_errors.extend(score_errors)

        is_skipped = matched_user is not None and not has_any_score and not row_errors
        if is_skipped:
            row_warnings.append('ไม่มีคะแนนในแถวนี้ (ข้าม คะแนนเดิมในระบบคงเดิม)')

        can_override = matched_user is None and not score_errors and has_any_score
        if can_override:
            has_resolvable_unmatched = True

        is_valid = matched_user is not None and not row_errors and has_any_score
        if is_valid:
            valid_count += 1
            if row_warnings:
                warning_count += 1
        elif is_skipped:
            skipped_count += 1
        else:
            invalid_count += 1

        user_info = None
        if matched_user:
            code = matched_user.camp1_code or ''
            user_info = {
                'pk': matched_user.pk,
                'camp1_code': code,
                'name': matched_user.get_full_name(),
                'is_current_cohort': bool(yy_prefix and code.upper().startswith(yy_prefix))
                or (year is not None and getattr(matched_user, 'academic_year', None) == int(year)),
            }

        preview_rows.append({
            'rownum': row_idx,
            'raw_code': raw_code,
            'raw_name': raw_name,
            'user': user_info,
            'user_pk': matched_user.pk if matched_user else None,
            'user_code': matched_user.camp1_code if matched_user else raw_code,
            'user_name': matched_user.get_full_name() if matched_user else raw_name,
            'match_by': match_by,
            'suggestions': suggestions,
            'can_override': can_override,
            'has_score_errors': bool(score_errors),
            'scores': row_scores,
            'scores_to_save': row_scores_data,
            'is_valid': is_valid,
            'is_skipped': is_skipped,
            'errors': row_errors,
            'warnings': row_warnings,
        })

    has_problems = invalid_count > 0 or bool(missing_sub_cols) or bool(extra_score_cols) or \
        (code_col_idx is None and name_col_idx is None)

    return {
        'success': True,
        'has_header': has_header,
        'identity_info': identity_info,
        'matched_cols': matched_cols,
        'missing_sub_cols': missing_sub_cols,
        'ignored_cols': ignored_cols,
        'column_issues': column_issues,
        'rows': preview_rows,
        'valid_count': valid_count,
        'invalid_count': invalid_count,
        'skipped_count': skipped_count,
        'warning_count': warning_count,
        'total_rows': len(preview_rows),
        'has_problems': has_problems,
        'has_resolvable_unmatched': has_resolvable_unmatched,
        'exam_no': exam_no,
    }


def execute_score_import(preview_data, admin_user, exam_no, selected_sub_cols, request=None,
                         override_users=None, allowed_user_ids=None):
    """
    บันทึกคะแนนจริงจากผลการ Preview:
    นำเข้าเฉพาะแถวที่ผ่านการตรวจสอบ (is_valid=True)
    หรือแถวที่ผู้ใช้เลือกจับคู่นักเรียนเองจากตาราง Preview (override_users: {rownum: user_pk})
    """
    saved_count = 0
    errors = []
    sub_col_map = {sc.pk: sc for sc in selected_sub_cols}
    override_users = override_users or {}

    from composn.models import User
    user_cache = {}

    for r in preview_data.get('rows', []):
        uid = None
        if r.get('is_valid') and r.get('user_pk'):
            uid = r['user_pk']
        elif r.get('can_override') and not r.get('has_score_errors') and r.get('scores_to_save'):
            chosen = override_users.get(str(r.get('rownum'))) or override_users.get(r.get('rownum'))
            if chosen:
                try:
                    chosen_int = int(chosen)
                except (TypeError, ValueError):
                    chosen_int = None
                if chosen_int and (allowed_user_ids is None or chosen_int in allowed_user_ids):
                    uid = chosen_int
        if not uid:
            continue

        u = user_cache.get(uid) or User.objects.filter(pk=uid).first()
        if not u:
            continue
        user_cache[uid] = u

        for sc_pk, sc_str in (r.get('scores_to_save') or {}).items():
            sc = sub_col_map.get(int(sc_pk))
            if not sc:
                continue
            try:
                sc_val = Decimal(sc_str)
                save_score(u, sc, exam_no, sc_val, sc.default_max_score,
                           recorded_by=admin_user, request=request)
                saved_count += 1
            except (InvalidOperation, FoodError) as e:
                errors.append(f'{u.camp1_code} {sc.name}: {str(e)}')

    log_action(admin_user, 'import_scores', target=admin_user,
               new_value={'saved_count': saved_count, 'exam_no': exam_no}, request=request)

    return saved_count, errors
