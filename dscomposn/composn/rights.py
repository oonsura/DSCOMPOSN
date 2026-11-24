"""
ตรรกะระบบสิทธิ์ COMPOSN

กฎ:
- มีรหัสค่าย 1 อย่างเดียว ใช้สิทธิ์ค่าย 1 ได้
- มีรหัสค่าย 2 ใช้สิทธิ์ค่าย 2 ได้ และไม่สามารถใช้สิทธิ์ค่าย 1 ได้ทุกกรณี
- ใช้สิทธิ์ได้สูงสุด 5 ครั้งตลอดชีวิต (นับจากจำนวนแถวใน ARCHIVE_RECORD)
- ระดับชั้นในปีนั้นต้องไม่เกิน ม.6 / ปวช.3
- เปิดรอบแล้วไม่ยืนยัน = สละสิทธิ์ในปีนั้น (บันทึกตอนปิดรอบ)
- รหัสใช้สิทธิ์: ค่าย 1 ใช้รหัสค่าย 1 (camp1_code) / ค่าย 2 ใช้รหัสค่าย 2 (camp2_code)
"""
from dataclasses import dataclass, field

from django.db import IntegrityError, transaction
from django.db.models import Q
from django.utils import timezone

from . import levels
from .audit import log_action
from .models import (
    EXPIRED_TEXT, ArchiveRecord, RightsRound, SystemSetting, UseRecord, User,
)


class RightsError(Exception):
    pass


@dataclass
class Eligibility:
    ok: bool
    message: str
    rights_round: RightsRound = None
    grade_level: str = ''
    used_count: int = 0
    max_rights: int = 5
    code: str = ''  # รหัสสถานะสั้นสำหรับ UI/ทดสอบ
    extra: dict = field(default_factory=dict)


def _is_year_approved(year):
    from food.models import ScoreApproval
    return ScoreApproval.is_year_approved(year)


def get_open_round(camp_level):
    return RightsRound.objects.filter(is_open=True, camp_level=camp_level).order_by('-opened_at').first()


def is_current_camp_student(user, year=None):
    """นักเรียนปัจจุบันในปีการศึกษาที่ระบุ (เข้าค่ายปีเดียวกับหรือมากกว่าปีนั้น — ไม่ใช่ผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบัน)"""
    if not user or user.role != User.ROLE_MEMBER:
        return False
    setting_year = int(year or SystemSetting.get().current_academic_year)
    base_y = user.base_academic_year
    return bool(base_y and int(base_y) >= setting_year)


def year_record(user, academic_year):
    """USE_RECORD ของผู้ใช้ในปีนั้น (ไม่รวมแถวที่ถูกลบ)"""
    return UseRecord.objects.filter(user=user, academic_year=academic_year).order_by('-id').first()


def check_eligibility(user, rights_round=None, check_schedule=True):
    """ตรวจสอบว่าผู้ใช้ยืนยันใช้สิทธิ์ได้หรือไม่ (ใช้สิทธิ์ได้เฉพาะผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบันเท่านั้น)"""
    setting = SystemSetting.get()
    max_rights = setting.max_rights
    used = user.rights_used_count()
    base = dict(used_count=used, max_rights=max_rights)

    if not user.is_active:
        return Eligibility(False, 'บัญชีนี้ถูกปิดการใช้งาน', code='inactive', **base)
    if user.role != User.ROLE_MEMBER:
        return Eligibility(False, 'เฉพาะนักเรียน/ผู้ใช้สิทธิ์เท่านั้นที่ยืนยันใช้สิทธิ์ได้', code='not_member', **base)

    target_year = rights_round.academic_year if rights_round is not None else setting.current_academic_year
    if is_current_camp_student(user, target_year):
        return Eligibility(
            False,
            f'นักเรียนปัจจุบันประจำปีการศึกษา {target_year} ไม่ต้องยืนยันการใช้สิทธิ์ '
            f'(การใช้สิทธิ์สำหรับผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบันเท่านั้น)',
            rights_round=rights_round,
            code='current_student',
            **base,
        )

    if rights_round is None:
        rights_round = get_open_round(user.camp_level)
        if rights_round is None:
            other = get_open_round(1 if user.camp_level == 2 else 2)
            if other and user.camp_level == 2:
                return Eligibility(False, 'ท่านมีรหัสค่าย 2 จึงไม่สามารถใช้สิทธิ์ค่าย 1 ได้ทุกกรณี',
                                   rights_round=other, code='camp2_cannot_use_camp1', **base)
            if other and user.camp_level == 1:
                return Eligibility(False, 'ท่านยังไม่มีรหัสค่าย 2 จึงไม่สามารถใช้สิทธิ์ค่าย 2 ได้',
                                   rights_round=other, code='no_camp2', **base)
            return Eligibility(False, 'ขณะนี้ยังไม่เปิดรอบใช้สิทธิ์', code='no_round', **base)
        if is_current_camp_student(user, rights_round.academic_year):
            return Eligibility(
                False,
                f'นักเรียนปัจจุบันประจำปีการศึกษา {rights_round.academic_year} ไม่ต้องยืนยันการใช้สิทธิ์ '
                f'(การใช้สิทธิ์สำหรับผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบันเท่านั้น)',
                rights_round=rights_round,
                code='current_student',
                **base,
            )

    if not rights_round.is_open:
        return Eligibility(False, 'รอบใช้สิทธิ์นี้ปิดแล้ว', rights_round=rights_round, code='round_closed', **base)
    if _is_year_approved(rights_round.academic_year):
        return Eligibility(
            False,
            f'ปีการศึกษา {rights_round.academic_year} อนุมัติประกาศผลแล้ว ไม่สามารถใช้สิทธิ์ได้',
            rights_round=rights_round,
            code='results_approved',
            **base,
        )
    if check_schedule and not rights_round.is_within_schedule():
        return Eligibility(False, 'ขณะนี้อยู่นอกช่วงเวลาเปิด-ปิดการใช้สิทธิ์', rights_round=rights_round,
                           code='outside_schedule', **base)
    if rights_round.camp_level == 1 and user.camp2_code:
        return Eligibility(False, 'ท่านมีรหัสค่าย 2 จึงไม่สามารถใช้สิทธิ์ค่าย 1 ได้ทุกกรณี',
                           rights_round=rights_round, code='camp2_cannot_use_camp1', **base)
    if rights_round.camp_level == 2 and not user.camp2_code:
        return Eligibility(False, 'ท่านยังไม่มีรหัสค่าย 2 จึงไม่สามารถใช้สิทธิ์ค่าย 2 ได้',
                           rights_round=rights_round, code='no_camp2', **base)

    rec = year_record(user, rights_round.academic_year)
    if rec is not None:
        if rec.is_used:
            can_cancel = bool(rights_round.is_open and rights_round.is_within_schedule())
            return Eligibility(
                False,
                f'ท่านใช้สิทธิ์ในปีการศึกษา {rights_round.academic_year} แล้ว (รหัส {rec.rights_code})',
                rights_round=rights_round,
                grade_level=rec.grade_level,
                code='already_used',
                extra={'can_cancel': can_cancel, 'use_record': rec},
                **base,
            )
        return Eligibility(False, f'ท่านสละสิทธิ์ในปีการศึกษา {rights_round.academic_year} แล้ว',
                           rights_round=rights_round, code='waived', **base)

    if used >= max_rights:
        return Eligibility(False, f'ท่านใช้สิทธิ์ครบ {max_rights} ครั้งแล้ว', rights_round=rights_round,
                           code='max_reached', **base)

    grade = user.level_for_year(rights_round.academic_year)
    if grade == EXPIRED_TEXT or (grade and not levels.is_within_limit(grade)):
        return Eligibility(False, 'หมดสิทธิ์ (ระดับชั้นเกิน ม.6 / ปวช.3)', rights_round=rights_round,
                           grade_level=EXPIRED_TEXT, code='expired', **base)
    if not grade:
        return Eligibility(False, f'ไม่พบข้อมูลระดับชั้นในปีการศึกษา {rights_round.academic_year} '
                                  f'กรุณาติดต่อผู้ดูแลระบบ', rights_round=rights_round,
                           code='no_level', **base)

    return Eligibility(True, f'ท่านมีสิทธิ์ใช้สิทธิ์ค่ายที่ {rights_round.camp_level} '
                             f'ปีการศึกษา {rights_round.academic_year}',
                       rights_round=rights_round, grade_level=grade, code='eligible', **base)


def user_rights_code(user, camp_level):
    """รหัสที่ใช้ในการใช้สิทธิ์: ค่าย 1 ใช้รหัสค่าย 1 (camp1_code), ค่าย 2 ใช้รหัสค่าย 2 (camp2_code)"""
    if int(camp_level) == 2:
        if not user.camp2_code:
            raise RightsError('ไม่พบรหัสค่าย 2 ของนักเรียน')
        return user.camp2_code
    return user.camp1_code


def confirm_rights(user, rights_round=None, request=None):
    """ยืนยันใช้สิทธิ์ด้วยรหัสค่าย 1 (สำหรับค่าย 1) หรือรหัสค่าย 2 (สำหรับค่าย 2) คืนค่า UseRecord หรือ raise RightsError"""
    for attempt in range(5):
        try:
            with transaction.atomic():
                if rights_round is not None:
                    rights_round = RightsRound.objects.select_for_update().get(pk=rights_round.pk)
                else:
                    r = get_open_round(user.camp_level)
                    rights_round = RightsRound.objects.select_for_update().get(pk=r.pk) if r else None
                user = User.objects.select_for_update().get(pk=user.pk)
                elig = check_eligibility(user, rights_round)
                if not elig.ok:
                    raise RightsError(elig.message)
                rr = elig.rights_round
                code = user_rights_code(user, rr.camp_level)
                now = timezone.now()
                rec = UseRecord.objects.create(
                    user=user, rights_round=rr, academic_year=rr.academic_year, camp_level=rr.camp_level,
                    rights_code=code, grade_level=elig.grade_level, confirmed_at=now, is_used=True,
                    camp_code_snapshot=code,
                    full_name_snapshot=user.get_full_name(), school_snapshot=user.school_name,
                    center_snapshot=user.camp_center,
                )
                ArchiveRecord.objects.create(
                    user=user, use_record=rec, academic_year=rr.academic_year, camp_level=rr.camp_level,
                    rights_code=code, used_at=now,
                )
        except IntegrityError:
            continue  # รหัสชนกัน ลองใหม่
        log_action(user, 'confirm_rights', target=rec, new_value={
            'rights_code': rec.rights_code, 'academic_year': rec.academic_year,
            'camp_level': rec.camp_level, 'grade_level': rec.grade_level,
        }, request=request)
        return rec
    raise RightsError('ระบบไม่สามารถออกรหัสใช้สิทธิ์ได้ กรุณาลองใหม่อีกครั้ง')


def cancel_rights(user, rights_round=None, request=None):
    """นักเรียนยกเลิกการใช้สิทธิ์ของตนเองได้ทุกเมื่อ ก่อนหมดเวลา (รอบยังเปิดอยู่และอยู่ในช่วงเวลา)"""
    with transaction.atomic():
        user = User.objects.select_for_update().get(pk=user.pk)
        if rights_round is not None:
            rights_round = RightsRound.objects.select_for_update().get(pk=rights_round.pk)
        else:
            r = get_open_round(user.camp_level)
            rights_round = RightsRound.objects.select_for_update().get(pk=r.pk) if r else None

        if rights_round is None or not rights_round.is_open or not rights_round.is_within_schedule():
            raise RightsError('หมดเวลาการใช้สิทธิ์แล้ว หรือรอบการใช้สิทธิ์ปิดแล้ว ไม่สามารถยกเลิกได้')
        if _is_year_approved(rights_round.academic_year):
            raise RightsError(f'ปีการศึกษา {rights_round.academic_year} อนุมัติประกาศผลแล้ว ไม่สามารถยกเลิกการใช้สิทธิ์ได้')

        rec = year_record(user, rights_round.academic_year)
        if rec is None or not rec.is_used:
            raise RightsError('ท่านยังไม่ได้ยืนยันใช้สิทธิ์ในปีการศึกษานี้')

        old_code = rec.rights_code
        old_year = rec.academic_year
        old_camp = rec.camp_level
        ArchiveRecord.objects.filter(use_record=rec).delete(user=user)
        rec.delete(user=user)

    log_action(user, 'cancel_rights', target_model='UseRecord', target_id=rec.pk, old_value={
        'rights_code': old_code, 'academic_year': old_year, 'camp_level': old_camp,
    }, request=request)
    return old_code


def open_round(academic_year, camp_level, admin, title='', request=None,
               scheduled_open_at=None, scheduled_close_at=None):
    academic_year = int(academic_year)
    camp_level = int(camp_level)
    if camp_level not in (1, 2):
        raise RightsError('ค่ายไม่ถูกต้อง')
    if _is_year_approved(academic_year):
        raise RightsError(f'ปีการศึกษา {academic_year} อนุมัติประกาศผลแล้ว ไม่สามารถเปิดรอบใช้สิทธิ์ได้ กรุณายกเลิกประกาศผลก่อน')
    if RightsRound.objects.filter(is_open=True, camp_level=camp_level).exists():
        raise RightsError(f'มีรอบใช้สิทธิ์ค่าย {camp_level} ที่เปิดอยู่แล้ว กรุณาปิดรอบเดิมก่อน')
    rr = RightsRound.objects.create(academic_year=academic_year, camp_level=camp_level, title=title,
                                    is_open=True, opened_by=admin,
                                    scheduled_open_at=scheduled_open_at, scheduled_close_at=scheduled_close_at)
    log_action(admin, 'open_round', target=rr, new_value={
        'academic_year': academic_year, 'camp_level': camp_level, 'title': title}, request=request)
    return rr


def reopen_round(rights_round, admin, request=None):
    """เปิดรอบที่เคยปิดไปแล้วขึ้นมาใหม่ (และล้างรายการสละสิทธิ์ในรอบนั้นเพื่อให้ยืนยันได้)"""
    if rights_round.is_open:
        raise RightsError('รอบนี้เปิดอยู่แล้ว')
    if _is_year_approved(rights_round.academic_year):
        raise RightsError(f'ปีการศึกษา {rights_round.academic_year} อนุมัติประกาศผลแล้ว ไม่สามารถเปิดรอบใช้สิทธิ์ใหม่ได้ กรุณายกเลิกประกาศผลก่อน')
    if RightsRound.objects.filter(is_open=True, camp_level=rights_round.camp_level).exists():
        raise RightsError(f'มีรอบใช้สิทธิ์ค่าย {rights_round.camp_level} อื่นที่เปิดอยู่แล้ว กรุณาปิดรอบนั้นก่อน')
    with transaction.atomic():
        # ลบ (hard delete) เฉพาะแถวที่ถูกทำเครื่องหมายว่าสละสิทธิ์ (is_used=False) ตอนปิดรอบเดิม
        UseRecord.objects.filter(rights_round=rights_round, is_used=False).hard_delete()
        rights_round.is_open = True
        rights_round.closed_at = None
        rights_round.closed_by = None
        rights_round.save(update_fields=['is_open', 'closed_at', 'closed_by'])
    log_action(admin, 'open_round', target=rights_round, new_value={'reopened': True}, request=request)
    return rights_round


def eligible_unconfirmed_users(rights_round):
    """ผู้ใช้ที่มีสิทธิ์ในรอบนี้แต่ยังไม่ยืนยัน (จะถือว่าสละสิทธิ์เมื่อปิดรอบ)"""
    qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True)
    if rights_round.camp_level == 1:
        qs = qs.filter(camp2_code__isnull=True)
    else:
        qs = qs.filter(camp2_code__isnull=False)
    result = []
    for u in qs:
        if check_eligibility(u, rights_round, check_schedule=False).ok:
            result.append(u)
    return result


def close_round(rights_round, admin, mark_waived=True, request=None):
    """ปิดรอบ + บันทึกสละสิทธิ์ + Backup Excel อัตโนมัติ"""
    from .excel import backup_use_records

    if not rights_round.is_open:
        raise RightsError('รอบนี้ปิดไปแล้ว')
    waived = 0
    with transaction.atomic():
        if mark_waived:
            for u in eligible_unconfirmed_users(rights_round):
                UseRecord.objects.create(
                    user=u, rights_round=rights_round, academic_year=rights_round.academic_year,
                    camp_level=rights_round.camp_level, rights_code='', is_used=False,
                    grade_level=u.level_for_year(rights_round.academic_year),
                    camp_code_snapshot=u.camp2_code if rights_round.camp_level == 2 else u.camp1_code,
                    full_name_snapshot=u.get_full_name(), school_snapshot=u.school_name,
                    center_snapshot=u.camp_center,
                )
                waived += 1
        rights_round.is_open = False
        rights_round.closed_at = timezone.now()
        rights_round.closed_by = admin
        rights_round.save()
    path = backup_use_records(rights_round.academic_year, rights_round.camp_level)
    rights_round.backup_file = path.name
    rights_round.save(update_fields=['backup_file'])
    log_action(admin, 'close_round', target=rights_round, new_value={
        'waived': waived, 'backup_file': path.name}, request=request)
    log_action(admin, 'backup', target_model='UseRecord', new_value={'file': path.name}, request=request)
    return waived, path


def member_rights_summary(user):
    """ข้อมูลสรุปสำหรับหน้า Member"""
    setting = SystemSetting.get()
    elig = check_eligibility(user)
    year = elig.rights_round.academic_year if elig.rights_round else setting.current_academic_year
    camp_level = elig.rights_round.camp_level if elig.rights_round else user.camp_level
    grade_now = user.level_for_year(year)
    is_current = is_current_camp_student(user, setting.current_academic_year)

    # นักเรียนที่มีสิทธิ์ในปีการศึกษานั้น (เฉพาะผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบันเท่านั้น)
    has_rights_in_year = False if is_current else bool(
        elig.ok
        or elig.code == 'already_used'
        or (
            user.is_active
            and user.role == User.ROLE_MEMBER
            and user.rights_used_count() < setting.max_rights
            and grade_now
            and grade_now != EXPIRED_TEXT
            and levels.is_within_limit(grade_now)
            and elig.code not in ('current_student', 'camp2_cannot_use_camp1', 'no_camp2', 'waived', 'expired', 'max_reached', 'no_level')
        )
    )
    if is_current:
        status = 'นักเรียนปัจจุบัน'
    elif elig.ok:
        status = f'มีสิทธิ์ใช้สิทธิ์ค่ายที่ {camp_level}'
    elif elig.code in ('expired', 'max_reached'):
        status = 'หมดสิทธิ์'
    elif elig.code == 'already_used':
        status = f'ใช้สิทธิ์ค่ายที่ {camp_level} ปีนี้แล้ว'
    elif elig.code == 'waived':
        status = 'สละสิทธิ์ปีนี้'
    elif user.rights_used_count() >= setting.max_rights or grade_now == EXPIRED_TEXT:
        status = 'หมดสิทธิ์'
    elif has_rights_in_year:
        status = f'มีสิทธิ์ใช้สิทธิ์ค่ายที่ {camp_level} (รอเปิดรอบ)'
    else:
        status = 'รอเปิดรอบ'

    can_cancel = False if is_current else bool(elig.extra.get('can_cancel'))
    history = ArchiveRecord.objects.filter(user=user).order_by('academic_year', 'used_at')
    year_records = UseRecord.objects.filter(user=user).order_by('academic_year')
    return {
        'eligibility': elig,
        'is_current_student': is_current,
        'status': status,
        'camp_level': camp_level,
        'has_rights_in_year': has_rights_in_year,
        'can_cancel': can_cancel,
        'academic_year': year,
        'used_count': elig.used_count,
        'remaining': max(0, setting.max_rights - elig.used_count),
        'max_rights': setting.max_rights,
        'history': history,
        'year_records': year_records,
        'future_levels': user.future_levels(),
    }


def get_active_camp_for_year(academic_year):
    """คืนค่า (camp_level, rights_round) ตามรอบที่เปิดให้ใช้สิทธิ์อยู่ ณ เวลานั้น (หากยังไม่เปิดรอบ ใช้ค่าย 1 เป็นค่าเริ่มต้น)"""
    rr = RightsRound.objects.filter(academic_year=int(academic_year), is_open=True).order_by('-opened_at').first()
    if not rr:
        rr = RightsRound.objects.filter(academic_year=int(academic_year)).order_by('-opened_at').first()
    return (rr.camp_level if rr else 1), rr


def preview_camp1_rights_import(code_items, academic_year=None, camp_level=None):
    """ตรวจสอบรายชื่อรหัสค่าย 1 ว่าสามารถใช้สิทธิ์ในปีการศึกษานั้น (ตามค่ายที่กำลังเปิดใช้สิทธิ์อยู่ ณ เวลานั้น) ได้หรือไม่ อย่างไร ก่อนนำเข้าจริง
    code_items = [(rownum, camp1_code), ...] หรือ [camp1_code, ...]
    """
    setting = SystemSetting.get()
    year = int(academic_year or setting.current_academic_year)
    active_camp, active_rr = get_active_camp_for_year(year)
    target_camp = int(camp_level or active_camp)
    max_rights = setting.max_rights
    rows = []
    valid_count = 0
    invalid_count = 0
    seen = {}

    for idx, item in enumerate(code_items, start=1):
        if isinstance(item, (tuple, list)) and len(item) >= 2:
            rownum, raw_code = item[0], item[1]
        else:
            rownum, raw_code = idx, item
        code = str(raw_code or '').strip().upper()
        if not code:
            continue

        user = User.objects.filter(Q(camp1_code__iexact=code) | Q(camp2_code__iexact=code)).first()
        valid = False
        reason = ''
        grade = ''
        used = 0
        full_name = '-'
        school_name = '-'
        center = '-'

        if not user:
            reason = f'ไม่พบรหัส "{code}" ในฐานข้อมูลนักเรียน'
        elif user.role != User.ROLE_MEMBER:
            full_name = user.get_full_name()
            reason = f'เป็นบัญชี {user.get_role_display()} ไม่ใช่นักเรียน/สมาชิก'
        else:
            code = user.camp1_code
            full_name = user.get_full_name() or '-'
            school_name = user.school_name or '-'
            center = user.camp_center or '-'
            used = user.rights_used_count()
            grade = user.level_for_year(year)

            if code in seen:
                reason = f'รหัสค่าย 1 ซ้ำกับแถวที่ {seen[code]} ในชุดข้อมูลที่นำเข้า'
            elif not user.is_active:
                reason = 'บัญชีนี้ถูกปิดการใช้งาน (Inactive)'
            elif is_current_camp_student(user, year):
                reason = f'เป็นนักเรียนปัจจุบันประจำปีการศึกษา {year} ไม่ต้องใช้สิทธิ์ (ใช้สิทธิ์ได้เฉพาะผู้ที่เคยเข้าค่ายมาก่อนปีปัจจุบันเท่านั้น)'
            elif target_camp == 1 and user.camp2_code:
                reason = f'มีรหัสค่าย 2 ({user.camp2_code}) แล้ว จึงไม่สามารถใช้สิทธิ์ค่ายที่ 1 ได้ทุกกรณี'
            elif target_camp == 2 and not user.camp2_code:
                reason = 'ยังไม่มีรหัสค่าย 2 จึงไม่สามารถใช้สิทธิ์ค่ายที่ 2 ได้'
            else:
                rec = year_record(user, year)
                if rec and rec.is_used:
                    reason = f'ใช้สิทธิ์ในปีการศึกษา {year} ไปแล้ว (รหัสใช้สิทธิ์ {rec.rights_code})'
                elif used >= max_rights:
                    reason = f'ใช้สิทธิ์ครบจำนวนสูงสุด {max_rights} ครั้งตลอดชีวิตแล้ว'
                elif grade == EXPIRED_TEXT or (grade and not levels.is_within_limit(grade)):
                    reason = f'หมดสิทธิ์ในปีการศึกษา {year} เนื่องจากระดับชั้นเกิน ม.6 / ปวช.3'
                elif not grade:
                    reason = f'ไม่พบข้อมูลระดับชั้นในปีการศึกษา {year} (ปีเริ่มต้น {user.base_academic_year or "-"} ชั้น {user.current_level or "-"})'
                else:
                    valid = True
                    extra_note = ' (เปลี่ยนจากสถานะสละสิทธิ์เป็นใช้สิทธิ์)' if (rec and not rec.is_used) else ''
                    reason = f'สามารถใช้สิทธิ์ค่ายที่ {target_camp} ปีการศึกษา {year} ได้ (ระดับชั้น {grade}){extra_note}'
            seen[code] = rownum

        if valid:
            valid_count += 1
        else:
            invalid_count += 1

        rows.append({
            'rownum': rownum,
            'camp1_code': code,
            'full_name': full_name,
            'school_name': school_name,
            'center': center,
            'grade_level': grade or '-',
            'used_count': used,
            'max_rights': max_rights,
            'valid': valid,
            'reason': reason,
        })

    return {
        'academic_year': year,
        'camp_level': target_camp,
        'rights_round': active_rr,
        'rows': rows,
        'total': len(rows),
        'valid_count': valid_count,
        'invalid_count': invalid_count,
    }


def import_camp1_rights(codes, admin, academic_year=None, camp_level=None, request=None):
    """บันทึกการใช้สิทธิ์ในปีการศึกษานั้น (ตามค่ายที่เปิดให้ใช้สิทธิ์อยู่ ณ เวลานั้น) ให้แก่รายการรหัสค่าย 1 ที่ผ่านการตรวจสอบแล้ว"""
    setting = SystemSetting.get()
    year = int(academic_year or setting.current_academic_year)
    if _is_year_approved(year):
        raise RightsError(f'ปีการศึกษา {year} อนุมัติประกาศผลแล้ว ไม่สามารถนำเข้าผู้ใช้สิทธิ์ได้ กรุณายกเลิกประกาศผลก่อน')
    preview = preview_camp1_rights_import(codes, academic_year=year, camp_level=camp_level)
    target_camp = preview['camp_level']
    rr = preview['rights_round']
    imported = 0
    errors = []
    now = timezone.now()

    with transaction.atomic():
        for r in preview['rows']:
            if not r['valid']:
                errors.append((r['rownum'], f"{r['camp1_code']}: {r['reason']}"))
                continue
            user = User.objects.select_for_update().get(camp1_code=r['camp1_code'])
            # หากมีรายการสละสิทธิ์ (is_used=False) ในปีนั้นอยู่ก่อน ให้ลบออกก่อนบันทึกใช้สิทธิ์
            UseRecord.objects.filter(user=user, academic_year=year, is_used=False).hard_delete()
            code = user_rights_code(user, target_camp)
            rec = UseRecord.objects.create(
                user=user,
                rights_round=rr,
                academic_year=year,
                camp_level=target_camp,
                rights_code=code,
                grade_level=r['grade_level'],
                confirmed_at=now,
                is_used=True,
                camp_code_snapshot=code,
                full_name_snapshot=user.get_full_name(),
                school_snapshot=user.school_name,
                center_snapshot=user.camp_center,
            )
            ArchiveRecord.objects.create(
                user=user,
                use_record=rec,
                academic_year=year,
                camp_level=target_camp,
                rights_code=code,
                used_at=now,
            )
            imported += 1

    log_action(admin, 'import_rights', target_model='UseRecord', new_value={
        'academic_year': year, 'camp_level': target_camp, 'imported': imported, 'errors': len(errors),
    }, request=request)
    return {'imported': imported, 'errors': errors, 'total': preview['total'], 'academic_year': year,
            'camp_level': target_camp}
