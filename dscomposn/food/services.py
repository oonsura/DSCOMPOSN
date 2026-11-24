"""
Business logic ของแอป Food (เวลาเรียน / คะแนน)
"""
from decimal import Decimal, InvalidOperation

from django.db import transaction

from composn.audit import log_action

from .models import (
    Attendance, Score, ScoreApproval, TrainingDate, absence_status, current_year,
    is_future_attendance_date, year_locked_message,
)


class FoodError(Exception):
    pass


def _resolve_attendance_lock_years(student, date, period, academic_year=None):
    """รวบรวมปีการศึกษาที่เกี่ยวข้องกับรายการเวลาเรียน เพื่อตรวจสอบสถานะอนุมัติประกาศผล"""
    years = []
    if academic_year:
        try:
            years.append(int(academic_year))
        except (TypeError, ValueError):
            pass
    existing = Attendance.objects.filter(user=student, date=date, session_period=period).first()
    if existing and existing.academic_year not in years:
        years.append(existing.academic_year)
    if not years:
        td = TrainingDate.objects.filter(date=date).first()
        if td:
            years.append(td.academic_year)
        else:
            years.append(current_year())
    return years


def save_attendance(student, date, period, status, reason='', note='', recorded_by=None, request=None,
                    academic_year=None):
    """บันทึก/แก้ไขเวลาเรียน 1 รายการ (ไม่ซ้ำ user+date+period) คืนค่า (attendance, absence_status)"""
    if is_future_attendance_date(date):
        raise FoodError('ยังไม่ถึงวันที่ดังกล่าว ไม่สามารถบันทึกเวลาเรียนล่วงหน้าได้')
    if period not in dict(Attendance.PERIOD_CHOICES):
        raise FoodError('ช่วงเวลาไม่ถูกต้อง')
    if status not in Attendance.ALL_VALID_STATUSES:
        raise FoodError('สถานะไม่ถูกต้อง')
    for chk_year in _resolve_attendance_lock_years(student, date, period, academic_year=academic_year):
        if ScoreApproval.is_year_approved(chk_year):
            raise FoodError(year_locked_message(chk_year))
    with transaction.atomic():
        att = Attendance.objects.select_for_update().filter(user=student, date=date, session_period=period).first()
        old = None
        if att:
            old = {'status': att.status, 'reason': att.reason, 'note': att.note}
            eff_reason = '' if status == Attendance.STATUS_PRESENT else (att.reason if reason is None else (reason or ''))[:100]
            eff_note = att.note if note is None else (note or '')
            if old == {'status': status, 'reason': eff_reason, 'note': eff_note}:
                return att, absence_status(student, att.academic_year)
            att.status, att.reason, att.note = status, eff_reason, eff_note
            att.recorded_by = recorded_by
            att.save()
        else:
            eff_reason = '' if status == Attendance.STATUS_PRESENT else (reason or '')[:100]
            eff_note = note or ''
            kwargs = {}
            if academic_year:
                kwargs['academic_year'] = academic_year
            att = Attendance.objects.create(user=student, date=date, session_period=period, status=status,
                                            reason=eff_reason, note=eff_note, recorded_by=recorded_by, **kwargs)
    log_action(recorded_by, 'save_attendance', target=att, old_value=old, new_value={
        'student': student.camp1_code, 'date': str(date), 'period': period, 'status': status, 'reason': eff_reason,
        'note': eff_note}, request=request)
    return att, absence_status(student, att.academic_year)


def delete_attendance(att, by=None, request=None):
    if att and ScoreApproval.is_year_approved(att.academic_year):
        raise FoodError(year_locked_message(att.academic_year))
    att.delete(user=by)
    log_action(by, 'delete_attendance', target=att, old_value={
        'student': att.user.camp1_code, 'date': str(att.date), 'period': att.session_period,
        'status': att.status}, request=request)
    return absence_status(att.user, att.academic_year)


def to_decimal(value):
    if value is None:
        return None
    s = str(value).strip().replace(',', '')
    if s == '':
        return None
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise FoodError(f'คะแนน "{value}" ไม่ใช่ตัวเลข')
    if not d.is_finite():
        raise FoodError(f'คะแนน "{value}" ไม่ใช่ตัวเลข')
    if d < 0:
        raise FoodError('คะแนนต้องไม่ติดลบ')
    return d


def save_score(student, category, exam_no, score, max_score, recorded_by=None, request=None):
    """บันทึกคะแนน: score=None -> ลบ (soft delete) รายการเดิม"""
    if ScoreApproval.is_year_approved(category.academic_year):
        raise FoodError(year_locked_message(category.academic_year))
    existing = Score.objects.filter(user=student, category=category, exam_no=exam_no).first()
    if score is None:
        if existing:
            existing.delete(user=recorded_by)
            log_action(recorded_by, 'delete_score', target=existing, old_value={
                'student': student.camp1_code, 'category': category.name, 'exam_no': exam_no,
                'score': str(existing.score)}, request=request)
        return None
    if not max_score or max_score <= 0:
        raise FoodError('คะแนนเต็มต้องมากกว่า 0')
    if score > max_score:
        raise FoodError(f'{student.camp1_code}: คะแนน ({score}) มากกว่าคะแนนเต็ม ({max_score})')
    if existing:
        if existing.score == score and existing.max_score == max_score:
            return existing
        old = {'score': str(existing.score), 'max_score': str(existing.max_score)}
        existing.score, existing.max_score, existing.recorded_by = score, max_score, recorded_by
        existing.save()
        obj = existing
    else:
        old = None
        obj = Score.objects.create(user=student, category=category, exam_no=exam_no, score=score,
                                   max_score=max_score, recorded_by=recorded_by)
    log_action(recorded_by, 'save_score', target=obj, old_value=old, new_value={
        'student': student.camp1_code, 'category': category.name, 'exam_no': exam_no, 'score': str(score),
        'max_score': str(max_score)}, request=request)
    return obj
