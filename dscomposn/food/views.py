"""
Views ของแอป Food: เวลาเรียน / คะแนน / ตั้งค่ากำหนดคะแนน-สัดส่วน-ผู้ผ่านการคัดเลือก-วันอบรม-วันสอบ /
เอกสารประกาศผล / แบบลงชื่อเข้าอบรมและวันสอบ
"""
import datetime as dt
import re
from decimal import Decimal
from urllib.parse import quote, urlencode

from django.contrib import messages
from django.db import transaction
from django.db.models import Case, Count, IntegerField, Q, Value, When
from django.http import HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.clickjacking import xframe_options_sameorigin
from django.views.decorators.http import require_POST
from openpyxl import load_workbook

from composn.audit import log_action
from composn.context_processors import get_academic_year_choices
from composn.decorators import admin_required, staff_perm_required, staff_required
from composn.excel import build_workbook, workbook_response
from composn.models import MAX_ACADEMIC_YEAR, MIN_ACADEMIC_YEAR, SystemSetting, UseRecord, User
from composn.templatetags.composn_tags import format_date_thai_full, format_date_thai_long, get_thai_day_name

from .models import (AbsenceReason, Attendance, CertificateRecipient, CertificateRecord, CertificateSetting,
                     CertificateYearStatus, QualificationResult, Score, ScoreApproval, ScoreCategory, TrainingDate,
                     WarningColor, absence_status, approve_year_scores, approved_years_set, compute_year_ranking,
                     current_year, is_future_attendance_date, non_training_dates_for_year, score_members_for_year,
                     sort_announcement_rows, to_ce_date, total_score, unapprove_year_scores, year_locked_message)
from .services import FoodError, delete_attendance, save_attendance, save_score, to_decimal


def _members(request, year=None):
    year_int = int(year or current_year())
    yy = f'{year_int % 100:02d}'
    qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=year_int)
    )
    center = request.GET.get('center', '') or request.POST.get('center', '')
    q = (request.GET.get('q', '') or '').strip()
    if center in ('DS', 'KM'):
        qs = qs.filter(camp_center=center)
    if q:
        qs = qs.filter(Q(camp1_code__icontains=q) | Q(camp2_code__icontains=q) | Q(first_name__icontains=q) |
                       Q(last_name__icontains=q))
    return qs.order_by('camp1_code'), center, q


def _score_members_for_year(request, year):
    """คืนรายชื่อนักเรียนสำหรับกรอก/สรุปคะแนนในปีการศึกษา `year`:
    1) นักเรียนปีการศึกษานั้น (ดูจากรหัส 2 ตัวหน้าของ camp1_code หรือ academic_year == year)
    2) นักเรียนปีอื่นที่ขอใช้สิทธิ์ในปีการศึกษานั้น (ยืนยันใช้สิทธิ์แล้ว is_used=True)
    คืนค่า (queryset, center, q, used_rights_user_ids_set)
    เรียงลำดับ: นักเรียนปัจจุบันขึ้นก่อน ตามด้วยนักเรียนใช้สิทธิ์ และเรียงตาม camp1_code"""
    year_int = int(year)
    yy = f'{year_int % 100:02d}'

    # ผู้ใช้สิทธิ์ในปีนั้นที่มีการยืนยันใช้สิทธิ์แล้ว
    used_qs = UseRecord.objects.filter(
        academic_year=year_int,
        is_used=True,
    )
    used_user_ids = set(used_qs.values_list('user_id', flat=True))

    is_current_q = ~Q(pk__in=used_user_ids) & (Q(camp1_code__istartswith=yy) | Q(academic_year=year_int))

    qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        (Q(camp1_code__istartswith=yy) | Q(academic_year=year_int)) | Q(pk__in=used_user_ids)
    )
    center = request.GET.get('center', '') or request.POST.get('center', '')
    q = (request.GET.get('q', '') or '').strip()
    if center in ('DS', 'KM'):
        qs = qs.filter(camp_center=center)
    if q:
        qs = qs.filter(Q(camp1_code__icontains=q) | Q(camp2_code__icontains=q) | Q(first_name__icontains=q) |
                       Q(last_name__icontains=q))

    # เรียงนักเรียนปัจจุบันขึ้นก่อน (cohort_order=0) ตามด้วยนักเรียนใช้สิทธิ์ (cohort_order=1) แล้วเรียงตาม camp1_code
    qs = qs.annotate(
        cohort_order=Case(
            When(is_current_q, then=Value(0)),
            default=Value(1),
            output_field=IntegerField(),
        )
    ).order_by('cohort_order', 'camp1_code')

    return qs, center, q, used_user_ids


def _parse_date(s):
    try:
        return dt.date.fromisoformat(s)
    except (TypeError, ValueError):
        return timezone.localdate()


def _parse_dt_local(val):
    if not val:
        return None
    try:
        naive = dt.datetime.fromisoformat(val.strip())
        return timezone.make_aware(naive) if timezone.is_naive(naive) else naive
    except ValueError:
        return None


# =============================================================================
# Attendance & แบบลงชื่อเข้าอบรม / สอบท้ายค่าย
# =============================================================================
@staff_perm_required('staff_can_attendance')
def staff_attendance(request):
    today = timezone.localdate()
    period = request.GET.get('period') or request.POST.get('period') or Attendance.PERIOD_MORNING
    if period not in dict(Attendance.PERIOD_CHOICES):
        period = Attendance.PERIOD_MORNING
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    members_qs, center, q = _members(request, year=year)
    members = list(members_qs)
    training_dates = list(TrainingDate.objects.filter(academic_year=year).order_by('date', 'id'))
    for idx, td in enumerate(training_dates):
        td.orig_idx = idx
        td.is_future = is_future_attendance_date(td.date, today=today)
        td.can_record = (td.event_type != TrainingDate.TYPE_HOLIDAY) and (not td.is_future)
        td.active_period_list = td.active_periods()
        td.has_morning = Attendance.PERIOD_MORNING in td.active_period_list
        td.has_afternoon = Attendance.PERIOD_AFTERNOON in td.active_period_list
        td.col_span = len(td.active_period_list) or 1

    raw_date = (request.GET.get('date') or request.POST.get('date') or '').strip()
    if raw_date:
        date = _parse_date(raw_date)
    else:
        # เลือกวันนี้หากตรงกับวันในตารางอบรม มิฉะนั้นเลือกวันอบรม/สอบล่าสุดที่มาถึงแล้ว
        exact_today = next((td.date for td in training_dates if to_ce_date(td.date) == today), None)
        arrived_tds = [td.date for td in training_dates if td.can_record]
        date = exact_today or (arrived_tds[-1] if arrived_tds else today)

    active_td = next((td for td in training_dates if td.date == date or to_ce_date(td.date) == to_ce_date(date)), None)
    for td in training_dates:
        td.is_pinned = bool(active_td is not None and td.pk == active_td.pk)
    if active_td is not None:
        display_training_dates = [active_td] + [td for td in training_dates if td.pk != active_td.pk]
    else:
        display_training_dates = list(training_dates)

    is_holiday = bool(active_td and active_td.event_type == TrainingDate.TYPE_HOLIDAY)
    is_exam_active = bool(active_td and active_td.event_type == TrainingDate.TYPE_EXAM)
    if is_exam_active and active_td.active_period_list and period not in active_td.active_period_list:
        period = active_td.active_period_list[0]
    is_future = is_future_attendance_date(date, today=today)
    is_year_approved = ScoreApproval.is_year_approved(year)
    can_record_active = (not is_holiday) and (not is_future)
    can_edit_attendance = can_record_active and (not is_year_approved)

    # ค่าเริ่มต้น = เข้าอบรม/เข้าสอบ (present) สำหรับวันอบรมและวันสอบที่มาถึงแล้วตามช่วงเวลาที่กำหนด (ไม่สร้างให้วันอนาคต/วันหยุด/ช่วงที่ไม่มีสอบ และไม่อัปเดตเพิ่มหากปีนั้นอนุมัติประกาศผลแล้ว)
    both_periods = [Attendance.PERIOD_MORNING, Attendance.PERIOD_AFTERNOON]
    arrived_tds_obj = [td for td in training_dates if td.can_record]
    if arrived_tds_obj and members and not is_year_approved:
        member_ids = [m.pk for m in members]
        arrived_td_dates = [td.date for td in arrived_tds_obj]
        existing_triples = set(
            Attendance.objects.filter(
                user_id__in=member_ids,
                date__in=arrived_td_dates,
                session_period__in=both_periods,
            ).values_list('user_id', 'date', 'session_period')
        )
        to_create = [
            Attendance(
                user_id=m.pk, academic_year=year, date=td.date, session_period=p,
                status=Attendance.STATUS_PRESENT, recorded_by=request.user,
            )
            for m in members for td in arrived_tds_obj for p in td.active_period_list
            if (m.pk, td.date, p) not in existing_triples
        ]
        if to_create:
            Attendance.objects.bulk_create(to_create, ignore_conflicts=True)

    records = {a.user_id: a for a in Attendance.objects.filter(date=date, session_period=period)}
    # วันสอบ และ วันหยุด ไม่นับเป็นวันอบรม (ไม่นับรวมในยอดขาดอบรม)
    excl_dates = non_training_dates_for_year(year)
    abs_qs = Attendance.objects.filter(academic_year=year, status__in=Attendance.ABSENT_STATUSES)
    if excl_dates:
        abs_qs = abs_qs.exclude(date__in=excl_dates)
    abs_counts = dict(abs_qs.values('user_id').annotate(c=Count('id')).values_list('user_id', 'c'))

    # เตรียมข้อมูลตารางรวมทุกวันที่ตั้งค่าไว้ (วันอบรมมี 2 ช่อง เช้า/บ่าย, วันสอบแสดงตามช่วงเวลาสอบ เช้า/บ่าย/ทั้งวัน)
    td_dates = [td.date for td in training_dates]
    matrix_att = {}
    if td_dates:
        for a in Attendance.objects.filter(date__in=td_dates, session_period__in=both_periods):
            matrix_att[(a.user_id, a.date, a.session_period)] = a

    rows = []
    for m in members:
        date_cells = []
        for td in display_training_dates:
            m_att = matrix_att.get((m.pk, td.date, Attendance.PERIOD_MORNING))
            a_att = matrix_att.get((m.pk, td.date, Attendance.PERIOD_AFTERNOON))
            m_status = m_att.status if m_att else (Attendance.STATUS_PRESENT if (td.can_record and td.has_morning) else '')
            a_status = a_att.status if a_att else (Attendance.STATUS_PRESENT if (td.can_record and td.has_afternoon) else '')
            date_cells.append({
                'td': td,
                'morning_att': m_att,
                'morning_status': m_status,
                'afternoon_att': a_att,
                'afternoon_status': a_status,
                'is_future': td.is_future,
                'can_record': td.can_record,
                'has_morning': td.has_morning,
                'has_afternoon': td.has_afternoon,
            })
        cur_att = records.get(m.pk)
        cur_status = cur_att.status if cur_att else (Attendance.STATUS_PRESENT if can_record_active else '')
        rows.append({
            'user': m,
            'att': cur_att,
            'status': cur_status,
            'date_cells': date_cells,
            'absence': absence_status(m, year, count=abs_counts.get(m.pk, 0)),
        })
    stats = {s: sum(1 for r in rows if r['status'] == s) for s, _ in Attendance.STATUS_CHOICES}
    stats['total_absent'] = sum(stats.get(s, 0) for s in Attendance.ABSENT_STATUSES)
    period_label = dict(Attendance.PERIOD_CHOICES).get(period, 'ภาคเช้า')
    exam_status_choices = [
        (Attendance.STATUS_PRESENT, 'เข้าสอบ'),
        (Attendance.STATUS_ABSENT, 'ขาดสอบ'),
    ]
    return render(request, 'food/staff_attendance.html', {
        'rows': rows, 'year': year, 'date': date, 'today': today, 'period': period,
        'period_label': period_label,
        'center': center, 'q': q, 'stats': stats,
        'period_choices': Attendance.PERIOD_CHOICES, 'status_choices': Attendance.STATUS_CHOICES,
        'exam_status_choices': exam_status_choices,
        'reasons': AbsenceReason.objects.filter(is_active=True),
        'colors': WarningColor.objects.all(), 'setting': SystemSetting.get(),
        'training_dates': training_dates,
        'display_training_dates': display_training_dates,
        'active_td': active_td,
        'is_holiday': is_holiday, 'is_exam_active': is_exam_active,
        'is_future': is_future, 'can_record_active': can_record_active,
        'is_year_approved': is_year_approved, 'can_edit_attendance': can_edit_attendance,
    })


@staff_perm_required('staff_can_attendance')
@require_POST
def attendance_save_api(request):
    """บันทึกแบบ Real-time (AJAX) คืนค่าจำนวนครั้งที่ขาดล่าสุด"""
    try:
        uid = int(request.POST.get('user_id') or 0)
    except (ValueError, TypeError):
        return JsonResponse({'ok': False, 'error': 'ไม่พบข้อมูลนักเรียน'}, status=400)
    student = User.objects.filter(pk=uid, role=User.ROLE_MEMBER).first()
    if not student:
        return JsonResponse({'ok': False, 'error': 'ไม่พบข้อมูลนักเรียน'}, status=404)
    date = _parse_date(request.POST.get('date'))
    period = request.POST.get('period')
    try:
        year_req = request.POST.get('year')
        att_year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        att_year = current_year()
    if is_future_attendance_date(date):
        return JsonResponse({'ok': False, 'error': 'ยังไม่ถึงวันที่ดังกล่าว ไม่สามารถบันทึกเวลาเรียนล่วงหน้าได้'}, status=400)
    td_match = TrainingDate.objects.filter(academic_year=att_year, date=date).first() or TrainingDate.objects.filter(date=date).first()
    lock_years = [att_year]
    if td_match and td_match.academic_year not in lock_years:
        lock_years.append(td_match.academic_year)
    for ly in lock_years:
        if ScoreApproval.is_year_approved(ly):
            return JsonResponse({'ok': False, 'error': year_locked_message(ly)}, status=403)
    if td_match and td_match.event_type == TrainingDate.TYPE_HOLIDAY:
        return JsonResponse({'ok': False, 'error': 'วันที่เลือกเป็นวันหยุด ไม่มีการลงเวลาเรียน'}, status=400)
    if td_match and td_match.event_type == TrainingDate.TYPE_EXAM and period not in td_match.active_periods():
        return JsonResponse({'ok': False, 'error': 'วันสอบดังกล่าวไม่ได้จัดสอบในช่วงเวลาที่เลือก'}, status=400)
    try:
        if request.POST.get('action') == 'delete':
            att = Attendance.objects.filter(user=student, date=date, session_period=period).first()
            status = delete_attendance(att, request.user, request) if att else absence_status(student, att_year)
            return JsonResponse({'ok': True, 'deleted': True, 'absence': status})
        att, status = save_attendance(student, date, period, request.POST.get('status'),
                                      request.POST.get('reason'), request.POST.get('note'),
                                      recorded_by=request.user, request=request, academic_year=att_year)
    except FoodError as e:
        return JsonResponse({'ok': False, 'error': str(e)}, status=400)
    return JsonResponse({'ok': True, 'id': att.pk, 'status': att.status, 'absence': status,
                         'recorded_at': timezone.localtime(att.recorded_at).strftime('%H:%M:%S')})


@staff_perm_required('staff_can_attendance')
def attendance_summary(request):
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    members, center, q = _members(request, year=year)
    setting = SystemSetting.get()
    # วันสอบ และ วันหยุด ไม่นับเป็นวันอบรม
    att_qs = Attendance.objects.filter(academic_year=year)
    excl_dates = non_training_dates_for_year(year)
    if excl_dates:
        att_qs = att_qs.exclude(date__in=excl_dates)
    agg = {r['user_id']: r for r in att_qs.values('user_id').annotate(
        present=Count('id', filter=Q(status=Attendance.STATUS_PRESENT)),
        absent_only=Count('id', filter=Q(status=Attendance.STATUS_ABSENT)),
        sick=Count('id', filter=Q(status=Attendance.STATUS_SICK)),
        personal=Count('id', filter=Q(status=Attendance.STATUS_PERSONAL)),
        absent=Count('id', filter=Q(status__in=Attendance.ABSENT_STATUSES)),
    )}
    rows = []
    for m in members:
        a = agg.get(m.pk, {'present': 0, 'absent_only': 0, 'sick': 0, 'personal': 0, 'absent': 0})
        rows.append({
            'user': m, 'present': a['present'], 'absent_only': a['absent_only'],
            'sick': a['sick'], 'personal': a['personal'], 'absent': a['absent'],
            'absence': absence_status(m, year, count=a['absent']),
        })
    if request.GET.get('only') == 'warn':
        rows = [r for r in rows if r['absent'] > 0]
    rows.sort(key=lambda r: (-r['absent'], r['user'].camp1_code))
    if request.GET.get('export') == 'xlsx':
        data = [[r['user'].camp1_code, r['user'].get_full_name(), r['present'], r['absent_only'],
                 r['sick'], r['personal'], r['absent'], r['absence']['text']] for r in rows]
        log_action(request.user, 'export', target_model='Attendance', new_value={'summary_year': year},
                   request=request)
        return workbook_response(build_workbook([(f'สรุปเวลาเรียน {year}', [
            'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'เข้าอบรม', 'ขาดอบรม', 'ลาป่วย', 'ลากิจ', 'รวมขาดอบรม', 'สถานะ'], data)]),
            f'สรุปเวลาเรียน_{year}.xlsx')
    return render(request, 'food/attendance_summary.html', {
        'rows': rows, 'year': year, 'center': center, 'q': q, 'setting': setting,
        'only': request.GET.get('only', ''),
    })


@staff_perm_required('staff_can_attendance')
def attendance_history(request, user_id):
    student = get_object_or_404(User, pk=user_id, role=User.ROLE_MEMBER)
    if request.method == 'POST':
        att = get_object_or_404(Attendance, pk=request.POST.get('id'), user=student)
        if ScoreApproval.is_year_approved(att.academic_year):
            messages.error(request, year_locked_message(att.academic_year))
            return redirect(request.path)
        try:
            delete_attendance(att, request.user, request)
        except FoodError as e:
            messages.error(request, str(e))
        else:
            messages.success(request, 'ลบรายการ (Soft Delete) แล้ว')
        return redirect(request.path)
    return render(request, 'food/attendance_history.html', {
        'student': student, 'records': Attendance.objects.filter(user=student).select_related('recorded_by'),
        'absence': absence_status(student),
        'approved_years': approved_years_set(),
    })


@staff_perm_required('staff_can_signin_sheet')
def signin_sheet(request):
    """แบบลงชื่อเข้าอบรม / วันสอบท้ายค่าย (รองรับช่วงวันที่และ Export PDF)
    - วันอบรม: เฉพาะนักเรียนปัจจุบัน (ไม่มีนักเรียนใช้สิทธิ์), 1 วัน = 2 แผ่น (ภาคเช้า + ภาคบ่าย)
    - วันสอบ: แยกแผ่นนักเรียนปัจจุบัน กับ นักเรียนใช้สิทธิ์ คนละแผ่นกัน"""
    from .signin_pdf import build_signin_pdf

    setting = SystemSetting.get()
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year = int(setting.current_academic_year)

    mode = (request.GET.get('mode') or 'auto').strip()  # auto, training, exam
    if mode not in ('auto', 'training', 'exam'):
        mode = 'auto'
    try:
        camp = int(request.GET.get('camp') or 1)
    except (ValueError, TypeError):
        camp = 1
    if camp not in (1, 2):
        camp = 1
    period_opt = request.GET.get('period', 'both')
    if period_opt not in ('both', 'morning', 'afternoon'):
        period_opt = 'both'
    center = request.GET.get('center', '').upper()
    if center not in ('DS', 'KM'):
        center = ''

    all_tds = list(TrainingDate.objects.filter(academic_year=year).order_by('date', 'id'))
    training_dates = [td for td in all_tds if td.event_type != TrainingDate.TYPE_HOLIDAY]
    td_by_ce = {to_ce_date(td.date): td for td in all_tds}

    selected_td = None
    td_id_raw = (request.GET.get('td_id') or '').strip()
    date_start_str = (request.GET.get('date_start') or '').strip()
    date_end_str = (request.GET.get('date_end') or '').strip()
    legacy_date_str = (request.GET.get('date') or '').strip()
    range_all = request.GET.get('range_all') == '1'

    if range_all and training_dates:
        ce_start = to_ce_date(training_dates[0].date)
        ce_end = to_ce_date(training_dates[-1].date)
        mode = 'auto'
    elif td_id_raw:
        try:
            selected_td = TrainingDate.objects.filter(pk=int(td_id_raw)).first()
        except ValueError:
            selected_td = None
        if selected_td:
            ce_start = ce_end = to_ce_date(selected_td.date)
            if selected_td.event_type in ('training', 'exam'):
                mode = selected_td.event_type
        else:
            ce_start = ce_end = to_ce_date(_parse_date(date_start_str or legacy_date_str))
    elif date_start_str or date_end_str:
        ce_start = to_ce_date(_parse_date(date_start_str or date_end_str))
        ce_end = to_ce_date(_parse_date(date_end_str or date_start_str))
    elif legacy_date_str:
        ce_start = ce_end = to_ce_date(_parse_date(legacy_date_str))
    else:
        ce_start = ce_end = timezone.localdate()

    if ce_end < ce_start:
        ce_start, ce_end = ce_end, ce_start
    if (ce_end - ce_start).days > 60:
        ce_end = ce_start + dt.timedelta(days=60)
        messages.warning(request, 'กำหนดช่วงวันที่สูงสุด 60 วันต่อครั้ง ระบบปรับวันที่สิ้นสุดให้อัตโนมัติ')

    # สร้างรายการวันที่ในช่วงที่เลือก (ข้ามวันหยุดอัตโนมัติ)
    is_multi_day = (ce_start != ce_end)
    day_entries = []
    skipped_holidays = 0
    cur_d = ce_start
    while cur_d <= ce_end:
        td_match = td_by_ce.get(cur_d)
        if td_match and td_match.event_type == TrainingDate.TYPE_HOLIDAY:
            skipped_holidays += 1
            cur_d += dt.timedelta(days=1)
            continue
        if mode in ('training', 'exam'):
            day_mode = mode
        else:
            day_mode = (
                td_match.event_type
                if (td_match and td_match.event_type in ('training', 'exam'))
                else 'training'
            )
        day_name = get_thai_day_name(cur_d.weekday())
        day_entries.append({
            'date': cur_d,
            'day_mode': day_mode,
            'td': td_match,
            'day_name': day_name,
            'date_thai_header': f'วัน{day_name}ที่ {format_date_thai_long(cur_d)}',
        })
        cur_d += dt.timedelta(days=1)

    # 1. รายชื่อนักเรียนปัจจุบัน (ใช้ในวันอบรม และแยกเป็นแผ่นนักเรียนปัจจุบันในวันสอบ)
    yy = f'{int(year) % 100:02d}'
    cur_qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=int(year))
    )
    if camp == 2:
        cur_qs = cur_qs.filter(camp2_code__isnull=False).exclude(camp2_code='')
    if center:
        cur_qs = cur_qs.filter(camp_center=center)
    cur_qs = cur_qs.order_by('camp2_code' if camp == 2 else 'camp1_code')

    current_user_ids = set()
    current_rows = []
    for idx, u in enumerate(cur_qs, 1):
        current_user_ids.add(u.pk)
        disp_code = (u.camp2_code if (camp == 2 and u.camp2_code) else u.camp1_code)
        current_rows.append({
            'no': idx,
            'code': disp_code,
            'camp1_code': u.camp1_code,
            'camp_code': disp_code,
            'full_name': u.get_full_name(),
            'school': u.school_name,
            'center': u.camp_center,
            'grade': u.level_for_year(year) or u.current_level,
        })

    # 2. รายชื่อนักเรียนใช้สิทธิ์ (ไม่มีในวันอบรม — แสดงเฉพาะวันสอบและขึ้นแผ่นแยกคนละแผ่นกับนักเรียนปัจจุบัน)
    rights_qs = UseRecord.objects.filter(
        academic_year=year, camp_level=camp, is_used=True
    ).exclude(user_id__in=current_user_ids).select_related('user')
    if center:
        rights_qs = rights_qs.filter(center_snapshot=center)
    rights_qs = rights_qs.order_by('rights_code', 'user__camp1_code')

    rights_rows = []
    for idx, rec in enumerate(rights_qs, 1):
        rights_rows.append({
            'no': idx,
            'code': rec.rights_code or rec.camp_code_snapshot,
            'camp1_code': rec.user.camp1_code,
            'camp_code': rec.camp_code_snapshot,
            'full_name': rec.full_name_snapshot or rec.user.get_full_name(),
            'school': rec.school_snapshot or rec.user.school_name,
            'center': rec.center_snapshot or rec.user.camp_center,
            'grade': rec.grade_level or rec.user.level_for_year(year) or rec.user.current_level,
        })

    # สร้างแผ่น A4 แนวนอน (สูงสุด 52 คน/แผ่น · 2 ฝั่งซ้าย-ขวา)
    if period_opt == 'morning':
        periods_list = [('morning', 'ภาคเช้า')]
    elif period_opt == 'afternoon':
        periods_list = [('afternoon', 'ภาคบ่าย')]
    else:
        periods_list = [('morning', 'ภาคเช้า'), ('afternoon', 'ภาคบ่าย')]

    per_sheet = 52

    def _append_group_sheets(target_list, group_rows, entry, p_key, p_label, group_key, group_label):
        chunks = [group_rows[i:i + per_sheet] for i in range(0, len(group_rows), per_sheet)] if group_rows else [[]]
        for page_idx, chunk in enumerate(chunks, 1):
            half = (len(chunk) + 1) // 2
            left_list = chunk[:half]
            right_list = chunk[half:]
            paired = [
                {'left': left_list[i], 'right': right_list[i] if i < len(right_list) else None}
                for i in range(half)
            ]
            if entry['day_mode'] == 'exam':
                title_l4 = f'แบบลงชื่อเข้าสอบ — {group_label} ({p_label})'
            else:
                title_l4 = f'แบบลงชื่อเข้าอบรม ({p_label})'
            target_list.append({
                'date': entry['date'],
                'date_thai_header': entry['date_thai_header'],
                'day_mode': entry['day_mode'],
                'group_key': group_key,
                'group_label': group_label,
                'title_line4': title_l4,
                'period_key': p_key,
                'period_label': p_label,
                'page_no': page_idx,
                'total_pages': len(chunks),
                'count': len(chunk),
                'paired_rows': paired,
            })

    sheets = []
    for entry in day_entries:
        for p_key, p_label in periods_list:
            if entry['day_mode'] == 'exam':
                if current_rows:
                    _append_group_sheets(sheets, current_rows, entry, p_key, p_label, 'current', 'นักเรียนปัจจุบัน')
                if rights_rows:
                    _append_group_sheets(sheets, rights_rows, entry, p_key, p_label, 'rights', 'นักเรียนใช้สิทธิ์')
                if not current_rows and not rights_rows:
                    _append_group_sheets(sheets, [], entry, p_key, p_label, 'current', 'นักเรียนปัจจุบัน')
            else:
                # วันอบรม: เฉพาะนักเรียนปัจจุบันเท่านั้น ไม่มีนักเรียนใช้สิทธิ์
                _append_group_sheets(sheets, current_rows, entry, p_key, p_label, 'current', 'นักเรียนปัจจุบัน')

    has_exam_day = any(e['day_mode'] == 'exam' for e in day_entries)
    date_tag = f'{ce_start}' if ce_start == ce_end else f'{ce_start}_ถึง_{ce_end}'

    # Export PDF (A4 แนวนอน 50–52 คน/แผ่น)
    if request.GET.get('export') == 'pdf':
        pdf_bytes = build_signin_pdf(sheets, setting, camp=camp)
        log_action(request.user, 'export', target_model='Attendance',
                   new_value={'signin_pdf': date_tag, 'camp': camp, 'sheets': len(sheets)}, request=request)
        fname = f'แบบลงชื่อ_ค่าย{camp}_{date_tag}.pdf'
        resp = HttpResponse(pdf_bytes, content_type='application/pdf')
        resp['Content-Disposition'] = f"attachment; filename=\"signin_camp{camp}_{ce_start}.pdf\"; filename*=UTF-8''{quote(fname)}"
        return resp

    # Export Word (.docx)
    if request.GET.get('export') == 'docx':
        from .signin_docx import build_signin_docx, HAS_DOCX
        if not HAS_DOCX:
            messages.error(request, 'ไม่พบ python-docx กรุณาติดตั้ง: pip install python-docx')
            return redirect(request.path + '?' + request.GET.urlencode().replace('export=docx', ''))
        docx_bytes = build_signin_docx(sheets, setting, camp=camp)
        log_action(request.user, 'export', target_model='Attendance',
                   new_value={'signin_docx': date_tag, 'camp': camp, 'sheets': len(sheets)}, request=request)
        fname = f'แบบลงชื่อ_ค่าย{camp}_{date_tag}.docx'
        resp = HttpResponse(docx_bytes,
                            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        resp['Content-Disposition'] = f"attachment; filename=\"signin_camp{camp}_{ce_start}.docx\"; filename*=UTF-8''{quote(fname)}"
        return resp

    # Export Excel
    if request.GET.get('export') == 'xlsx':
        headers = ['ลำดับ', 'เลขประจำตัวค่าย', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'โรงเรียน',
                   'ลงชื่อ (ภาคเช้า)', 'ลงชื่อ (ภาคบ่าย)', 'หมายเหตุ']
        wb_sheets = []
        wb_sheets.append((
            'นักเรียนปัจจุบัน',
            headers,
            [[r['no'], r['code'], r['full_name'], r['grade'], r['school'], '', '', ''] for r in current_rows],
        ))
        if has_exam_day or mode == 'exam':
            wb_sheets.append((
                'นักเรียนใช้สิทธิ์',
                headers,
                [[r['no'], r['code'], r['full_name'], r['grade'], r['school'], '', '', ''] for r in rights_rows],
            ))
        title_str = f'แบบลงชื่อ_ค่าย{camp}_{date_tag}'
        wb = build_workbook(wb_sheets)
        return workbook_response(wb, f'{title_str}.xlsx')

    # ค่าสำหรับแสดงผลบนหน้าเว็บ
    first_header = day_entries[0]['date_thai_header'] if day_entries else f'วัน{get_thai_day_name(ce_start.weekday())}ที่ {format_date_thai_long(ce_start)}'
    pdf_query = urlencode({
        'year': year, 'mode': mode, 'camp': camp, 'period': period_opt,
        'center': center, 'date_start': ce_start.isoformat(), 'date_end': ce_end.isoformat(),
        'export': 'pdf',
    })
    xlsx_query = urlencode({
        'year': year, 'mode': mode, 'camp': camp, 'period': period_opt,
        'center': center, 'date_start': ce_start.isoformat(), 'date_end': ce_end.isoformat(),
        'export': 'xlsx',
    })
    docx_query = urlencode({
        'year': year, 'mode': mode, 'camp': camp, 'period': period_opt,
        'center': center, 'date_start': ce_start.isoformat(), 'date_end': ce_end.isoformat(),
        'export': 'docx',
    })

    return render(request, 'food/signin_sheet.html', {
        'rows': current_rows,
        'current_rows': current_rows,
        'rights_rows': rights_rows,
        'sheets': sheets,
        'day_entries': day_entries,
        'skipped_holidays': skipped_holidays,
        'has_exam_day': has_exam_day,
        'is_multi_day': is_multi_day,
        'year': year, 'mode': mode, 'camp': camp,
        'period_opt': period_opt, 'center': center,
        'date': ce_start,
        'date_start': ce_start,
        'date_end': ce_end,
        'date_full': format_date_thai_full(ce_start),
        'date_thai_header': first_header,
        'day_name': get_thai_day_name(ce_start.weekday()),
        'training_dates': training_dates,
        'selected_td': selected_td,
        'setting': setting,
        'pdf_query': pdf_query,
        'xlsx_query': xlsx_query,
        'docx_query': docx_query,
    })


# =============================================================================
# Scores (กรอกรายบุคคล / นำเข้า Excel / Copy-Paste จาก Excel / Export)
# =============================================================================
def _leaf_score_categories(year):
    """คืนรายการหมวดคะแนนที่ใช้กรอกคะแนนได้ (ถ้าหัวข้อใหญ่มีหัวข้อย่อย ให้เลือกกรอกที่หัวข้อย่อย; ถ้าไม่มีให้กรอกที่หัวข้อใหญ่)"""
    main_cats = ScoreCategory.objects.filter(academic_year=year, parent__isnull=True).order_by('order', 'id')
    result = []
    for m_idx, mc in enumerate(main_cats, 1):
        subs = list(mc.subcategories.all().order_by('order', 'id'))
        if subs:
            for s_idx, sc in enumerate(subs, 1):
                sc.display_code = f'{m_idx}.{s_idx}'
                sc.display_label = f'{m_idx}.{s_idx} {sc.name} ({mc.name})'
                result.append(sc)
        else:
            mc.display_code = f'{m_idx}'
            mc.display_label = f'{m_idx}. {mc.name}'
            result.append(mc)
    return result


@staff_perm_required('staff_can_scores')
def staff_scores(request):
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    main_categories = list(ScoreCategory.objects.filter(academic_year=year, parent__isnull=True).order_by('order', 'id'))
    for m_idx, mc in enumerate(main_categories, 1):
        mc.main_no = m_idx
        subs = list(mc.subcategories.all().order_by('order', 'id'))
        for s_idx, sc in enumerate(subs, 1):
            sc.sub_no = f'{m_idx}.{s_idx}'
            sc.display_code = sc.sub_no
            sc.display_label = f'{sc.sub_no} {sc.name}'
        mc.sub_list = subs

    # รองรับทั้งการเลือกตามหัวข้อใหญ่ (main_id) หรือเลือกจาก category (เผื่อลิงก์เดิมส่ง category เป็นหัวข้อย่อยมา)
    main_id = request.GET.get('main_id') or request.POST.get('main_id')
    cat_id = request.GET.get('category') or request.POST.get('category')
    selected_main = None
    if main_id:
        selected_main = next((mc for mc in main_categories if str(mc.pk) == str(main_id)), None)
    elif cat_id:
        for mc in main_categories:
            if str(mc.pk) == str(cat_id) or any(str(sc.pk) == str(cat_id) for sc in mc.sub_list):
                selected_main = mc
                break
    if not selected_main and main_categories:
        selected_main = main_categories[0]

    # รายการคอลัมน์คะแนนที่จะแสดงพร้อมกันในตาราง (ถ้าหัวข้อใหญ่มีหัวข้อย่อย แสดงทุกหัวข้อย่อยพร้อมกัน; ถ้าไม่มี แสดงหัวข้อใหญ่นั้น 1 คอลัมน์)
    if selected_main:
        if selected_main.sub_list:
            sub_cols = selected_main.sub_list
        else:
            selected_main.display_code = f'{selected_main.main_no}'
            selected_main.display_label = f'{selected_main.main_no}. {selected_main.name}'
            sub_cols = [selected_main]
    else:
        sub_cols = []

    category = sub_cols[0] if sub_cols else None
    if cat_id and sub_cols:
        category = next((c for c in sub_cols if str(c.pk) == str(cat_id)), category)

    try:
        exam_no = max(1, int(request.GET.get('exam_no') or request.POST.get('exam_no') or 1))
    except ValueError:
        exam_no = 1
    members, center, q, used_user_ids = _score_members_for_year(request, year)
    is_year_approved = ScoreApproval.is_year_approved(year)

    # Export คะแนนรวมทุกหัวข้อใหญ่และทุกหัวข้อย่อยเป็น Excel
    if request.GET.get('export') == 'all_xlsx':
        return _export_all_scores_workbook(request, year)

    # Export คะแนนของหัวข้อใหญ่นี้ (รวมทุกหัวข้อย่อย) เป็น Excel
    if request.GET.get('export') == 'xlsx' and selected_main and sub_cols:
        scores_qs = Score.objects.filter(category__in=sub_cols, exam_no=exam_no)
        scores_map = {(s.user_id, s.category_id): s for s in scores_qs}
        headers = ['รหัสค่าย 1', 'ชื่อ-นามสกุล'] + \
                  [f'{sc.display_label} (เต็ม {sc.default_max_score})' for sc in sub_cols] + ['รวมคะแนนที่ได้']
        xrows = []
        for m in members:
            row_vals = [m.camp1_code, m.get_full_name()]
            sum_sc = Decimal('0')
            has_sc = False
            for sc in sub_cols:
                s_obj = scores_map.get((m.pk, sc.pk))
                if s_obj is not None:
                    row_vals.append(s_obj.score)
                    sum_sc += s_obj.score
                    has_sc = True
                else:
                    row_vals.append('')
            row_vals.append(sum_sc if has_sc else '')
            xrows.append(row_vals)
        wb = build_workbook([(f'{selected_main.name}'[:31], headers, xrows)])
        return workbook_response(wb, f'คะแนน_{selected_main.name}.xlsx')

    if request.method == 'POST' and selected_main and sub_cols:
        def _scores_url(extra=None, anchor=''):
            params = {'year': year, 'main_id': selected_main.pk, 'exam_no': exam_no, 'center': center, 'q': q}
            if extra:
                params.update(extra)
            return f'{request.path}?{urlencode(params)}{anchor}'

        action = request.POST.get('action', 'save_manual')
        if action != 'cancel_import_preview' and is_year_approved:
            messages.error(request, year_locked_message(year))
            return redirect(_scores_url())

        if not selected_main.is_currently_open() and not request.user.is_admin_role:
            messages.error(request, f'หมวดคะแนน "{selected_main.name}" ปิดการกรอกคะแนนแล้ว หรืออยู่นอกเวลาที่กำหนด')
            return redirect(_scores_url())

        target_cat_id = request.POST.get('category')
        target_cat = next((c for c in sub_cols if str(c.pk) == str(target_cat_id)), None)
        writable_sub_cols = sub_cols if request.user.is_admin_role else [sc for sc in sub_cols if sc.is_currently_open()]

        # 1) อัปโหลดไฟล์ Excel (.xlsx) เพื่อตรวจสอบคอลัมน์และรายชื่อ พร้อมแสดงผล Preview ปัญหาก่อนนำเข้า
        if action == 'import_excel':
            allowed_members_qs, _ = score_members_for_year(year)
            excel_file = request.FILES.get('excel_file')
            if not excel_file:
                messages.error(request, 'กรุณาเลือกไฟล์ Excel (.xlsx) ก่อนกดตรวจสอบ')
                return redirect(_scores_url())

            from .excel_import import preview_score_excel
            preview_res = preview_score_excel(excel_file, allowed_members_qs, sub_cols, exam_no=exam_no, year=year)
            if not preview_res.get('success'):
                messages.error(request, preview_res.get('error_message', 'ไม่สามารถอ่านไฟล์ Excel ได้'))
                return redirect(_scores_url())

            preview_res['file_name'] = getattr(excel_file, 'name', '')
            preview_res['main_label'] = f'{selected_main.main_no}. {selected_main.name}'
            request.session['score_import_preview'] = {
                'main_id': selected_main.pk,
                'exam_no': exam_no,
                'year': year,
                'data': preview_res,
            }
            if preview_res.get('has_problems'):
                msg = f'ตรวจสอบไฟล์ "{preview_res["file_name"]}" สำหรับหัวข้อ "{preview_res["main_label"]}": ' \
                      f'พร้อมนำเข้า {preview_res["valid_count"]} คน, พบปัญหา {preview_res["invalid_count"]} คน'
                if preview_res.get('missing_sub_cols') or any('มากกว่าหัวข้อย่อย' in i for i in preview_res.get('column_issues', [])):
                    msg += ' และจำนวนคอลัมน์คะแนนไม่ตรงกับหัวข้อย่อย'
                messages.warning(request, msg + ' กรุณาตรวจสอบตาราง Preview ด้านล่าง')
            else:
                messages.success(request, f'ตรวจสอบไฟล์ "{preview_res["file_name"]}" สำหรับหัวข้อ "{preview_res["main_label"]}": '
                                          f'ข้อมูลถูกต้องพร้อมนำเข้า {preview_res["valid_count"]} คน กรุณาตรวจสอบตาราง Preview ด้านล่างแล้วกดยืนยันนำเข้า')

            return redirect(_scores_url({'show_preview': '1'}, '#import-preview-section'))

        # ยืนยันนำเข้าคะแนนหลังตรวจสอบ Preview
        if action == 'confirm_import_excel':
            pdata = request.session.get('score_import_preview')
            if not pdata or not pdata.get('data'):
                messages.error(request, 'ไม่พบข้อมูล Preview หรือหมดอายุ กรุณาอัปโหลดไฟล์ใหม่อีกครั้ง')
                return redirect(_scores_url())
            if str(pdata.get('main_id')) != str(selected_main.pk) or pdata.get('exam_no') != exam_no:
                messages.error(request, 'ข้อมูล Preview ไม่ตรงกับหัวข้อคะแนน/ครั้งที่สอบที่เลือก กรุณาอัปโหลดไฟล์ใหม่อีกครั้ง')
                return redirect(_scores_url())
            request.session.pop('score_import_preview', None)

            override_users = {
                k[len('override_user_'):]: v.strip()
                for k, v in request.POST.items()
                if k.startswith('override_user_') and v.strip()
            }
            all_year_qs, _ = score_members_for_year(year)
            allowed_ids = set(all_year_qs.values_list('pk', flat=True))

            from .excel_import import execute_score_import
            saved, errors = execute_score_import(
                pdata['data'], request.user, exam_no, writable_sub_cols, request=request,
                override_users=override_users, allowed_user_ids=allowed_ids,
            )
            if errors:
                messages.warning(request, f'พบข้อผิดพลาดบางรายการ ({len(errors)} รายการ): ' + '; '.join(errors[:5]))
            messages.success(request, f'นำเข้าคะแนนจากไฟล์ Excel สำเร็จเรียบร้อย ({saved} รายการ)')
            return redirect(_scores_url())

        # ยกเลิก Preview
        if action == 'cancel_import_preview':
            request.session.pop('score_import_preview', None)
            return redirect(_scores_url())

        # 2) วางข้อมูลที่ Copy มาจาก Excel / Google Sheets (paste_excel)
        if action == 'paste_excel':
            from .excel_import import norm_thai_name
            all_year_qs, _ = score_members_for_year(year)
            user_map = {}
            name_owners = {}
            all_year_by_pk = {}
            for u in all_year_qs:
                all_year_by_pk[u.pk] = u
                user_map[u.camp1_code.upper()] = u
                if u.camp2_code:
                    user_map[u.camp2_code.upper()] = u
                for n in {norm_thai_name(u.get_full_name()), norm_thai_name(f'{u.first_name}{u.last_name}')}:
                    if n:
                        name_owners.setdefault(n, set()).add(u.pk)
            user_by_name = {n: all_year_by_pk[next(iter(pks))] for n, pks in name_owners.items() if len(pks) == 1}

            saved, errors = 0, []
            raw_rows = []
            raw_text = request.POST.get('paste_data', '')
            for line in raw_text.splitlines():
                if not line.strip():
                    continue
                if '\t' in line:
                    parts = [p.strip() for p in line.split('\t')]
                elif ',' in line:
                    parts = [p.strip() for p in line.split(',')]
                elif ';' in line:
                    parts = [p.strip() for p in line.split(';')]
                else:
                    parts = [p.strip() for p in re.split(r'\s{2,}', line.strip()) if p.strip() != '']
                    if len(parts) == 1:
                        parts = line.strip().split()
                raw_rows.append(parts)

            # ตรวจสอบว่ามีรหัสค่ายหรือชื่อ-นามสกุลนำหน้าหรือไม่
            has_codes = any(bool(r and re.match(r'^\d{2}(DS|KM)\d{2}$', str(r[0] or '').strip().upper())) for r in raw_rows)
            has_names = (not has_codes) and any(bool(r and norm_thai_name(r[0]) in user_by_name) for r in raw_rows)
            paste_orientation = request.POST.get('paste_orientation', 'auto')

            if has_codes or has_names:
                for r in raw_rows:
                    if not r:
                        continue
                    c0_raw = str(r[0] or '').strip()
                    c0 = c0_raw.upper()
                    if has_codes:
                        if not re.match(r'^\d{2}(DS|KM)\d{2}$', c0):
                            continue
                        u = user_map.get(c0)
                        if not u:
                            errors.append(f'ไม่พบรหัส {c0}')
                            continue
                        # ตรวจสอบว่าเป็นไฟล์ที่ Export ออกไป (มีคอลัมน์ รหัส1, ชื่อ นำหน้า 2 คอลัมน์ หรือแบบเดิม 5 คอลัมน์) หรือเป็น [รหัส, คะแนนย่อย1, คะแนนย่อย2, ...]
                        if len(r) >= 5 + len(sub_cols) and not re.match(r'^-?\d+(\.\d+)?$', str(r[2] or '')):
                            score_cells = r[5:5 + len(sub_cols)]
                        elif len(r) >= 2 + len(sub_cols) and not re.match(r'^-?\d+(\.\d+)?$', str(r[1] or '')):
                            score_cells = r[2:2 + len(sub_cols)]
                        else:
                            score_cells = r[1:]
                    else:
                        if not c0_raw or any(kw in c0_raw.lower() for kw in ['ชื่อ', 'รหัส', 'name', 'code']):
                            continue
                        u = user_by_name.get(norm_thai_name(c0_raw))
                        if not u:
                            errors.append(f'ไม่พบชื่อ "{c0_raw}" ในระบบปีการศึกษา {year}')
                            continue
                        score_cells = r[1:]

                    # ถ้าระบุ target_cat เฉพาะเจาะจง 1 หมวด และวางมาแค่ 1-2 คอลัมน์คะแนน
                    if target_cat and len(sub_cols) == 1:
                        sc_raw = score_cells[0] if score_cells else ''
                        mx_raw = score_cells[1] if len(score_cells) > 1 else None
                        if sc_raw == '':
                            continue
                        try:
                            sc_val = to_decimal(sc_raw)
                            mx_val = to_decimal(mx_raw) if mx_raw not in (None, '') else (to_decimal(request.POST.get('default_max')) or target_cat.default_max_score)
                            save_score(u, target_cat, exam_no, sc_val, mx_val, recorded_by=request.user, request=request)
                            saved += 1
                        except FoodError as e:
                            errors.append(str(e))
                    else:
                        # วางพร้อมกันหลายหัวข้อย่อยตามลำดับคอลัมน์
                        for idx_col, sc_col in enumerate(sub_cols):
                            if idx_col >= len(score_cells):
                                break
                            cell_val = score_cells[idx_col]
                            if cell_val == '':
                                continue
                            try:
                                sc_val = to_decimal(cell_val)
                                save_score(u, sc_col, exam_no, sc_val, sc_col.default_max_score,
                                           recorded_by=request.user, request=request)
                                saved += 1
                            except FoodError as e:
                                errors.append(str(e))
            else:
                # กรณีคัดลอกเฉพาะคะแนนตัวเลขมาวาง (ไม่มีรหัสค่าย/ชื่อนำหน้า) รองรับทั้งแนวตั้ง แนวนอน และตาราง 2D
                members_list = list(members)
                num_rows = len(raw_rows)
                num_cols = max((len(r) for r in raw_rows), default=0)

                target_sc_col = target_cat or sub_cols[0]
                target_col_idx = sub_cols.index(target_sc_col) if target_sc_col in sub_cols else 0
                remaining_cols = len(sub_cols) - target_col_idx

                # ถ้าเลือกแนวตั้ง หรือข้อมูลเป็นคอลัมน์เดียวหลายแถว หรือแถวเดียวที่ยาวเกินจำนวนคอลัมน์ที่เหลือ -> วางแนวตั้ง
                # ส่วนตาราง 2 มิติ (หลายแถว x หลายคอลัมน์) ให้วางตามแถว x คอลัมน์
                is_vertical = (paste_orientation == 'vertical') or (
                    paste_orientation != 'horizontal' and (
                        num_cols == 1 or len(sub_cols) == 1 or (num_rows == 1 and num_cols > remaining_cols)
                    )
                )

                if is_vertical:
                    # วางแนวตั้ง: หยอดคะแนนลงตามลำดับรายชื่อนักเรียนในหัวข้อคะแนนที่เลือก
                    flat_vals = []
                    for r in raw_rows:
                        flat_vals.extend(r)
                    for idx_u, sc_raw in enumerate(flat_vals):
                        if idx_u >= len(members_list):
                            break
                        if sc_raw == '':
                            continue
                        u = members_list[idx_u]
                        try:
                            sc_val = to_decimal(sc_raw)
                            save_score(u, target_sc_col, exam_no, sc_val, target_sc_col.default_max_score,
                                       recorded_by=request.user, request=request)
                            saved += 1
                        except FoodError as e:
                            errors.append(str(e))
                else:
                    # วางแนวนอน / ตาราง 2D: หยอดคะแนนข้ามหัวข้อย่อยของนักเรียนแต่ละคน
                    for idx_u, r in enumerate(raw_rows):
                        if idx_u >= len(members_list):
                            break
                        u = members_list[idx_u]
                        for idx_col, cell_val in enumerate(r):
                            actual_col = target_col_idx + idx_col
                            if actual_col >= len(sub_cols):
                                break
                            sc_col = sub_cols[actual_col]
                            if cell_val == '':
                                continue
                            try:
                                sc_val = to_decimal(cell_val)
                                save_score(u, sc_col, exam_no, sc_val, sc_col.default_max_score,
                                           recorded_by=request.user, request=request)
                                saved += 1
                            except FoodError as e:
                                errors.append(str(e))

            if errors:
                messages.warning(request, f'พบข้อผิดพลาด {len(errors)} รายการ: ' + '; '.join(errors[:8]))
            messages.success(request, f'นำเข้า/วางคะแนนสำเร็จ {saved} รายการ')
            return redirect(_scores_url())

        # 2) บันทึกแบบกรอกในตาราง (รองรับทั้งแบบหลายหัวข้อย่อยพร้อมกัน score_<user_id>_<cat_id> และแบบเดิม score_<user_id>)
        saved, errors = 0, []
        try:
            default_max = to_decimal(request.POST.get('default_max'))
        except FoodError as e:
            default_max = None
            errors.append(str(e))
        is_admin = request.user.is_admin_role
        open_sub_ids = {sc.pk for sc in sub_cols if is_admin or sc.is_currently_open()}
        for m in members:
            # กรณีส่งแบบเก่า score_<user_id>
            legacy_key = f'score_{m.pk}'
            if legacy_key in request.POST and target_cat:
                if target_cat.pk not in open_sub_ids:
                    continue
                try:
                    sc_val = to_decimal(request.POST.get(legacy_key))
                    mx_val = to_decimal(request.POST.get(f'max_{m.pk}')) or default_max or target_cat.default_max_score
                    save_score(m, target_cat, exam_no, sc_val, mx_val, recorded_by=request.user, request=request)
                    saved += 1
                except FoodError as e:
                    errors.append(str(e))
                continue

            # กรณีส่งแบบตาราง Excel หลายคอลัมน์พร้อมกัน score_<user_id>_<cat_id>
            for sc_col in sub_cols:
                cell_key = f'score_{m.pk}_{sc_col.pk}'
                if cell_key not in request.POST or sc_col.pk not in open_sub_ids:
                    continue
                try:
                    sc_val = to_decimal(request.POST.get(cell_key))
                    mx_val = sc_col.default_max_score or default_max or Decimal('100')
                    save_score(m, sc_col, exam_no, sc_val, mx_val, recorded_by=request.user, request=request)
                    saved += 1
                except FoodError as e:
                    errors.append(str(e))

        if errors:
            messages.error(request, 'บันทึกไม่สำเร็จบางรายการ: ' + '; '.join(errors[:10]))
        messages.success(request, f'บันทึกคะแนนเรียบร้อย ({saved} ช่อง)')
        return redirect(f'{request.path}?{urlencode({"year": year, "main_id": selected_main.pk, "exam_no": exam_no, "center": center, "q": q})}')

    # เตรียมข้อมูลแสดงตารางแบบ Excel (แถว = นักเรียน, คอลัมน์ = ทุกหัวข้อย่อยของหัวข้อใหญ่นี้)
    scores_qs = Score.objects.filter(category__in=sub_cols, exam_no=exam_no) if sub_cols else Score.objects.none()
    scores_map = {(s.user_id, s.category_id): s for s in scores_qs}
    yy_prefix = f'{int(year) % 100:02d}'
    rows = []
    for m in members:
        m.display_level = m.level_for_year(year)
        cells = []
        row_sum = Decimal('0')
        has_any = False
        for sc_col in sub_cols:
            s_obj = scores_map.get((m.pk, sc_col.pk))
            if s_obj is not None:
                row_sum += s_obj.score
                has_any = True
            cells.append({'category': sc_col, 'score': s_obj})
        is_current_cohort = m.camp1_code.upper().startswith(yy_prefix) or (m.academic_year == int(year))
        rows.append({
            'user': m,
            'is_current_cohort': is_current_cohort,
            'is_rights_user': not is_current_cohort,
            'has_used_rights': m.pk in used_user_ids,
            'cells': cells,
            'score': cells[0]['score'] if cells else None,
            'row_sum': row_sum.quantize(Decimal('0.01')) if has_any else None,
        })

    # แสดง Preview ที่ยังค้างอยู่เสมอ (แม้ออกไปหน้าอื่นแล้วกลับมา) จนกว่าจะกดยืนยันหรือยกเลิก
    score_import_preview = None
    pending_import = None
    pdata = request.session.get('score_import_preview')
    if pdata and pdata.get('data') and selected_main:
        if str(pdata.get('main_id')) == str(selected_main.pk) and pdata.get('exam_no') == exam_no:
            score_import_preview = pdata.get('data')
        else:
            pending_import = {
                'year': pdata.get('year', year),
                'main_id': pdata.get('main_id'),
                'exam_no': pdata.get('exam_no'),
                'main_label': pdata['data'].get('main_label', ''),
                'file_name': pdata['data'].get('file_name', ''),
            }

    all_year_members = []
    if score_import_preview and score_import_preview.get('has_resolvable_unmatched'):
        all_year_members = list(score_members_for_year(year)[0])

    return render(request, 'food/staff_scores.html', {
        'main_categories': main_categories,
        'selected_main': selected_main,
        'sub_cols': sub_cols,
        'categories': _leaf_score_categories(year),
        'category': category,
        'exam_no': exam_no,
        'rows': rows,
        'center': center,
        'q': q,
        'year': year,
        'is_open': selected_main.is_currently_open() if selected_main else False,
        'is_year_approved': is_year_approved,
        'score_import_preview': score_import_preview,
        'pending_import': pending_import,
        'all_year_members': all_year_members,
    })


def _compute_score_ranking(request, year, only_visible=True):
    members, center, q, _ = _score_members_for_year(request, year)
    qs = ScoreCategory.objects.filter(academic_year=year, parent__isnull=True).order_by('order', 'id')
    if only_visible:
        qs = qs.filter(is_visible=True)
    categories = list(qs)
    setting = SystemSetting.get()
    # จัดอันดับจากนักเรียนทั้งหมดในปี แล้วค่อยกรองตามศูนย์/คำค้น เพื่อให้อันดับไม่เพี้ยน
    all_members, _ = score_members_for_year(year)
    rows = compute_year_ranking(year, members=all_members, only_visible=only_visible)
    allowed = set(members.values_list('pk', flat=True))
    rows = [r for r in rows if r['user'].pk in allowed]
    return rows, categories, center, q, setting


def _export_all_scores_workbook(request, year):
    """Export คะแนนทุกหัวข้อใหญ่และทุกหัวข้อย่อยในปีการศึกษาที่ระบุเป็นไฟล์ Excel"""
    rows, categories, center, q, setting = _compute_score_ranking(request, year, only_visible=False)
    total_weight = sum((c.weight for c in categories), Decimal('0'))
    cat_subs = [
        (idx, c, list(c.subcategories.all().order_by('order', 'id')))
        for idx, c in enumerate(categories, 1)
    ]
    headers = ['อันดับ', 'รหัสค่าย 1', 'รหัสค่าย 2', 'ชื่อ-นามสกุล', 'โรงเรียน']
    for idx, c, subs in cat_subs:
        if subs:
            for sub_idx, sc in enumerate(subs, 1):
                headers.append(f'{idx}.{sub_idx} {sc.name} (เต็ม {sc.default_max_score})')
            headers.append(f'{idx}. รวม {c.name} (สัดส่วน {c.weight})')
        else:
            headers.append(f'{idx}. {c.name} (สัดส่วน {c.weight})')
    headers.extend([f'คะแนนรวม ({total_weight})', 'ผลการคัดเลือก'])

    data = []
    for r in rows:
        row_vals = [
            r['rank'],
            r['user'].camp1_code,
            r['user'].camp2_code or '',
            r['user'].get_full_name(),
            r['user'].school_name,
        ]
        for b, (_, c, subs) in zip(r['breakdown'], cat_subs):
            if subs:
                sub_map = {sb['category'].pk: sb for sb in b.get('sub_breakdown', [])}
                for sc in subs:
                    sb = sub_map.get(sc.pk)
                    row_vals.append(sb['weighted'] if sb and sb['weighted'] is not None else '')
            row_vals.append(b['weighted'] if b['weighted'] is not None else '')
        row_vals.extend([r['total'], 'ผ่านการคัดเลือก' if r['passed'] else 'ไม่ผ่าน'])
        data.append(row_vals)

    log_action(request.user, 'export', target_model='Score', new_value={'summary_all_year': year}, request=request)
    wb = build_workbook([(f'คะแนนทั้งหมด {year}'[:31], headers, data)])
    return workbook_response(wb, f'คะแนนทั้งหมด_หัวข้อใหญ่ย่อย_{year}.xlsx')


@staff_perm_required('staff_can_scores')
def score_summary(request):
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    rows, categories, center, q, setting = _compute_score_ranking(request, year, only_visible=False)
    if request.GET.get('export') == 'xlsx':
        return _export_all_scores_workbook(request, year)
    total_weight = sum((c.weight for c in categories), Decimal('0'))
    return render(request, 'food/score_summary.html', {
        'rows': rows, 'categories': categories, 'year': year, 'center': center, 'q': q,
        'total_weight': total_weight, 'setting': setting,
    })


def _update_announcement_settings_from_post(setting, post_data):
    """อัปเดตฟิลด์การตั้งค่าเอกสารประกาศผล (ข้อความ + กั้นหน้า-กั้นหลัง + ขนาดตัวอักษร + ความกว้างคอลัมน์) จาก POST"""
    from .announcement_pdf import to_arabic_num
    if 'announcement_school_name' in post_data:
        setting.announcement_school_name = to_arabic_num(post_data.get('announcement_school_name', '').strip())
    if 'announcement_title' in post_data:
        setting.announcement_title = to_arabic_num(post_data.get('announcement_title', '').strip())
    if 'announcement_center_name' in post_data:
        setting.announcement_center_name = to_arabic_num(post_data.get('announcement_center_name', '').strip())
    if 'announcement_body_p1' in post_data:
        setting.announcement_body_p1 = to_arabic_num(post_data.get('announcement_body_p1', '').strip())
    if 'announcement_body_p2' in post_data:
        setting.announcement_body_p2 = to_arabic_num(post_data.get('announcement_body_p2', '').strip())
    if 'announcement_date_text' in post_data:
        setting.announcement_date_text = to_arabic_num(post_data.get('announcement_date_text', '').strip())
    if 'announcement_director_name' in post_data:
        setting.announcement_director_name = to_arabic_num(post_data.get('announcement_director_name', '').strip())
    if 'announcement_director_position' in post_data:
        setting.announcement_director_position = to_arabic_num(post_data.get('announcement_director_position', '').strip())
    if 'announcement_list_title' in post_data:
        setting.announcement_list_title = to_arabic_num(post_data.get('announcement_list_title', '').strip())
    if 'announcement_note' in post_data:
        setting.announcement_note = to_arabic_num(post_data.get('announcement_note', '').strip())

    # ระยะขอบ (Margins) มม.
    if 'announcement_margin_top_mm' in post_data:
        setting.announcement_margin_top_mm = _parse_float(post_data.get('announcement_margin_top_mm'), 25.0, 5.0, 60.0)
    if 'announcement_margin_bottom_mm' in post_data:
        setting.announcement_margin_bottom_mm = _parse_float(post_data.get('announcement_margin_bottom_mm'), 25.0, 5.0, 60.0)
    if 'announcement_margin_left_mm' in post_data:
        setting.announcement_margin_left_mm = _parse_float(post_data.get('announcement_margin_left_mm'), 25.0, 5.0, 60.0)
    if 'announcement_margin_right_mm' in post_data:
        setting.announcement_margin_right_mm = _parse_float(post_data.get('announcement_margin_right_mm'), 25.0, 5.0, 60.0)

    # ขนาดตัวอักษรทุกส่วน ระยะบรรทัด ระยะย่อหน้า และความสูงตราครุฑ
    if 'announcement_font_size_title1' in post_data:
        setting.announcement_font_size_title1 = _parse_float(post_data.get('announcement_font_size_title1'), 17.0, 10.0, 28.0)
    if 'announcement_font_size_title2' in post_data:
        setting.announcement_font_size_title2 = _parse_float(post_data.get('announcement_font_size_title2'), 15.0, 10.0, 26.0)
    if 'announcement_font_size_title3' in post_data:
        setting.announcement_font_size_title3 = _parse_float(post_data.get('announcement_font_size_title3'), 14.0, 10.0, 24.0)
    if 'announcement_font_size_body' in post_data:
        setting.announcement_font_size_body = _parse_float(post_data.get('announcement_font_size_body'), 14.5, 9.0, 24.0)
    if 'announcement_font_size_director_name' in post_data:
        setting.announcement_font_size_director_name = _parse_float(post_data.get('announcement_font_size_director_name'), 14.5, 9.0, 24.0)
    if 'announcement_font_size_director_pos' in post_data:
        setting.announcement_font_size_director_pos = _parse_float(post_data.get('announcement_font_size_director_pos'), 13.0, 8.0, 22.0)
    if 'announcement_font_size_list_title1' in post_data:
        setting.announcement_font_size_list_title1 = _parse_float(post_data.get('announcement_font_size_list_title1'), 14.0, 9.0, 24.0)
    if 'announcement_font_size_list_title2' in post_data:
        setting.announcement_font_size_list_title2 = _parse_float(post_data.get('announcement_font_size_list_title2'), 12.0, 8.0, 22.0)
    if 'announcement_font_size_table' in post_data:
        setting.announcement_font_size_table = _parse_float(post_data.get('announcement_font_size_table'), 10.0, 7.0, 18.0)
    if 'announcement_line_spacing' in post_data:
        setting.announcement_line_spacing = _parse_float(post_data.get('announcement_line_spacing'), 1.85, 1.0, 3.0)
    if 'announcement_paragraph_spacing_mm' in post_data:
        setting.announcement_paragraph_spacing_mm = _parse_float(
            post_data.get('announcement_paragraph_spacing_mm'), 4.0, 0.0, 30.0
        )
    if 'announcement_indent_mm' in post_data:
        setting.announcement_indent_mm = _parse_float(post_data.get('announcement_indent_mm'), 25.0, 5.0, 60.0)
    if 'announcement_garuda_height_mm' in post_data:
        setting.announcement_garuda_height_mm = _parse_float(post_data.get('announcement_garuda_height_mm'), 30.0, 12.0, 50.0)

    # ความกว้างช่องตาราง (มม.)
    if 'announcement_col_w_no' in post_data:
        setting.announcement_col_w_no = _parse_float(post_data.get('announcement_col_w_no'), 14.0, 5.0, 40.0)
    if 'announcement_col_w_camp2' in post_data:
        setting.announcement_col_w_camp2 = _parse_float(post_data.get('announcement_col_w_camp2'), 24.0, 10.0, 60.0)
    if 'announcement_col_w_rights' in post_data:
        setting.announcement_col_w_rights = _parse_float(post_data.get('announcement_col_w_rights'), 24.0, 10.0, 60.0)
    if 'announcement_col_w_name' in post_data:
        setting.announcement_col_w_name = _parse_float(post_data.get('announcement_col_w_name'), 50.0, 20.0, 120.0)
    if 'announcement_col_w_school' in post_data:
        setting.announcement_col_w_school = _parse_float(post_data.get('announcement_col_w_school'), 54.0, 20.0, 120.0)
    if 'announcement_col_w_grade' in post_data:
        setting.announcement_col_w_grade = _parse_float(post_data.get('announcement_col_w_grade'), 14.0, 5.0, 40.0)


def _process_qualification_post(request, action, year, setting, approval, is_approved):
    """ประมวลผลคำสั่งอนุมัติประกาศผลและกำหนดค่าเอกสารประกาศผล (ใช้ร่วมกันทั้ง score_announcement และ score_categories)"""
    if action == 'save_qualification':
        if is_approved:
            if 'max_qualified_count' in request.POST:
                new_mq = max(0, int(request.POST.get('max_qualified_count') or 0))
                if new_mq != approval.max_qualified_count:
                    raise FoodError(
                        f'ปีการศึกษา {year} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนจำนวนผู้ผ่านการคัดเลือกได้ '
                        f'กรุณายกเลิกประกาศผลก่อน'
                    )
            if 'passing_score' in request.POST:
                new_ps = to_decimal(request.POST.get('passing_score')) or Decimal('0')
                if new_ps != approval.passing_score:
                    raise FoodError(
                        f'ปีการศึกษา {year} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนคะแนนเกณฑ์ผ่านได้ '
                        f'กรุณายกเลิกประกาศผลก่อน'
                    )
        else:
            setting.passing_score = to_decimal(request.POST.get('passing_score')) or Decimal('0')
            setting.max_qualified_count = max(0, int(request.POST.get('max_qualified_count') or 0))
        _update_announcement_settings_from_post(setting, request.POST)
        setting.save()
        log_action(
            request.user, 'update_settings', target=setting,
            new_value={
                'passing_score': str(setting.passing_score),
                'max_qualified': setting.max_qualified_count,
                'announcement_title': setting.announcement_title,
                'margin_l': setting.announcement_margin_left_mm,
                'margin_r': setting.announcement_margin_right_mm,
            },
            request=request,
        )
        messages.success(request, 'บันทึกเกณฑ์ผู้ผ่านการคัดเลือกและเอกสารประกาศผลเรียบร้อย')
    elif action == 'approve_scores':
        if is_approved:
            messages.info(request, f'ปีการศึกษา {year} อนุมัติประกาศผลแล้ว หากต้องการคำนวณผลใหม่ กรุณายกเลิกประกาศผลก่อน')
        else:
            appr, ranked = approve_year_scores(year, request.user)
            passed_n = sum(1 for r in ranked if r['passed'])
            log_action(
                request.user, 'update_settings', target=appr,
                new_value={'approve_scores': year, 'total': len(ranked), 'passed': passed_n},
                request=request,
            )
            messages.success(
                request,
                f'อนุมัติประกาศผลคะแนนปีการศึกษา {year} แล้ว '
                f'(ทั้งหมด {len(ranked)} คน ผ่านการคัดเลือกเข้าค่าย 2 จำนวน {passed_n} คน)',
            )
    elif action == 'unapprove_scores':
        if not is_approved:
            messages.info(request, f'ปีการศึกษา {year} ยังไม่ได้อนุมัติประกาศผล')
        else:
            confirm_password = request.POST.get('confirm_password') or ''
            if not confirm_password:
                raise FoodError('กรุณากรอกรหัสผ่านปัจจุบันของผู้ดูแลระบบเพื่อยืนยันการยกเลิกประกาศผล')
            if not request.user.check_password(confirm_password):
                log_action(
                    request.user, 'update_settings', target_model='ScoreApproval',
                    new_value={'unapprove_scores_failed': year, 'reason': 'invalid_password'},
                    request=request,
                )
                raise FoodError('รหัสผ่านปัจจุบันของผู้ดูแลระบบไม่ถูกต้อง ไม่สามารถยกเลิกประกาศผลได้')
            unapprove_year_scores(year)
            log_action(
                request.user, 'update_settings', target_model='ScoreApproval',
                new_value={'unapprove_scores': year, 'password_verified': True},
                request=request,
            )
            messages.success(
                request,
                f'ยกเลิกการอนุมัติประกาศผลคะแนนปีการศึกษา {year} แล้ว '
                f'(ปลดล็อกให้บันทึกเวลาและคะแนนได้อีกครั้ง)',
            )


@xframe_options_sameorigin
@staff_perm_required('staff_can_scores')
def score_announcement(request):
    """เอกสารประกาศผลการคัดเลือก (พร้อมพิมพ์ A4, Export PDF, Export Excel และแท็บอนุมัติประกาศผล)"""
    from .announcement_pdf import (
        build_announcement_pdf, format_announcement_date_text,
        format_announcement_list_title, format_default_announcement_date,
        format_announcement_body_p1, format_announcement_body_p2,
        format_default_announcement_p1, format_default_announcement_p2,
        format_default_announcement_p2_template,
        get_announcement_html_lines, get_announcement_page1_metrics,
    )

    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()

    tab = (request.GET.get('tab') or request.POST.get('tab') or 'document').strip()
    if tab not in ('document', 'qualification') or not request.user.is_admin_role:
        tab = 'document'

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action in ('save_qualification', 'approve_scores', 'unapprove_scores'):
            if not request.user.is_admin_role:
                return HttpResponseForbidden('เฉพาะผู้ดูแลระบบเท่านั้น')
            setting = SystemSetting.get()
            approval_obj = ScoreApproval.objects.filter(academic_year=year).select_related('approved_by').first()
            is_approved_flag = bool(approval_obj and approval_obj.is_approved)
            try:
                _process_qualification_post(request, action, year, setting, approval_obj, is_approved_flag)
            except (FoodError, ValueError) as e:
                messages.error(request, str(e))
            return redirect(f"{reverse('food:score_announcement')}?year={year}&tab=qualification")

        if action == 'save_announcement_text':
            if not request.user.is_admin_role:
                return HttpResponseForbidden('เฉพาะผู้ดูแลระบบเท่านั้น')
            setting = SystemSetting.get()
            _update_announcement_settings_from_post(setting, request.POST)
            setting.save()
            log_action(request.user, 'update_settings', target=setting, new_value={
                'announcement_school_name': setting.announcement_school_name,
                'announcement_title': setting.announcement_title,
                'announcement_center_name': setting.announcement_center_name,
                'announcement_body_p1': setting.announcement_body_p1,
                'announcement_body_p2': setting.announcement_body_p2,
                'announcement_date_text': setting.announcement_date_text,
                'announcement_director_name': setting.announcement_director_name,
                'announcement_director_position': setting.announcement_director_position,
                'announcement_list_title': setting.announcement_list_title,
                'announcement_note': setting.announcement_note,
                'margin_l': setting.announcement_margin_left_mm,
                'margin_r': setting.announcement_margin_right_mm,
                'font_body': setting.announcement_font_size_body,
                'font_table': setting.announcement_font_size_table,
            }, request=request)
            messages.success(request, 'บันทึกข้อมูลและรูปแบบประกาศผลเรียบร้อยแล้ว')
            params = {'year': year}
            if request.POST.get('only') in ('passed', 'all'):
                params['only'] = request.POST.get('only')
            if request.POST.get('q', '').strip():
                params['q'] = request.POST.get('q', '').strip()
            return redirect(f"{reverse('food:score_announcement')}?{urlencode(params)}")

    approval = ScoreApproval.objects.filter(academic_year=year, is_approved=True).select_related('approved_by').first()
    only_passed = request.GET.get('only', 'passed') == 'passed'
    frozen_qrs = list(
        QualificationResult.objects.filter(academic_year=year)
        .select_related('user')
        .order_by('rank')
    ) if approval else []
    if approval and frozen_qrs:
        setting = SystemSetting.get()
        members, center, q, _ = _score_members_for_year(request, year)
        allowed_ids = set(members.values_list('pk', flat=True))
        all_ranked = [
            {'user': qr.user, 'total': qr.total, 'rank': qr.rank, 'passed': qr.passed, 'breakdown': []}
            for qr in frozen_qrs if qr.user.is_active
        ]
        rows = [r for r in all_ranked if r['user'].pk in allowed_ids]
    else:
        rows, categories, center, q, setting = _compute_score_ranking(request, year)
        all_members, _ = score_members_for_year(year)
        all_ranked = compute_year_ranking(year, members=all_members, only_visible=True)

    display_rows = [r for r in rows if r['passed']] if only_passed else list(rows)

    # เรียงลำดับจากรหัสนักเรียนปัจจุบัน (น้อยไปมาก) ตามด้วยรหัสนักเรียนใช้สิทธิ์ (น้อยไปมาก)
    used_records = {ur.user_id: ur.rights_code for ur in UseRecord.objects.filter(academic_year=year, is_used=True)}
    all_passed_sorted = sort_announcement_rows([r for r in all_ranked if r['passed']], year, used_records=used_records)
    preview_km_map = {
        r['user'].pk: (r['user'].camp2_code or f'{str(year)[-2:]}KM{idx:02d}')
        for idx, r in enumerate(all_passed_sorted, 1)
    }

    display_rows = sort_announcement_rows(display_rows, year, used_records=used_records)

    # เติมข้อมูลรหัสค่าย 2, รหัสเดิม/สิทธิ์, ระดับชั้น สำหรับเอกสารประกาศผล
    for idx, r in enumerate(display_rows, 1):
        u = r['user']
        r['camp2_code'] = (u.camp2_code or preview_km_map.get(u.pk) or f'{str(year)[-2:]}KM{idx:02d}') if r['passed'] else (u.camp2_code or '-')
        r['rights_code'] = used_records.get(u.pk) or u.camp1_code
        r['grade'] = u.level_for_year(year) or u.current_level or ''

    if request.GET.get('export') == 'pdf':
        pdf_bytes = build_announcement_pdf(display_rows, year, setting)
        resp = HttpResponse(pdf_bytes, content_type='application/pdf')
        resp['Content-Disposition'] = f'inline; filename="announcement_{year}.pdf"'
        resp['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        resp['Pragma'] = 'no-cache'
        return resp

    if request.GET.get('export') == 'docx':
        from .announcement_docx import build_announcement_docx, HAS_DOCX
        if not HAS_DOCX:
            messages.error(request, 'ไม่พบ python-docx กรุณาติดตั้ง: pip install python-docx')
            return redirect(request.path + '?' + urlencode({'year': year, 'only': 'passed' if only_passed else 'all'}))
        docx_bytes = build_announcement_docx(display_rows, year, setting)
        fname = f'ประกาศผล_ค่าย2_ปีการศึกษา{year}.docx'
        resp = HttpResponse(docx_bytes,
                            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        resp['Content-Disposition'] = f"attachment; filename=\"announcement_{year}.docx\"; filename*=UTF-8''{quote(fname)}"
        return resp

    if request.GET.get('export') == 'xlsx':
        headers = ['ลำดับ', 'เลขประจำตัวค่าย 2', 'เลขประจำตัวค่าย 1', 'ชื่อ นามสกุล', 'โรงเรียน', 'ระดับ']
        data = [[idx, r['camp2_code'], r['rights_code'], r['user'].get_full_name(),
                 r['user'].school_name, r['grade']]
                for idx, r in enumerate(display_rows, 1)]
        wb = build_workbook([(f'ประกาศ_{year}', headers, data)])
        return workbook_response(wb, f'ประกาศ_รายชื่อค่าย2-{str(year)[-2:]}.xlsx')

    from .certificate_pdf import get_certificate_data_for_students
    cert_students_data = get_certificate_data_for_students(year)
    cert_eligible_students = [s for s in cert_students_data if s['is_eligible']]
    cert_lecturers = list(CertificateRecipient.objects.filter(
        academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER
    ))
    cert_committees = list(CertificateRecipient.objects.filter(
        academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE
    ))

    default_list_title = f'ประกาศรายชื่อผู้ที่ผ่านการคัดเลือกเข้า สอวน.คอมพิวเตอร์ ค่ายที่ 2 ปีการศึกษา {year}'
    passed_count = sum(1 for r in rows if r['passed'])
    doc_row_count = len(display_rows)
    p1_lines, p2_lines = get_announcement_html_lines(setting, doc_row_count)

    total_col_w = (
        (setting.announcement_col_w_no or 14.0) +
        (setting.announcement_col_w_camp2 or 24.0) +
        (setting.announcement_col_w_rights or 24.0) +
        (setting.announcement_col_w_name or 50.0) +
        (setting.announcement_col_w_school or 54.0) +
        (setting.announcement_col_w_grade or 14.0)
    )
    if total_col_w <= 0:
        total_col_w = 176.0
    col_w_pct = {
        'no': round(((setting.announcement_col_w_no or 14.0) / total_col_w) * 100, 2),
        'camp2': round(((setting.announcement_col_w_camp2 or 24.0) / total_col_w) * 100, 2),
        'rights': round(((setting.announcement_col_w_rights or 24.0) / total_col_w) * 100, 2),
        'name': round(((setting.announcement_col_w_name or 50.0) / total_col_w) * 100, 2),
        'school': round(((setting.announcement_col_w_school or 54.0) / total_col_w) * 100, 2),
        'grade': round(((setting.announcement_col_w_grade or 14.0) / total_col_w) * 100, 2),
    }

    from .announcement_pdf import to_arabic_num
    page1_metrics = get_announcement_page1_metrics(setting)
    return render(request, 'food/score_announcement.html', {
        'rows': display_rows, 'all_count': len(rows),
        'passed_count': passed_count,
        'doc_row_count': doc_row_count,
        'approval': approval,
        'is_approved': bool(approval),
        'is_year_approved': bool(approval),
        'tab': tab,
        'effective_max_qualified': approval.max_qualified_count if approval else setting.max_qualified_count,
        'year': year, 'center': center, 'q': q, 'only_passed': only_passed, 'setting': setting,
        'col_w_pct': col_w_pct, 'total_col_w': total_col_w,
        'date_indent_mm': page1_metrics['date_indent_mm'],
        'sig_box_w_mm': page1_metrics['sig_box_w_mm'],
        'sig_left_from_margin_mm': page1_metrics['sig_left_from_margin_mm'],
        'announcement_school_name_display': to_arabic_num(setting.announcement_school_name),
        'announcement_title_display': to_arabic_num(setting.announcement_title),
        'announcement_center_name_display': to_arabic_num(setting.announcement_center_name),
        'announcement_director_name_display': to_arabic_num(setting.announcement_director_name),
        'announcement_note_display': to_arabic_num(setting.announcement_note),
        'announcement_date_display': format_announcement_date_text(setting),
        'default_announcement_date': format_default_announcement_date(),
        'announcement_list_title_display': format_announcement_list_title(setting, year),
        'default_announcement_list_title': default_list_title,
        'announcement_body_p1_display': format_announcement_body_p1(setting),
        'announcement_body_p2_display': format_announcement_body_p2(setting, doc_row_count),
        'p1_lines': p1_lines,
        'p2_lines': p2_lines,
        'default_announcement_p1': format_default_announcement_p1(setting),
        'default_announcement_p2': format_default_announcement_p2(doc_row_count),
        'default_announcement_p2_template': format_default_announcement_p2_template(),
        'edit_announcement_body_p1': to_arabic_num(setting.announcement_body_p1) or format_default_announcement_p1(setting),
        'edit_announcement_body_p2': to_arabic_num(setting.announcement_body_p2) or format_default_announcement_p2_template(),
        'edit_announcement_date_text': to_arabic_num(setting.announcement_date_text) or format_default_announcement_date(),
        'edit_announcement_list_title': to_arabic_num(setting.announcement_list_title) or default_list_title,
        'edit_announcement_director_position': to_arabic_num(setting.announcement_director_position),
        'director_position_lines': [to_arabic_num(l.strip()) for l in (setting.announcement_director_position or '').splitlines() if l.strip()],
        'cert_eligible_students': cert_eligible_students,
        'cert_lecturers': cert_lecturers,
        'cert_committees': cert_committees,
        'cert_eligible_count': len(cert_eligible_students),
        'cert_lecturer_count': len(cert_lecturers),
        'cert_committee_count': len(cert_committees),
        'cert_total_count': len(cert_eligible_students) + len(cert_lecturers) + len(cert_committees),
    })


def _parse_float(raw, default, lo, hi):
    """แปลงค่าทศนิยมสำหรับตำแหน่ง/ขนาดตัวอักษรพร้อมจำกัดช่วง [lo, hi]"""
    try:
        val = float(str(raw).strip())
    except (TypeError, ValueError):
        return float(default)
    return max(float(lo), min(float(hi), val))


def _parse_hex_color(raw, default='#000000'):
    val = str(raw or '').strip()
    return val if re.match(r'^#[0-9a-fA-F]{6}$', val) else default


def _validated_image_upload(uploaded):
    """ตรวจสอบว่าไฟล์ที่อัปโหลดเป็นรูปภาพจริง (Pillow verify) แล้วรีเซ็ตตัวชี้ไฟล์"""
    from PIL import Image
    if not uploaded:
        return None, 'กรุณาเลือกไฟล์รูปภาพ'
    try:
        uploaded.seek(0)
        img = Image.open(uploaded)
        img.verify()
        uploaded.seek(0)
        return uploaded, None
    except Exception:
        return None, 'ไฟล์รูปภาพไม่ถูกต้อง กรุณาอัปโหลดไฟล์ PNG หรือ JPG'


def _normalize_cert_desc(text):
    """ปรับรูปแบบขึ้นบรรทัดใหม่/ช่องว่างท้ายบรรทัดเพื่อเทียบกับข้อความเริ่มต้นของกลุ่ม"""
    raw = str(text or '').replace('\r\n', '\n').replace('\r', '\n').strip()
    return '\n'.join(ln.rstrip() for ln in raw.splitlines())


def _resolve_recipient_description(raw_desc, rtype, cert_setting, year):
    """คืนค่า '' เมื่อข้อความว่างหรือตรงกับค่าเริ่มต้นของกลุ่ม เพื่อให้อัปเดตตามกลุ่มได้"""
    description = (raw_desc or '').strip()
    default_raw = cert_setting.default_line3_for(rtype)
    default_rendered = cert_setting.render_text(default_raw, year)
    norm_desc = _normalize_cert_desc(description)
    if not norm_desc or norm_desc in (_normalize_cert_desc(default_raw), _normalize_cert_desc(default_rendered)):
        return ''
    return description


@staff_perm_required('staff_can_scores')
def admin_certificates(request):
    """หน้าจัดการเกียรติบัตร (นักเรียน / วิทยากร / คณะกรรมการ / ตั้งค่าเกียรติบัตร)"""
    from django.urls import reverse
    from .certificate_pdf import get_certificate_data_for_students, parse_recipients_upload

    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()

    tab = (request.GET.get('tab') or request.POST.get('tab') or 'students').strip()
    if tab not in ('students', 'lecturer', 'committee', 'settings'):
        tab = 'students'
    if request.method != 'POST' and tab == 'settings' and not request.user.is_admin_role:
        messages.warning(request, 'เฉพาะผู้ดูแลระบบเท่านั้นที่สามารถตั้งค่าเกียรติบัตรได้')
        return redirect(reverse('food:admin_certificates') + f'?year={year}&tab=students')

    setting = SystemSetting.get()
    cert_setting = CertificateSetting.get()

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()

        if action == 'toggle_cert_publish':
            if not request.user.is_admin_role:
                return HttpResponseForbidden('เฉพาะผู้ดูแลระบบเท่านั้น')
            cert_status, _ = CertificateYearStatus.objects.get_or_create(academic_year=year)
            old_pub = cert_status.is_published
            cert_status.is_published = not old_pub
            cert_status.updated_by = request.user
            cert_status.save(update_fields=['is_published', 'updated_by', 'updated_at'])
            log_action(
                request.user, 'toggle_cert_publish', target=cert_status,
                old_value={'academic_year': year, 'is_published': old_pub},
                new_value={'academic_year': year, 'is_published': cert_status.is_published},
                request=request,
            )
            if cert_status.is_published:
                messages.success(request, f'เปิดแสดงเกียรติบัตรให้นักเรียนประจำปีการศึกษา {year} เรียบร้อยแล้ว')
            else:
                messages.info(request, f'ปิดการแสดงเกียรติบัตรให้นักเรียนประจำปีการศึกษา {year} เรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={tab}')

        if action == 'save_settings':
            if not request.user.is_admin_role:
                return HttpResponseForbidden('เฉพาะผู้ดูแลระบบเท่านั้น')
            if 'background_image' in request.FILES:
                from composn.image_utils import ImageUploadError, compress_certificate_bg_image
                try:
                    compressed_bg = compress_certificate_bg_image(request.FILES['background_image'])
                except ImageUploadError as img_err:
                    messages.error(request, str(img_err))
                    return redirect(reverse('food:admin_certificates') + f'?year={year}&tab=settings')
                if cert_setting.background_image:
                    try:
                        cert_setting.background_image.delete(save=False)
                    except Exception:
                        pass
                cert_setting.background_image = compressed_bg
            elif request.POST.get('clear_background_image') == '1':
                if cert_setting.background_image:
                    try:
                        cert_setting.background_image.delete(save=False)
                    except Exception:
                        pass
                cert_setting.background_image = None

            if 'header1_text' in request.POST:
                cert_setting.header1_text = request.POST.get('header1_text', '').strip()
            cert_setting.header1_y = _parse_float(request.POST.get('header1_y'), cert_setting.header1_y, 0, 210)
            cert_setting.header1_font_size = _parse_float(
                request.POST.get('header1_font_size'), cert_setting.header1_font_size, 8, 96
            )
            cert_setting.header1_line_spacing = _parse_float(
                request.POST.get('header1_line_spacing'), cert_setting.header1_line_spacing, 4, 30
            )
            if 'has_header_fields' in request.POST or 'header1_bold' in request.POST:
                cert_setting.header1_bold = request.POST.get('header1_bold') == '1'
            cert_setting.header1_color = _parse_hex_color(request.POST.get('header1_color'), cert_setting.header1_color)

            if 'header2_text' in request.POST:
                cert_setting.header2_text = request.POST.get('header2_text', '').strip()
            cert_setting.header2_y = _parse_float(request.POST.get('header2_y'), cert_setting.header2_y, 0, 210)
            cert_setting.header2_font_size = _parse_float(
                request.POST.get('header2_font_size'), cert_setting.header2_font_size, 8, 96
            )
            if 'has_header_fields' in request.POST or 'header2_bold' in request.POST:
                cert_setting.header2_bold = request.POST.get('header2_bold') == '1'
            cert_setting.header2_color = _parse_hex_color(request.POST.get('header2_color'), cert_setting.header2_color)

            cert_setting.line1_y = _parse_float(request.POST.get('line1_y'), cert_setting.line1_y, 0, 210)
            cert_setting.line1_font_size = _parse_float(request.POST.get('line1_font_size'), cert_setting.line1_font_size, 8, 96)
            cert_setting.line1_bold = request.POST.get('line1_bold') == '1'
            cert_setting.line1_color = _parse_hex_color(request.POST.get('line1_color'), cert_setting.line1_color)

            cert_setting.line2_y = _parse_float(request.POST.get('line2_y'), cert_setting.line2_y, 0, 210)
            cert_setting.line2_font_size = _parse_float(request.POST.get('line2_font_size'), cert_setting.line2_font_size, 8, 96)
            cert_setting.line2_bold = request.POST.get('line2_bold') == '1'
            cert_setting.line2_color = _parse_hex_color(request.POST.get('line2_color'), cert_setting.line2_color)
            if 'student_default_province' in request.POST:
                cert_setting.student_default_province = request.POST.get('student_default_province', '').strip()

            cert_setting.line3_y = _parse_float(request.POST.get('line3_y'), cert_setting.line3_y, 0, 210)
            cert_setting.line3_font_size = _parse_float(request.POST.get('line3_font_size'), cert_setting.line3_font_size, 8, 96)
            cert_setting.line3_bold = request.POST.get('line3_bold') == '1'
            cert_setting.line3_color = _parse_hex_color(request.POST.get('line3_color'), cert_setting.line3_color)
            cert_setting.line3_line_spacing = _parse_float(
                request.POST.get('line3_line_spacing'), cert_setting.line3_line_spacing, 4, 30
            )
            if 'student_line3_text' in request.POST:
                cert_setting.student_line3_text = request.POST.get('student_line3_text', '').strip()
            if 'lecturer_line3_text' in request.POST:
                cert_setting.lecturer_line3_text = request.POST.get('lecturer_line3_text', '').strip()
            if 'committee_line3_text' in request.POST:
                cert_setting.committee_line3_text = request.POST.get('committee_line3_text', '').strip()

            if 'line4_text' in request.POST:
                cert_setting.line4_text = request.POST.get('line4_text', '').strip()
            cert_setting.line4_y = _parse_float(request.POST.get('line4_y'), cert_setting.line4_y, 0, 210)
            cert_setting.line4_font_size = _parse_float(request.POST.get('line4_font_size'), cert_setting.line4_font_size, 8, 96)
            cert_setting.line4_bold = request.POST.get('line4_bold') == '1'
            cert_setting.line4_color = _parse_hex_color(request.POST.get('line4_color'), cert_setting.line4_color)

            if 'has_qr_fields' in request.POST or 'show_qr_verify' in request.POST:
                cert_setting.show_qr_verify = request.POST.get('show_qr_verify') == '1'
            cert_setting.qr_x = _parse_float(request.POST.get('qr_x'), cert_setting.qr_x, 5, 290)
            cert_setting.qr_y = _parse_float(request.POST.get('qr_y'), cert_setting.qr_y, 5, 205)
            cert_setting.qr_size = _parse_float(request.POST.get('qr_size'), cert_setting.qr_size, 10, 60)
            cert_setting.ref_code_font_size = _parse_float(
                request.POST.get('ref_code_font_size'), cert_setting.ref_code_font_size, 6, 24
            )
            cert_setting.ref_code_color = _parse_hex_color(request.POST.get('ref_code_color'), cert_setting.ref_code_color)

            cert_setting.save()
            log_action(request.user, 'update_settings', target=cert_setting, new_value='บันทึกตั้งค่าเกียรติบัตร', request=request)
            messages.success(request, 'บันทึกการตั้งค่าเกียรติบัตรเรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab=settings')

        if action == 'reset_settings':
            if not request.user.is_admin_role:
                return HttpResponseForbidden('เฉพาะผู้ดูแลระบบเท่านั้น')
            cert_setting.reset_layout_defaults()
            cert_setting.save()
            log_action(request.user, 'update_settings', target=cert_setting, new_value='รีเซ็ตตำแหน่งและขนาดตัวอักษรเกียรติบัตร', request=request)
            messages.success(request, 'รีเซ็ตตำแหน่งแกน Y และขนาดตัวอักษรกลับค่าเริ่มต้นเรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab=settings')

        if action == 'add_recipient':
            rtype = request.POST.get('recipient_type')
            if rtype not in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
                rtype = CertificateRecipient.TYPE_LECTURER
            full_name = (request.POST.get('full_name') or '').strip()
            school_province = (request.POST.get('school_province') or '').strip()
            description = _resolve_recipient_description(request.POST.get('description'), rtype, cert_setting, year)
            if not full_name:
                messages.error(request, 'กรุณาระบุชื่อ - นามสกุล')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rtype}')
            rec = CertificateRecipient.objects.create(
                academic_year=year,
                recipient_type=rtype,
                full_name=full_name[:200],
                school_province=school_province[:255],
                description=description,
                created_by=request.user,
            )
            log_action(request.user, 'update_settings', target=rec, new_value={'add_recipient': rec.full_name, 'type': rtype}, request=request)
            messages.success(request, f'เพิ่มรายชื่อ "{rec.full_name}" เรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rtype}')

        if action == 'edit_recipient':
            rec = get_object_or_404(CertificateRecipient, pk=request.POST.get('recipient_id'))
            rtype = request.POST.get('recipient_type', rec.recipient_type)
            if rtype not in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
                rtype = rec.recipient_type
            full_name = (request.POST.get('full_name') or '').strip()
            if not full_name:
                messages.error(request, 'กรุณาระบุชื่อ - นามสกุล')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rec.recipient_type}')
            rec.recipient_type = rtype
            rec.full_name = full_name[:200]
            rec.school_province = (request.POST.get('school_province') or '').strip()[:255]
            rec.description = _resolve_recipient_description(request.POST.get('description'), rtype, cert_setting, year)
            rec.save()
            log_action(request.user, 'update_settings', target=rec, new_value={'edit_recipient': rec.full_name, 'type': rtype}, request=request)
            messages.success(request, f'แก้ไขข้อมูล "{rec.full_name}" เรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rtype}')

        if action == 'delete_recipient':
            rec = get_object_or_404(CertificateRecipient, pk=request.POST.get('recipient_id'))
            rtype = rec.recipient_type
            rname = rec.full_name
            rec.delete()
            log_action(request.user, 'update_settings', target_model='CertificateRecipient', new_value={'delete': rname, 'type': rtype}, request=request)
            messages.success(request, f'ลบรายชื่อ "{rname}" เรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rtype}')

        if action == 'delete_all_recipients':
            rtype = request.POST.get('recipient_type')
            if rtype not in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
                rtype = CertificateRecipient.TYPE_LECTURER
            cnt, _ = CertificateRecipient.objects.filter(academic_year=year, recipient_type=rtype).delete()
            log_action(request.user, 'update_settings', target_model='CertificateRecipient',
                       new_value={'delete_all': cnt, 'type': rtype, 'year': year}, request=request)
            messages.success(request, f'ลบรายชื่อทั้งหมด {cnt} รายการเรียบร้อยแล้ว')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={rtype}')

        if action == 'import_recipients':
            default_type = request.POST.get('default_type')
            if default_type not in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
                default_type = CertificateRecipient.TYPE_LECTURER
            up_file = request.FILES.get('import_file')
            if not up_file:
                messages.error(request, 'กรุณาเลือกไฟล์ Excel (.xlsx) หรือ CSV (.csv)')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={default_type}')
            parsed_rows, parse_errors = parse_recipients_upload(up_file, default_type=default_type)
            if not parsed_rows:
                messages.error(request, 'ไม่สามารถนำเข้าข้อมูลได้: ' + '; '.join(parse_errors[:5]))
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={default_type}')
            with transaction.atomic():
                objs = [
                    CertificateRecipient(
                        academic_year=year,
                        recipient_type=row['recipient_type'],
                        full_name=row['full_name'],
                        school_province=row['school_province'],
                        description=row['description'],
                        sort_order=idx,
                        created_by=request.user,
                    )
                    for idx, row in enumerate(parsed_rows, 1)
                ]
                CertificateRecipient.objects.bulk_create(objs)
            log_action(request.user, 'import', target_model='CertificateRecipient',
                       new_value={'year': year, 'imported': len(objs), 'default_type': default_type}, request=request)
            if parse_errors:
                messages.warning(request, f'นำเข้าสำเร็จ {len(objs)} รายการ (พบคำเตือน {len(parse_errors)} แถว: {"; ".join(parse_errors[:3])})')
            else:
                messages.success(request, f'นำเข้ารายชื่อสำเร็จ {len(objs)} รายการ')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={default_type}')

        if action == 'transfer_recipients':
            back_tab = 'committee' if request.POST.get('tab') == 'committee' else 'lecturer'
            try:
                source_year = int((request.POST.get('source_year') or '').strip())
            except (ValueError, TypeError):
                messages.error(request, 'กรุณาเลือกปีการศึกษาต้นทางที่ต้องการโอนรายชื่อ')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={back_tab}')
            if source_year == year:
                messages.error(request, f'ปีต้นทางต้องไม่ใช่ปีเดียวกับปีปลายทาง ({year})')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={back_tab}')

            transfer_scope = (request.POST.get('transfer_scope') or back_tab).strip()
            if transfer_scope not in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE, 'both'):
                transfer_scope = back_tab
            skip_duplicates = request.POST.get('skip_duplicates') == '1'
            reset_description = request.POST.get('reset_description') == '1'

            src_qs = CertificateRecipient.objects.filter(academic_year=source_year)
            if transfer_scope in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
                src_qs = src_qs.filter(recipient_type=transfer_scope)
            src_list = list(src_qs.order_by('recipient_type', 'sort_order', 'id'))
            if not src_list:
                messages.warning(request, f'ไม่พบรายชื่อในกลุ่มที่เลือกของปีการศึกษา {source_year}')
                return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={back_tab}')

            existing_keys = set()
            if skip_duplicates:
                existing_keys = {
                    (r.recipient_type, ' '.join((r.full_name or '').split()).lower())
                    for r in CertificateRecipient.objects.filter(academic_year=year)
                }

            order_counters = {
                CertificateRecipient.TYPE_LECTURER: CertificateRecipient.objects.filter(
                    academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER
                ).count(),
                CertificateRecipient.TYPE_COMMITTEE: CertificateRecipient.objects.filter(
                    academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE
                ).count(),
            }

            new_objs = []
            skipped = 0
            with transaction.atomic():
                for src in src_list:
                    norm_key = (src.recipient_type, ' '.join((src.full_name or '').split()).lower())
                    if skip_duplicates and norm_key in existing_keys:
                        skipped += 1
                        continue
                    existing_keys.add(norm_key)
                    order_counters[src.recipient_type] = order_counters.get(src.recipient_type, 0) + 1
                    new_objs.append(
                        CertificateRecipient(
                            academic_year=year,
                            recipient_type=src.recipient_type,
                            full_name=src.full_name,
                            school_province=src.school_province,
                            description='' if reset_description else src.description,
                            sort_order=order_counters[src.recipient_type],
                            created_by=request.user,
                        )
                    )
                if new_objs:
                    CertificateRecipient.objects.bulk_create(new_objs)

            log_action(
                request.user, 'import', target_model='CertificateRecipient',
                new_value={
                    'transfer_from': source_year, 'to_year': year, 'scope': transfer_scope,
                    'created': len(new_objs), 'skipped': skipped,
                },
                request=request,
            )
            if new_objs:
                msg = f'โอนรายชื่อจากปีการศึกษา {source_year} มายังปี {year} สำเร็จ {len(new_objs)} รายการ'
                if skipped:
                    msg += f' (ข้ามรายชื่อซ้ำ {skipped} รายการ)'
                messages.success(request, msg)
            else:
                messages.info(request, f'รายชื่อทั้งหมดจากปีการศึกษา {source_year} มีอยู่แล้วในปี {year} (ข้าม {skipped} รายการ)')
            return redirect(reverse('food:admin_certificates') + f'?year={year}&tab={back_tab}')

    # ข้อมูลแท็บ 1: นักเรียน (จัดกลุ่มตามสิทธิ์รับเกียรติบัตร)
    all_students_data = get_certificate_data_for_students(year)
    q = (request.GET.get('q') or '').strip()
    status_filter = request.GET.get('status') or 'all'
    if status_filter not in ('all', 'eligible', 'ineligible', 'rights'):
        status_filter = 'all'

    q_filtered = all_students_data
    if q:
        q_low = q.lower()
        q_filtered = [
            s for s in q_filtered
            if q_low in s['user'].camp1_code.lower()
            or q_low in (s['user'].get_full_name() or '').lower()
            or q_low in (s['user'].school_name or '').lower()
        ]

    eligible_students = [s for s in q_filtered if s['is_current'] and s['is_eligible']]
    ineligible_students = [s for s in q_filtered if s['is_current'] and not s['is_eligible']]
    rights_students = [s for s in q_filtered if not s['is_current']]

    if status_filter == 'eligible':
        filtered_data = eligible_students
    elif status_filter == 'ineligible':
        filtered_data = ineligible_students
    elif status_filter == 'rights':
        filtered_data = rights_students
    else:
        filtered_data = eligible_students + ineligible_students + rights_students

    total_current = sum(1 for s in all_students_data if s['is_current'])
    eligible_count = sum(1 for s in all_students_data if s['is_current'] and s['is_eligible'])
    ineligible_count = total_current - eligible_count
    rights_count = len(all_students_data) - total_current
    all_students_count = len(all_students_data)

    if request.GET.get('export') == 'xlsx':
        sheets = []
        if status_filter in ('all', 'eligible') and (eligible_students or status_filter == 'eligible'):
            sheets.append((
                'ได้รับเกียรติบัตร',
                ['ลำดับ', 'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'โรงเรียน', 'ขาดอบรม (ภาคเวลา)', 'การเข้าสอบ'],
                [
                    [
                        idx,
                        s['user'].camp1_code,
                        s['user'].get_full_name(),
                        s['user'].school_name,
                        f"{s['eligibility']['absences']}/{s['eligibility']['max_absences']}",
                        f"{s['eligibility']['attended_exam_slots']}/{s['eligibility']['total_exam_slots']}",
                    ]
                    for idx, s in enumerate(eligible_students, 1)
                ],
            ))
        if status_filter in ('all', 'ineligible') and (ineligible_students or status_filter == 'ineligible'):
            sheets.append((
                'ไม่ได้รับเกียรติบัตร',
                ['ลำดับ', 'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'โรงเรียน', 'ขาดอบรม (ภาคเวลา)', 'การเข้าสอบ', 'สาเหตุ'],
                [
                    [
                        idx,
                        s['user'].camp1_code,
                        s['user'].get_full_name(),
                        s['user'].school_name,
                        f"{s['eligibility']['absences']}/{s['eligibility']['max_absences']}",
                        f"{s['eligibility']['attended_exam_slots']}/{s['eligibility']['total_exam_slots']}",
                        s['eligibility']['reason'],
                    ]
                    for idx, s in enumerate(ineligible_students, 1)
                ],
            ))
        if status_filter in ('all', 'rights') and (rights_students or status_filter == 'rights'):
            sheets.append((
                'นักเรียนใช้สิทธิ์',
                ['ลำดับ', 'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'โรงเรียน', 'หมายเหตุ'],
                [
                    [
                        idx,
                        s['user'].camp1_code,
                        s['user'].get_full_name(),
                        s['user'].school_name,
                        s['eligibility']['reason'],
                    ]
                    for idx, s in enumerate(rights_students, 1)
                ],
            ))
        if not sheets:
            sheets.append((
                f'เกียรติบัตร_{year}',
                ['ลำดับ', 'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'โรงเรียน', 'ขาดอบรม (ภาคเวลา)', 'การเข้าสอบ'],
                [],
            ))
        wb = build_workbook(sheets)
        return workbook_response(wb, f'รายชื่อผู้รับเกียรติบัตร_{year}.xlsx')

    # ข้อมูลแท็บ 2 & 3: วิทยากร / คณะกรรมการ
    all_lecturers_qs = CertificateRecipient.objects.filter(academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER)
    all_committees_qs = CertificateRecipient.objects.filter(academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE)
    lecturer_count = all_lecturers_qs.count()
    committee_count = all_committees_qs.count()

    if q:
        q_filter = Q(full_name__icontains=q) | Q(school_province__icontains=q) | Q(description__icontains=q)
        lecturers = list(all_lecturers_qs.filter(q_filter))
        committees = list(all_committees_qs.filter(q_filter))
    else:
        lecturers = list(all_lecturers_qs)
        committees = list(all_committees_qs)

    for rec in lecturers:
        rec.display_description = cert_setting.render_text(rec.resolved_description(cert_setting), year)
    for rec in committees:
        rec.display_description = cert_setting.render_text(rec.resolved_description(cert_setting), year)

    existing_records = CertificateRecord.objects.filter(academic_year=year)
    stu_ref_map = {cr.user_id: cr.ref_code for cr in existing_records if cr.user_id}
    rec_ref_map = {cr.recipient_id: cr.ref_code for cr in existing_records if cr.recipient_id}
    for s in all_students_data:
        s['ref_code'] = stu_ref_map.get(s['user'].pk, '')
    for rec in lecturers:
        rec.ref_code = rec_ref_map.get(rec.pk, '')
    for rec in committees:
        rec.ref_code = rec_ref_map.get(rec.pk, '')

    # ตัวอย่างสำหรับพรีวิวแท็บตั้งค่า
    sample_student = next((s['user'] for s in all_students_data if s['is_eligible']), None) or (
        all_students_data[0]['user'] if all_students_data else None
    )
    sample_student_name = sample_student.get_full_name() if sample_student else 'เด็กชายนิธิดล ธรรมภูษิต'
    sample_student_base_school = (
        sample_student.school_name.strip()
        if sample_student and (sample_student.school_name or '').strip()
        else 'โรงเรียนสวนกุหลาบวิทยาลัย'
    )
    sample_student_school = (
        cert_setting.format_student_school(sample_student.school_name)
        if sample_student and sample_student.school_name
        else 'โรงเรียนสวนกุหลาบวิทยาลัย กรุงเทพมหานคร'
    )
    sample_lecturer = lecturers[0] if lecturers else None
    sample_committee = committees[0] if committees else None
    active_recipients = committees if tab == 'committee' else lecturers

    by_year = {}
    for row in (
        CertificateRecipient.objects.exclude(academic_year=year)
        .values('academic_year', 'recipient_type')
        .annotate(cnt=Count('id'))
    ):
        y_key = row['academic_year']
        entry = by_year.setdefault(y_key, {'year': y_key, 'lecturer_count': 0, 'committee_count': 0, 'total_count': 0})
        if row['recipient_type'] == CertificateRecipient.TYPE_LECTURER:
            entry['lecturer_count'] = row['cnt']
        elif row['recipient_type'] == CertificateRecipient.TYPE_COMMITTEE:
            entry['committee_count'] = row['cnt']
        entry['total_count'] += row['cnt']
    transfer_source_years = sorted(by_year.values(), key=lambda x: (0 if x['year'] < year else 1, -x['year']))

    cert_year_status = CertificateYearStatus.objects.filter(academic_year=year).first()
    return render(request, 'food/admin_certificates.html', {
        'students_data': filtered_data,
        'eligible_students': eligible_students,
        'ineligible_students': ineligible_students,
        'rights_students': rights_students,
        'year': year,
        'tab': tab,
        'q': q,
        'status_filter': status_filter,
        'total_current': total_current,
        'eligible_count': eligible_count,
        'ineligible_count': ineligible_count,
        'rights_count': rights_count,
        'all_students_count': all_students_count,
        'setting': setting,
        'cert_setting': cert_setting,
        'cert_year_status': cert_year_status,
        'cert_year_published': bool(cert_year_status and cert_year_status.is_published),
        'lecturers': lecturers,
        'committees': committees,
        'active_recipients': active_recipients,
        'lecturer_count': lecturer_count,
        'committee_count': committee_count,
        'transfer_source_years': transfer_source_years,
        'rendered_date_text': cert_setting.render_text(cert_setting.line4_text, year),
        'rendered_student_line3': cert_setting.render_text(cert_setting.student_line3_text, year),
        'rendered_lecturer_line3': cert_setting.render_text(cert_setting.lecturer_line3_text, year),
        'rendered_committee_line3': cert_setting.render_text(cert_setting.committee_line3_text, year),
        'sample_student_name': sample_student_name,
        'sample_student_base_school': sample_student_base_school,
        'sample_student_school': sample_student_school,
        'sample_lecturer_name': sample_lecturer.full_name if sample_lecturer else 'นายณัฏฐ์กวี ลืนภูเขียว',
        'sample_lecturer_school': sample_lecturer.school_province if sample_lecturer else 'โรงเรียนไตรมิตรวิทยาลัย กรุงเทพมหานคร',
        'sample_lecturer_desc': (
            cert_setting.render_text(sample_lecturer.resolved_description(cert_setting), year)
            if sample_lecturer else cert_setting.render_text(cert_setting.lecturer_line3_text, year)
        ),
        'sample_committee_name': sample_committee.full_name if sample_committee else 'นายอรรถธรรม ตั้งปริมณฑล',
        'sample_committee_school': sample_committee.school_province if sample_committee else 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
        'sample_committee_desc': (
            cert_setting.render_text(sample_committee.resolved_description(cert_setting), year)
            if sample_committee else cert_setting.render_text(cert_setting.committee_line3_text, year)
        ),
    })


@admin_required
def certificate_settings_redirect(request):
    """ทางลัดจากเมนู 5. ตั้งค่าระบบ ไปยังแท็บตั้งค่าเกียรติบัตร"""
    from django.urls import reverse
    year = request.GET.get('year') or current_year()
    return redirect(reverse('food:admin_certificates') + f'?year={year}&tab=settings')


@xframe_options_sameorigin
@staff_perm_required('staff_can_scores')
def certificate_preview_pdf(request):
    """สร้างตัวอย่าง PDF จริง 1 หน้า ตามค่าที่ตั้งใน CertificateSetting (รองรับ query overrides จากสไลเดอร์)"""
    import copy
    from .certificate_pdf import build_entries_certificate_pdf, get_certificate_data_for_students

    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    group = (request.GET.get('group') or 'student').strip()
    if group not in ('student', 'lecturer', 'committee'):
        group = 'student'

    setting = SystemSetting.get()
    cert_setting = copy.copy(CertificateSetting.get())
    if request.GET.get('clear_background_image') in ('1', 'true', 'True'):
        cert_setting.background_image = None

    # อนุญาต override พิกัด/ขนาด/ข้อความผ่าน query string ชั่วคราวสำหรับการกดดูตัวอย่างก่อนกดบันทึก
    for fld, lo, hi in (
        ('header1_y', 0, 210), ('header1_font_size', 8, 96), ('header1_line_spacing', 4, 30),
        ('header2_y', 0, 210), ('header2_font_size', 8, 96),
        ('line1_y', 0, 210), ('line1_font_size', 8, 96),
        ('line2_y', 0, 210), ('line2_font_size', 8, 96),
        ('line3_y', 0, 210), ('line3_font_size', 8, 96), ('line3_line_spacing', 4, 30),
        ('line4_y', 0, 210), ('line4_font_size', 8, 96),
        ('qr_x', 5, 290), ('qr_y', 5, 205), ('qr_size', 10, 60), ('ref_code_font_size', 6, 24),
    ):
        if fld in request.GET:
            setattr(cert_setting, fld, _parse_float(request.GET.get(fld), getattr(cert_setting, fld), lo, hi))
    for b_fld in ('header1_bold', 'header2_bold', 'line1_bold', 'line2_bold', 'line3_bold', 'line4_bold', 'show_qr_verify'):
        if b_fld in request.GET:
            setattr(cert_setting, b_fld, request.GET.get(b_fld) in ('1', 'true', 'True'))
    for c_fld in ('header1_color', 'header2_color', 'line1_color', 'line2_color', 'line3_color', 'line4_color', 'ref_code_color'):
        if c_fld in request.GET:
            setattr(cert_setting, c_fld, _parse_hex_color(request.GET.get(c_fld), getattr(cert_setting, c_fld)))
    for t_fld in (
        'header1_text', 'header2_text',
        'student_default_province', 'student_line3_text', 'lecturer_line3_text', 'committee_line3_text', 'line4_text',
    ):
        if t_fld in request.GET:
            setattr(cert_setting, t_fld, request.GET.get(t_fld, ''))

    preview_ref_code = f'DSPOSN-{year}-PREVIEW-000000'
    if group == 'lecturer':
        rec = CertificateRecipient.objects.filter(academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER).first()
        lec_desc = (
            cert_setting.lecturer_line3_text
            if 'lecturer_line3_text' in request.GET
            else (rec.resolved_description(cert_setting) if rec else cert_setting.lecturer_line3_text)
        )
        entry = {
            'full_name': rec.full_name if rec else 'นายณัฏฐ์กวี ลืนภูเขียว',
            'school_province': rec.school_province if rec else 'โรงเรียนไตรมิตรวิทยาลัย กรุงเทพมหานคร',
            'description': lec_desc,
            'code': '',
            'ref_code': preview_ref_code,
        }
    elif group == 'committee':
        rec = CertificateRecipient.objects.filter(academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE).first()
        com_desc = (
            cert_setting.committee_line3_text
            if 'committee_line3_text' in request.GET
            else (rec.resolved_description(cert_setting) if rec else cert_setting.committee_line3_text)
        )
        entry = {
            'full_name': rec.full_name if rec else 'นายอรรถธรรม ตั้งปริมณฑล',
            'school_province': rec.school_province if rec else 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
            'description': com_desc,
            'code': '',
            'ref_code': preview_ref_code,
        }
    else:
        students_data = get_certificate_data_for_students(year)
        sample_u = next((s['user'] for s in students_data if s['is_eligible']), None) or (
            students_data[0]['user'] if students_data else None
        )
        entry = {
            'full_name': sample_u.get_full_name() if sample_u else 'เด็กชายนิธิดล ธรรมภูษิต',
            'school_province': (
                cert_setting.format_student_school(sample_u.school_name)
                if sample_u and sample_u.school_name
                else 'โรงเรียนสวนกุหลาบวิทยาลัย กรุงเทพมหานคร'
            ),
            'description': cert_setting.student_line3_text,
            'code': getattr(sample_u, 'camp1_code', ''),
            'ref_code': preview_ref_code,
        }

    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    pdf_bytes = build_entries_certificate_pdf(
        [entry], year, setting=setting, cert_setting=cert_setting, verify_base_url=verify_base_url,
    )
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificate_preview_{year}_{group}.pdf"'
    return resp


@staff_perm_required('staff_can_scores')
def certificate_recipient_download(request, pk):
    """ดาวน์โหลดเกียรติบัตร PDF รายบุคคลของวิทยากรหรือคณะกรรมการ"""
    from .certificate_pdf import build_recipients_certificate_pdf
    rec = get_object_or_404(CertificateRecipient, pk=pk)
    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    pdf_bytes = build_recipients_certificate_pdf(
        [rec], rec.academic_year, SystemSetting.get(), CertificateSetting.get(), verify_base_url=verify_base_url,
    )
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificate_{rec.academic_year}_{rec.recipient_type}_{rec.pk}.pdf"'
    return resp


def _cert_export_back_url(return_to, year, tab='students'):
    """คืน URL ปลายทางเมื่อ export เกียรติบัตรไม่พบรายชื่อ (กลับหน้าประกาศผลหรือหน้าเกียรติบัตร)"""
    if (return_to or '').strip() == 'announcement':
        return reverse('food:score_announcement') + f'?year={year}'
    return reverse('food:admin_certificates') + f'?year={year}&tab={tab}'


@staff_perm_required('staff_can_scores')
def certificate_recipients_export_pdf(request):
    """ดาวน์โหลดเกียรติบัตร PDF รวมเล่มสำหรับกลุ่มวิทยากรหรือคณะกรรมการ"""
    from .certificate_pdf import build_recipients_certificate_pdf
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    rtype = (request.GET.get('type') or CertificateRecipient.TYPE_LECTURER).strip()
    return_to = request.GET.get('return_to')
    qs = CertificateRecipient.objects.filter(academic_year=year)
    if rtype in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
        qs = qs.filter(recipient_type=rtype)
    recipients = list(qs)
    if not recipients:
        messages.warning(request, 'ยังไม่มีรายชื่อสำหรับออกเกียรติบัตรในกลุ่มนี้')
        back_tab = rtype if rtype in ('lecturer', 'committee') else 'lecturer'
        return redirect(_cert_export_back_url(return_to, year, back_tab))
    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    pdf_bytes = build_recipients_certificate_pdf(
        recipients, year, SystemSetting.get(), CertificateSetting.get(), verify_base_url=verify_base_url,
    )
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificates_{rtype}_{year}.pdf"'
    return resp


@staff_perm_required('staff_can_scores')
def certificate_recipients_export_xlsx(request):
    """Export รายชื่อวิทยากร / คณะกรรมการ เป็น Excel"""
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    rtype = (request.GET.get('type') or CertificateRecipient.TYPE_LECTURER).strip()
    qs = CertificateRecipient.objects.filter(academic_year=year)
    if rtype in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
        qs = qs.filter(recipient_type=rtype)
    cert_setting = CertificateSetting.get()
    date_text = cert_setting.render_text(cert_setting.line4_text, year)
    headers = ['ลำดับ', 'ประเภท', 'ชื่อ-นามสกุล', 'โรงเรียน จังหวัด', 'รายละเอียด (ข้อ 3)', 'วันที่ออกเกียรติบัตร (ข้อ 4 ร่วมกัน)']
    data = [
        [
            idx,
            r.get_recipient_type_display(),
            r.full_name,
            r.school_province,
            cert_setting.render_text(r.resolved_description(cert_setting), year),
            date_text,
        ]
        for idx, r in enumerate(qs, 1)
    ]
    wb = build_workbook([(f'รายชื่อ_{rtype}_{year}', headers, data)])
    return workbook_response(wb, f'รายชื่อเกียรติบัตร_{rtype}_{year}.xlsx')


@staff_perm_required('staff_can_scores')
def certificate_recipients_template(request):
    """ดาวน์โหลดไฟล์ Excel ตัวอย่างสำหรับนำเข้ารายชื่อวิทยากร / คณะกรรมการ"""
    headers = ['ประเภท', 'ชื่อ-นามสกุล', 'โรงเรียน จังหวัด', 'รายละเอียด (ข้อ 3 - เว้นว่างใช้ค่าเริ่มต้น)']
    rows = [
        ['วิทยากร', 'นายณัฏฐ์กวี ลืนภูเขียว', 'โรงเรียนไตรมิตรวิทยาลัย กรุงเทพมหานคร', ''],
        ['คณะกรรมการ', 'นายอรรถธรรม ตั้งปริมณฑล', 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร', ''],
    ]
    wb = build_workbook([('ตัวอย่างนำเข้า', headers, rows)])
    return workbook_response(wb, 'ตัวอย่างนำเข้ารายชื่อ_วิทยากร_คณะกรรมการ.xlsx')


@staff_perm_required('staff_can_scores')
def certificate_download(request, user_id):
    """ดาวน์โหลดเกียรติบัตรของนักเรียนรายบุคคล (PDF)"""
    from django.urls import reverse
    from .certificate_pdf import check_certificate_eligibility, build_certificate_pdf
    user = get_object_or_404(User, pk=user_id)
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    setting = SystemSetting.get()
    elig = check_certificate_eligibility(user, year)
    if not elig['is_eligible']:
        messages.error(request, f'นักเรียน {user.get_full_name()} ({user.camp1_code}) ไม่ผ่านเกณฑ์ได้รับเกียรติบัตร: {elig["reason"]}')
        return redirect(reverse('food:admin_certificates') + f'?year={year}')
    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    pdf_bytes = build_certificate_pdf([user], year, setting, verify_base_url=verify_base_url)
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificate_{year}_{user.camp1_code}.pdf"'
    return resp


@staff_perm_required('staff_can_scores')
def certificates_export_all(request):
    """ดาวน์โหลดเกียรติบัตรของผู้ที่ผ่านเกณฑ์ทุกคนในปีการศึกษานั้น รวมเป็นไฟล์ PDF เดียว (Batch Export)"""
    from .certificate_pdf import get_certificate_data_for_students, build_certificate_pdf
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    return_to = request.GET.get('return_to')
    setting = SystemSetting.get()
    students_data = get_certificate_data_for_students(year)
    eligible_users = [s['user'] for s in students_data if s['is_eligible']]
    if not eligible_users:
        messages.warning(request, f'ไม่มีนักเรียนที่ผ่านเกณฑ์ได้รับเกียรติบัตรในปีการศึกษา {year}')
        return redirect(_cert_export_back_url(return_to, year, 'students'))
    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    pdf_bytes = build_certificate_pdf(eligible_users, year, setting, verify_base_url=verify_base_url)
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificates_all_{year}.pdf"'
    return resp


def _parse_int_set(raw_list):
    out = set()
    for val in raw_list:
        try:
            out.add(int(val))
        except (TypeError, ValueError):
            pass
    return out


@staff_perm_required('staff_can_scores')
def certificates_export_zip(request):
    """ดาวน์โหลดเกียรติบัตร PDF แยกเป็นไฟล์รายบุคคลบีบอัดเป็นไฟล์ .zip (เลือกเฉพาะคนหรือเลือกกลุ่มได้)"""
    from django.urls import reverse
    from .certificate_pdf import (
        build_certificates_zip, build_recipient_cert_entry, build_student_cert_entry,
        get_certificate_data_for_students,
    )

    data = request.POST if request.method == 'POST' else request.GET
    year_req = data.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()

    scope = (data.get('scope') or 'students').strip()
    if scope not in ('students', 'lecturer', 'committee', 'all'):
        scope = 'students'
    select_mode = (data.get('select_mode') or 'all').strip()
    if select_mode not in ('all', 'selected'):
        select_mode = 'all'

    back_tab = scope if scope in ('students', 'lecturer', 'committee') else 'students'
    return_to = data.get('return_to')
    setting = SystemSetting.get()
    cert_setting = CertificateSetting.get()
    items = []

    if scope == 'students':
        students_data = get_certificate_data_for_students(year)
        eligible_users = [s['user'] for s in students_data if s['is_eligible']]
        if select_mode == 'selected':
            sel_ids = _parse_int_set(data.getlist('student_ids'))
            eligible_users = [u for u in eligible_users if u.pk in sel_ids]
            if not eligible_users:
                messages.warning(request, 'กรุณาเลือกรายชื่อนักเรียนที่มีสิทธิ์รับเกียรติบัตรอย่างน้อย 1 คน')
                return redirect(_cert_export_back_url(return_to, year, 'students'))
        for u in eligible_users:
            fname = f'{u.camp1_code}_{u.get_full_name()}'.strip('_')
            items.append({
                'entry': build_student_cert_entry(u, year, cert_setting),
                'folder': '',
                'filename': fname,
            })
        zip_name = f'certificates_students_{year}.zip'

    elif scope in (CertificateRecipient.TYPE_LECTURER, CertificateRecipient.TYPE_COMMITTEE):
        qs = CertificateRecipient.objects.filter(academic_year=year, recipient_type=scope)
        if select_mode == 'selected':
            sel_ids = _parse_int_set(data.getlist('recipient_ids'))
            if not sel_ids:
                messages.warning(request, 'กรุณาเลือกรายชื่อที่ต้องการดาวน์โหลดอย่างน้อย 1 คน')
                return redirect(_cert_export_back_url(return_to, year, scope))
            qs = qs.filter(pk__in=sel_ids)
        for idx, rec in enumerate(qs, 1):
            items.append({
                'entry': build_recipient_cert_entry(rec, cert_setting),
                'folder': '',
                'filename': f'{idx:02d}_{rec.full_name}',
            })
        zip_name = f'certificates_{scope}_{year}.zip'

    else:
        has_toggles = (request.method == 'POST' and data.get('has_group_toggles') == '1')
        inc_stu = (data.get('include_students') == '1') if has_toggles else True
        inc_lec = (data.get('include_lecturer') == '1') if has_toggles else True
        inc_com = (data.get('include_committee') == '1') if has_toggles else True
        sel_stu_ids = _parse_int_set(data.getlist('student_ids')) if select_mode == 'selected' else None
        sel_rec_ids = _parse_int_set(data.getlist('recipient_ids')) if select_mode == 'selected' else None

        if inc_stu:
            students_data = get_certificate_data_for_students(year)
            eligible_users = [s['user'] for s in students_data if s['is_eligible']]
            if sel_stu_ids is not None:
                eligible_users = [u for u in eligible_users if u.pk in sel_stu_ids]
            for u in eligible_users:
                fname = f'{u.camp1_code}_{u.get_full_name()}'.strip('_')
                items.append({
                    'entry': build_student_cert_entry(u, year, cert_setting),
                    'folder': '01_นักเรียน',
                    'filename': fname,
                })
        if inc_lec:
            lec_qs = CertificateRecipient.objects.filter(
                academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER
            )
            if sel_rec_ids is not None:
                lec_qs = lec_qs.filter(pk__in=sel_rec_ids)
            for idx, rec in enumerate(lec_qs, 1):
                items.append({
                    'entry': build_recipient_cert_entry(rec, cert_setting),
                    'folder': '02_วิทยากร',
                    'filename': f'{idx:02d}_{rec.full_name}',
                })
        if inc_com:
            com_qs = CertificateRecipient.objects.filter(
                academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE
            )
            if sel_rec_ids is not None:
                com_qs = com_qs.filter(pk__in=sel_rec_ids)
            for idx, rec in enumerate(com_qs, 1):
                items.append({
                    'entry': build_recipient_cert_entry(rec, cert_setting),
                    'folder': '03_คณะกรรมการ',
                    'filename': f'{idx:02d}_{rec.full_name}',
                })
        zip_name = f'certificates_all_{year}.zip'

    if not items:
        messages.warning(request, f'ไม่พบรายชื่อที่มีสิทธิ์ออกเกียรติบัตรสำหรับดาวน์โหลดในปีการศึกษา {year}')
        return redirect(_cert_export_back_url(return_to, year, back_tab))

    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    zip_bytes = build_certificates_zip(
        items, year, setting=setting, cert_setting=cert_setting, verify_base_url=verify_base_url,
    )
    resp = HttpResponse(zip_bytes, content_type='application/zip')
    resp['Content-Disposition'] = f'attachment; filename="{zip_name}"'
    return resp


def my_certificate(request):
    """นักเรียนตรวจสอบและดาวน์โหลดเกียรติบัตรของตนเอง"""
    if not request.user.is_authenticated or request.user.role != User.ROLE_MEMBER:
        messages.error(request, 'เฉพาะนักเรียนเท่านั้นที่สามารถดาวน์โหลดเกียรติบัตรตนเองได้')
        return redirect('composn:login')
    from .certificate_pdf import check_certificate_eligibility, build_certificate_pdf
    default_year = request.user.base_academic_year or current_year()
    year_req = request.GET.get('year')
    try:
        year = int(year_req) if year_req else int(default_year)
    except (ValueError, TypeError):
        year = int(default_year)
    if not CertificateYearStatus.is_year_published(year):
        messages.warning(request, f'ผู้ดูแลระบบยังไม่เปิดให้แสดงและดาวน์โหลดเกียรติบัตรประจำปีการศึกษา {year}')
        return redirect('composn:profile')
    setting = SystemSetting.get()
    elig = check_certificate_eligibility(request.user, year)
    if not elig['is_eligible']:
        messages.warning(request, f'ท่านยังไม่ได้รับเกียรติบัตรในปีการศึกษา {year}: {elig["reason"]}')
        return redirect('composn:profile')
    from django.core.cache import cache
    cert_setting = CertificateSetting.get()
    verify_base_url = request.build_absolute_uri('/').rstrip('/')
    cache_key = (
        f'dscomposn:cert_pdf:{request.user.pk}:{year}:'
        f'{int(cert_setting.updated_at.timestamp())}:'
        f'{int(setting.updated_at.timestamp())}:'
        f'{int(request.user.updated_at.timestamp())}:'
        f'{request.get_host()}'
    )
    pdf_bytes = cache.get(cache_key)
    if pdf_bytes is None:
        pdf_bytes = build_certificate_pdf([request.user], year, setting, verify_base_url=verify_base_url)
        cache.set(cache_key, pdf_bytes, 600)
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="certificate_{year}_{request.user.camp1_code}.pdf"'
    return resp


def certificate_verify(request, ref_code=None):
    """หน้าสาธารณะตรวจสอบความถูกต้องของเกียรติบัตรด้วยเลขที่อ้างอิง / QR Code (ไม่ต้อง Login)"""
    from .certificate_pdf import check_certificate_eligibility
    code = (ref_code or request.GET.get('code') or '').strip().upper()
    setting = SystemSetting.get()
    record = None
    status = 'empty'
    status_detail = ''
    if code:
        record = CertificateRecord.objects.select_related('user', 'recipient').filter(ref_code__iexact=code).first()
        if not record:
            status = 'not_found'
        elif record.is_revoked:
            status = 'revoked'
        elif record.recipient_type == CertificateRecord.TYPE_STUDENT and record.user:
            elig = check_certificate_eligibility(record.user, record.academic_year)
            if not elig['is_eligible']:
                status = 'invalid'
                status_detail = elig.get('reason', '')
            else:
                status = 'valid'
        else:
            status = 'valid'
    return render(request, 'food/certificate_verify.html', {
        'setting': setting,
        'code': code,
        'record': record,
        'status': status,
        'status_detail': status_detail,
    })


# =============================================================================
# Admin: ตั้งค่ากำหนดคะแนน สัดส่วน / ผู้ผ่านการคัดเลือก / วันอบรม-วันสอบ / เหตุผล / สี
# =============================================================================
def _parse_bool_flag(val):
    return str(val or '').strip().lower() in ('1', 'true', 'on', 'yes')


def _valid_hex_or_empty(raw):
    val = str(raw or '').strip()
    return val if re.match(r'^#[0-9a-fA-F]{6}$', val) else ''


@admin_required
def score_categories(request):
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else current_year()
    except (ValueError, TypeError):
        year = current_year()
    tab = (request.GET.get('tab') or request.POST.get('tab') or 'scores').strip()
    if request.method == 'GET' and tab == 'qualification':
        return redirect(f"{reverse('food:score_announcement')}?year={year}&tab=qualification")
    if tab not in ('scores', 'dates'):
        tab = 'scores'
    setting = SystemSetting.get()
    approval = ScoreApproval.objects.filter(academic_year=year).select_related('approved_by').first()
    is_approved = bool(approval and approval.is_approved)
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'add_academic_year':
            from composn.views_admin import handle_add_academic_year_post
            handle_add_academic_year_post(request, setting)
            return redirect(f"{reverse('food:score_categories')}?year={setting.current_academic_year}&tab={tab}")
        if action == 'save_academic_year_settings':
            old_year = setting.current_academic_year
            old_max_abs = setting.max_absences
            cur_year_approved = ScoreApproval.is_year_approved(old_year)
            if request.POST.get('current_academic_year'):
                try:
                    req_year = int(request.POST.get('current_academic_year'))
                    allowed_years = set(get_academic_year_choices(old_year, setting=setting))
                    if req_year in allowed_years:
                        setting.current_academic_year = req_year
                    else:
                        messages.warning(
                            request,
                            f'ปีการศึกษา {req_year} ยังไม่มีในระบบ กรุณากดปุ่ม "เพิ่มปีการศึกษา" ก่อนเลือกใช้งาน',
                        )
                except (TypeError, ValueError):
                    pass
            if 'max_absences' in request.POST:
                try:
                    setting.max_absences = max(1, int(request.POST.get('max_absences') or old_max_abs))
                except (TypeError, ValueError):
                    pass
            if (
                cur_year_approved
                and int(setting.current_academic_year) == int(old_year)
                and int(setting.max_absences) != int(old_max_abs)
            ):
                setting.max_absences = old_max_abs
                messages.warning(
                    request,
                    f'ปีการศึกษา {old_year} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนจำนวนครั้งขาดอบรมสูงสุดได้ (คงค่าเดิมไว้)',
                )
            setting.save()
            log_action(
                request.user, 'update_settings', target=setting,
                old_value={'current_academic_year': old_year, 'max_absences': old_max_abs},
                new_value={'current_academic_year': setting.current_academic_year, 'max_absences': setting.max_absences},
                request=request,
            )
            messages.success(
                request,
                f'บันทึกการตั้งค่าปีการศึกษาเรียบร้อยแล้ว (ปีการศึกษาปัจจุบัน: {setting.current_academic_year})',
            )
            return redirect(f"{reverse('food:score_categories')}?year={setting.current_academic_year}&tab={tab}")

        if action in ('save_qualification', 'approve_scores', 'unapprove_scores'):
            tab = 'qualification'
        elif action in ('add_training_date', 'update_training_date', 'save_all_training_dates', 'delete_training_date'):
            tab = 'dates'
        else:
            tab = 'scores'
        try:
            if is_approved and action in (
                'create', 'update', 'create_sub', 'save_all_categories', 'delete',
                'add_training_date', 'update_training_date', 'save_all_training_dates', 'delete_training_date',
            ):
                if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    return JsonResponse({'ok': False, 'error': year_locked_message(year)}, status=403)
                raise FoodError(year_locked_message(year))
            if action in ('create', 'update', 'create_sub'):
                parent_id = request.POST.get('parent_id')
                parent_cat = ScoreCategory.objects.filter(pk=parent_id, academic_year=year, parent__isnull=True).first() if parent_id else None
                group_type = (request.POST.get('group_type') or ScoreCategory.GROUP_COLLECT).strip()
                raw_name = request.POST.get('name', '').strip()
                if action == 'create' and not parent_cat and not raw_name:
                    raw_name = group_type
                if not raw_name:
                    raise FoodError('กรุณาระบุชื่อหัวข้อคะแนน')
                weight = to_decimal(request.POST.get('weight'))
                if weight is None:
                    weight = Decimal('0')
                max_score = to_decimal(request.POST.get('default_max_score')) or Decimal('100')
                order = int(request.POST.get('order') or 0)
                is_visible = _parse_bool_flag(request.POST.get('is_visible')) if action == 'update' else _parse_bool_flag(request.POST.get('is_visible', '1'))
                is_open = _parse_bool_flag(request.POST.get('is_open')) if action == 'update' else _parse_bool_flag(request.POST.get('is_open', '1'))
                open_at = _parse_dt_local(request.POST.get('open_at'))
                close_at = _parse_dt_local(request.POST.get('close_at'))

                if action == 'create_sub':
                    if not parent_cat:
                        raise FoodError('ไม่พบหัวข้อใหญ่ที่ต้องการเพิ่มหัวข้อย่อย')
                    # รองรับการเพิ่มหลายหัวข้อย่อยพร้อมกัน (คั่นด้วยคอมม่า เช่น ข้อ1, ข้อ2, ข้อ3)
                    sub_names = [s.strip() for s in re.split(r'[,;\n]+', raw_name) if s.strip()]
                    base_order = order or (parent_cat.subcategories.count() + 1)
                    sub_score_val = to_decimal(request.POST.get('default_max_score') or request.POST.get('weight')) or Decimal('10')
                    last_cat = None
                    for idx, sname in enumerate(sub_names):
                        last_cat = ScoreCategory.objects.create(
                            parent=parent_cat, group_type=parent_cat.group_type,
                            name=sname, weight=sub_score_val, academic_year=year, order=base_order + idx,
                            default_max_score=sub_score_val, is_visible=is_visible,
                            is_open=is_open, open_at=None, close_at=None,
                        )
                    # อัปเดตคะแนนเต็มของหัวข้อใหญ่ให้เท่ากับผลรวมคะแนนของหัวข้อย่อยทั้งหมด
                    sub_total = sum((sc.default_max_score for sc in parent_cat.subcategories.all()), Decimal('0'))
                    parent_cat.default_max_score = sub_total
                    parent_cat.save(update_fields=['default_max_score'])
                    if last_cat:
                        log_action(request.user, 'score_category', target=last_cat, new_value={
                            'action': 'create_sub', 'parent': parent_cat.name, 'subs': sub_names}, request=request)
                    messages.success(request, f'เพิ่มหัวข้อย่อยใน "{parent_cat.name}" เรียบร้อย ({len(sub_names)} รายการ)')
                elif action == 'create':
                    box_color = _valid_hex_or_empty(request.POST.get('box_color'))
                    if not parent_cat and not box_color:
                        box_color = ScoreCategory.pick_random_unused_color(year)
                    cat = ScoreCategory.objects.create(
                        parent=parent_cat, group_type=group_type,
                        name=raw_name, weight=weight, academic_year=year, order=order,
                        default_max_score=max_score, box_color=box_color,
                        is_visible=is_visible, is_open=is_open, open_at=open_at, close_at=close_at,
                    )
                    log_action(request.user, 'score_category', target=cat, new_value={
                        'action': action, 'group_type': group_type, 'name': raw_name, 'weight': str(weight),
                        'is_visible': is_visible, 'box_color': cat.box_color}, request=request)
                    messages.success(request, 'เพิ่มหัวข้อคะแนนเรียบร้อย')
                else:
                    cat = get_object_or_404(ScoreCategory, pk=request.POST.get('id'))
                    if not cat.parent_id and request.POST.get('group_type'):
                        cat.group_type = group_type
                    if cat.parent_id:
                        sub_score_val = to_decimal(request.POST.get('default_max_score') or request.POST.get('weight')) or Decimal('10')
                        weight = sub_score_val
                        max_score = sub_score_val
                        cat.name, cat.weight, cat.order, cat.default_max_score = raw_name, weight, order, max_score
                        cat.is_visible = is_visible
                        cat.is_open = is_open
                        cat.open_at = None
                        cat.close_at = None
                        cat.save()
                        if cat.parent:
                            sub_total = sum((sc.default_max_score for sc in cat.parent.subcategories.all()), Decimal('0'))
                            cat.parent.default_max_score = sub_total
                            cat.parent.save(update_fields=['default_max_score'])
                    else:
                        # บันทึกหัวข้อใหญ่ พร้อมอัปเดตหัวข้อย่อยทั้งหมดภายใต้หัวข้อใหญ่นี้ในครั้งเดียว
                        subs = list(cat.subcategories.all())
                        sub_total = Decimal('0')
                        for sc in subs:
                            sid = str(sc.pk)
                            if f'sub_name_{sid}' in request.POST:
                                s_name = (request.POST.get(f'sub_name_{sid}') or sc.name).strip()
                                if s_name:
                                    sc.name = s_name
                                s_score = to_decimal(request.POST.get(f'sub_score_{sid}')) or Decimal('10')
                                sc.weight = s_score
                                sc.default_max_score = s_score
                                sc.order = int(request.POST.get(f'sub_order_{sid}') or sc.order)
                                sc.is_visible = _parse_bool_flag(request.POST.get(f'sub_visible_{sid}'))
                                sc.is_open = _parse_bool_flag(request.POST.get(f'sub_open_{sid}'))
                                sc.open_at = None
                                sc.close_at = None
                                sc.save()
                            sub_total += sc.default_max_score
                        if subs:
                            max_score = sub_total
                        box_color = _valid_hex_or_empty(request.POST.get('box_color'))
                        if box_color:
                            cat.box_color = box_color
                        cat.name, cat.weight, cat.order, cat.default_max_score = raw_name, weight, order, max_score
                        cat.is_visible = is_visible
                        cat.is_open = is_open
                        cat.open_at = open_at
                        cat.close_at = close_at
                        cat.save()
                    log_action(request.user, 'score_category', target=cat, new_value={
                        'action': action, 'name': raw_name, 'weight': str(weight), 'is_visible': is_visible},
                        request=request)
                    messages.success(request, f'บันทึกหัวข้อ "{cat.name}" และหัวข้อย่อยเรียบร้อย')
            elif action == 'save_all_categories':
                all_cats = list(ScoreCategory.objects.filter(academic_year=year))
                saved_cnt = 0
                for c in all_cats:
                    cid = str(c.pk)
                    if f'name_{cid}' not in request.POST:
                        continue
                    c_name = (request.POST.get(f'name_{cid}') or c.name).strip()
                    if not c_name:
                        continue
                    c.name = c_name
                    c.order = int(request.POST.get(f'order_{cid}') or c.order)
                    c.is_visible = _parse_bool_flag(request.POST.get(f'is_visible_{cid}'))
                    c.is_open = _parse_bool_flag(request.POST.get(f'is_open_{cid}'))
                    if not c.parent_id:
                        g_type = (request.POST.get(f'group_type_{cid}') or c.group_type).strip()
                        c.group_type = g_type
                        box_col = _valid_hex_or_empty(request.POST.get(f'box_color_{cid}'))
                        if box_col:
                            c.box_color = box_col
                        c.weight = to_decimal(request.POST.get(f'weight_{cid}')) or Decimal('0')
                        c.default_max_score = to_decimal(request.POST.get(f'default_max_score_{cid}')) or Decimal('100')
                        c.open_at = _parse_dt_local(request.POST.get(f'open_at_{cid}'))
                        c.close_at = _parse_dt_local(request.POST.get(f'close_at_{cid}'))
                    else:
                        sub_score = to_decimal(request.POST.get(f'default_max_score_{cid}') or request.POST.get(f'weight_{cid}')) or Decimal('10')
                        c.weight = sub_score
                        c.default_max_score = sub_score
                        c.open_at = None
                        c.close_at = None
                    c.save()
                    saved_cnt += 1
                # อัปเดตคะแนนรวมของหัวข้อใหญ่ที่มีหัวข้อย่อย ให้เท่ากับผลรวมคะแนนของหัวข้อย่อยทั้งหมดเสมอ
                for mc in ScoreCategory.objects.filter(academic_year=year, parent__isnull=True):
                    if mc.subcategories.exists():
                        mc.default_max_score = sum((sc.default_max_score for sc in mc.subcategories.all()), Decimal('0'))
                        mc.save(update_fields=['default_max_score'])
                log_action(request.user, 'score_category', target_model='ScoreCategory',
                           new_value={'action': 'save_all_categories', 'year': year, 'count': saved_cnt},
                           request=request)
                messages.success(request, f'บันทึกการตั้งค่าหัวข้อคะแนนทั้งหมดเรียบร้อย ({saved_cnt} รายการ)')
            elif action == 'delete':
                cat = get_object_or_404(ScoreCategory, pk=request.POST.get('id'))
                parent_to_sync = cat.parent if cat.parent_id else None
                sub_ids = list(cat.subcategories.values_list('id', flat=True))
                target_ids = [cat.id] + sub_ids
                alive_cnt = Score.objects.filter(category_id__in=target_ids).count()
                if alive_cnt > 0:
                    confirm_password = request.POST.get('confirm_password') or ''
                    if not confirm_password:
                        raise FoodError(f'หัวข้อ "{cat.name}" มีคะแนนที่บันทึกแล้ว {alive_cnt} รายการ กรุณายืนยันการลบด้วยการพิมพ์รหัสผ่าน')
                    if not request.user.check_password(confirm_password):
                        raise FoodError('รหัสผ่านยืนยันไม่ถูกต้อง ไม่สามารถลบหัวข้อคะแนนได้')
                old_log = {'action': 'delete', 'name': cat.name}
                if alive_cnt > 0:
                    old_log.update({'deleted_scores': alive_cnt, 'password_verified': True})
                with transaction.atomic():
                    Score.all_objects.filter(category_id__in=target_ids).hard_delete()
                    log_action(request.user, 'score_category', target=cat, old_value=old_log, request=request)
                    cat.delete()
                    if parent_to_sync and parent_to_sync.subcategories.exists():
                        parent_to_sync.default_max_score = sum(
                            (sc.default_max_score for sc in parent_to_sync.subcategories.all()), Decimal('0')
                        )
                        parent_to_sync.save(update_fields=['default_max_score'])
                if alive_cnt > 0:
                    messages.success(request, f'ลบหัวข้อคะแนน "{cat.name}" และคะแนนที่บันทึกไว้ {alive_cnt} รายการเรียบร้อยแล้ว')
                else:
                    messages.success(request, 'ลบหมวดคะแนนแล้ว')
            elif action in ('save_qualification', 'approve_scores', 'unapprove_scores'):
                _process_qualification_post(request, action, year, setting, approval, is_approved)
            elif action == 'add_training_date':
                raw_start = (request.POST.get('date_start') or request.POST.get('date') or '').strip()
                if not raw_start:
                    raise FoodError('กรุณาระบุวันที่เริ่มต้น')
                d_start = _parse_date(raw_start)
                raw_end = (request.POST.get('date_end') or '').strip()
                d_end = _parse_date(raw_end) if raw_end else d_start
                if d_end < d_start:
                    d_start, d_end = d_end, d_start
                if (d_end - d_start).days > 90:
                    raise FoodError('ช่วงวันที่ต้องไม่เกิน 90 วันต่อครั้ง')
                ev_type = request.POST.get('event_type', TrainingDate.TYPE_TRAINING)
                exam_period = request.POST.get('exam_period', TrainingDate.EXAM_BOTH)
                if exam_period not in dict(TrainingDate.EXAM_PERIOD_CHOICES):
                    exam_period = TrainingDate.EXAM_BOTH
                exam_type = request.POST.get('exam_type', TrainingDate.EXAM_TYPE_AUTO)
                if exam_type not in dict(TrainingDate.EXAM_TYPE_CHOICES):
                    exam_type = TrainingDate.EXAM_TYPE_AUTO
                title = request.POST.get('title', '').strip()
                created_count = 0
                cur = d_start
                last_td = None
                while cur <= d_end:
                    last_td, _ = TrainingDate.objects.update_or_create(
                        academic_year=year, date=cur,
                        defaults={
                            'event_type': ev_type,
                            'exam_period': exam_period,
                            'exam_type': exam_type if ev_type == TrainingDate.TYPE_EXAM else TrainingDate.EXAM_TYPE_AUTO,
                            'camp_level': 1,
                            'title': title,
                        },
                    )
                    created_count += 1
                    cur += dt.timedelta(days=1)
                if last_td:
                    log_action(request.user, 'update_settings', target=last_td,
                               new_value={'date_start': str(d_start), 'date_end': str(d_end),
                                          'type': ev_type, 'exam_period': exam_period, 'exam_type': exam_type,
                                          'title': title, 'count': created_count}, request=request)
                messages.success(request, f'เพิ่มวันที่อบรม/วันสอบ/วันหยุดเรียบร้อย ({created_count} วัน)')
            elif action == 'update_training_date':
                td = get_object_or_404(TrainingDate, pk=request.POST.get('td_id'))
                ev_type = request.POST.get('event_type', td.event_type)
                if ev_type not in dict(TrainingDate.TYPE_CHOICES):
                    ev_type = TrainingDate.TYPE_TRAINING
                td.event_type = ev_type
                if 'exam_period' in request.POST:
                    ep = request.POST.get('exam_period', td.exam_period)
                    if ep in dict(TrainingDate.EXAM_PERIOD_CHOICES):
                        td.exam_period = ep
                if 'exam_type' in request.POST:
                    et = request.POST.get('exam_type', td.exam_type)
                    if et in dict(TrainingDate.EXAM_TYPE_CHOICES):
                        td.exam_type = et
                if 'title' in request.POST:
                    td.title = request.POST.get('title', '').strip()
                td.save()
                log_action(request.user, 'update_settings', target=td,
                           new_value={'date': str(td.date), 'type': td.event_type,
                                      'exam_period': td.exam_period, 'exam_type': td.exam_type, 'title': td.title},
                           request=request)
                if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                    t_qs = TrainingDate.objects.filter(academic_year=year)
                    return JsonResponse({
                        'ok': True,
                        'event_type': td.event_type,
                        'exam_period': td.exam_period,
                        'exam_type': td.exam_type,
                        'display': td.get_event_type_display(),
                        'counts': {
                            'training': t_qs.filter(event_type=TrainingDate.TYPE_TRAINING).count(),
                            'exam': t_qs.filter(event_type=TrainingDate.TYPE_EXAM).count(),
                            'holiday': t_qs.filter(event_type=TrainingDate.TYPE_HOLIDAY).count(),
                            'total': t_qs.count(),
                        },
                    })
                messages.success(request, f'บันทึกการเปลี่ยนแปลงวันที่ {td.date} เรียบร้อย')
            elif action == 'save_all_training_dates':
                t_qs = list(TrainingDate.objects.filter(academic_year=year))
                saved_td = 0
                for td in t_qs:
                    tid = str(td.pk)
                    if (f'td_type_{tid}' not in request.POST and f'td_title_{tid}' not in request.POST
                            and f'td_exam_period_{tid}' not in request.POST
                            and f'td_exam_type_{tid}' not in request.POST):
                        continue
                    ev_type = request.POST.get(f'td_type_{tid}', td.event_type)
                    if ev_type in dict(TrainingDate.TYPE_CHOICES):
                        td.event_type = ev_type
                    if f'td_exam_period_{tid}' in request.POST:
                        ep = request.POST.get(f'td_exam_period_{tid}', td.exam_period)
                        if ep in dict(TrainingDate.EXAM_PERIOD_CHOICES):
                            td.exam_period = ep
                    if f'td_exam_type_{tid}' in request.POST:
                        et = request.POST.get(f'td_exam_type_{tid}', td.exam_type)
                        if et in dict(TrainingDate.EXAM_TYPE_CHOICES):
                            td.exam_type = et
                    if f'td_title_{tid}' in request.POST:
                        td.title = request.POST.get(f'td_title_{tid}', '').strip()
                    td.save()
                    saved_td += 1
                log_action(request.user, 'update_settings', target_model='TrainingDate',
                           new_value={'action': 'save_all_training_dates', 'year': year, 'count': saved_td},
                           request=request)
                messages.success(request, f'บันทึกข้อมูลวันที่อบรม/วันสอบ/วันหยุดทั้งหมดเรียบร้อย ({saved_td} วัน)')
            elif action == 'delete_training_date':
                td = get_object_or_404(TrainingDate, pk=request.POST.get('td_id'))
                td.delete()
                messages.success(request, 'ลบรายการวันที่แล้ว')
        except (FoodError, ValueError) as e:
            messages.error(request, str(e))
        if tab == 'qualification':
            return redirect(f"{reverse('food:score_announcement')}?year={year}&tab=qualification")
        return redirect(f'{request.path}?year={year}&tab={tab}')

    main_cats = list(
        ScoreCategory.objects.filter(academic_year=year, parent__isnull=True)
        .annotate(n=Count('scores', filter=Q(scores__is_deleted=False)))
        .order_by('order', 'id')
    )
    for m_idx, mc in enumerate(main_cats, 1):
        mc.main_no = m_idx
        if not (mc.box_color or '').strip():
            mc.box_color = ScoreCategory.pick_random_unused_color(year)
            ScoreCategory.objects.filter(pk=mc.pk).update(box_color=mc.box_color)
        subs = list(
            mc.subcategories.annotate(n=Count('scores', filter=Q(scores__is_deleted=False))).order_by('order', 'id')
        )
        for s_idx, sc in enumerate(subs, 1):
            sc.sub_no = f'{m_idx}.{s_idx}'
        mc.sub_list = subs
        mc.total_scores_count = mc.n + sum(sc.n for sc in subs)
        if subs:
            sub_sum = sum((sc.default_max_score for sc in subs), Decimal('0'))
            if mc.default_max_score != sub_sum:
                mc.default_max_score = sub_sum
                ScoreCategory.objects.filter(pk=mc.pk).update(default_max_score=sub_sum)

    training_dates = list(TrainingDate.objects.filter(academic_year=year))
    date_counts = {
        'training': sum(1 for t in training_dates if t.event_type == TrainingDate.TYPE_TRAINING),
        'exam': sum(1 for t in training_dates if t.event_type == TrainingDate.TYPE_EXAM),
        'holiday': sum(1 for t in training_dates if t.event_type == TrainingDate.TYPE_HOLIDAY),
        'total': len(training_dates),
    }
    available_years_desc = get_academic_year_choices(setting.current_academic_year, setting=setting)
    available_years_asc = sorted(available_years_desc)
    max_academic_year = max(available_years_asc) if available_years_asc else MIN_ACADEMIC_YEAR

    total_weight = sum((c.weight for c in main_cats), Decimal('0')).quantize(Decimal('0.01'))
    total_max_score = sum((c.default_max_score for c in main_cats), Decimal('0')).quantize(Decimal('0.01'))
    return render(request, 'food/score_categories.html', {
        'categories': main_cats,
        'group_choices': ScoreCategory.GROUP_CHOICES,
        'next_random_color': ScoreCategory.pick_random_unused_color(year),
        'year': year,
        'tab': tab,
        'total_weight': total_weight,
        'total_max_score': total_max_score,
        'setting': setting,
        'training_dates': training_dates,
        'exam_period_choices': TrainingDate.EXAM_PERIOD_CHOICES,
        'exam_type_choices': TrainingDate.EXAM_TYPE_CHOICES,
        'date_counts': date_counts,
        'approval': approval,
        'is_approved': is_approved,
        'available_years_desc': available_years_desc,
        'available_years_asc': available_years_asc,
        'min_academic_year': MIN_ACADEMIC_YEAR,
        'max_academic_year': max_academic_year,
        'next_academic_year': min(MAX_ACADEMIC_YEAR, max_academic_year + 1),
        'current_year_approved': ScoreApproval.is_year_approved(setting.current_academic_year),
    })


@admin_required
@require_POST
def colors_manage(request):
    action = request.POST.get('action')
    try:
        color = (request.POST.get('color') or '#ffc107').strip()
        if action in ('create', 'update'):
            if not re.match(r'^#[0-9a-fA-F]{6}$', color):
                raise FoodError('รหัสสีไม่ถูกต้อง')
            name = request.POST.get('name', '').strip() or color
            min_abs = max(1, int(request.POST.get('min_absences') or 1))
            if action == 'create':
                WarningColor.objects.create(name=name, color=color, min_absences=min_abs,
                                            order=WarningColor.objects.count() + 1)
            else:
                c = get_object_or_404(WarningColor, pk=request.POST.get('id'))
                c.name, c.color, c.min_absences = name, color, min_abs
                c.save()
        elif action == 'delete':
            get_object_or_404(WarningColor, pk=request.POST.get('id')).delete()
        elif action == 'reset':
            from .apps import DEFAULT_COLORS
            WarningColor.objects.all().delete()
            for i, (n, c, m) in enumerate(DEFAULT_COLORS, start=1):
                WarningColor.objects.create(name=n, color=c, min_absences=m, order=i)
        else:
            raise FoodError('คำสั่งไม่ถูกต้อง')
        log_action(request.user, 'update_settings', target_model='WarningColor',
                   new_value={'action': action, 'color': color}, request=request)
        messages.success(request, 'บันทึกสีเตือนภัยแล้ว')
    except (FoodError, ValueError) as e:
        messages.error(request, str(e))
    return redirect('composn:admin_settings')
