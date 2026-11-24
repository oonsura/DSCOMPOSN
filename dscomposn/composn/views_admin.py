"""
Views ฝั่ง Admin: จัดการผู้ใช้ / Import / คำนวณระดับชั้น / รอบใช้สิทธิ์ / คำขอโรงเรียน /
ตั้งค่าระบบ / Backup / Export / Audit Log / ถังขยะ (Soft Delete)
"""
import datetime as dt
import re

from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, ProtectedError, Q
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_POST

from . import levels, rights
from .audit import ACTION_LABELS, diff_snapshot, log_action, model_snapshot
from .context_processors import get_academic_year_choices
from .decorators import admin_required, staff_perm_required
from .excel import (CAMP2_COLUMNS, CAMP2_COLUMNS_TH, STUDENT_COLUMNS, STUDENT_COLUMNS_TH, USE_RECORD_HEADERS,
                    backup_use_records, build_workbook, camp2_template_workbook, history_workbook, import_camp2,
                    import_students, import_students_from_preview, list_backups, preview_students, read_rights_codes,
                    rights_template_workbook, safe_backup_path, students_template_workbook, students_workbook,
                    template_workbook, use_record_rows, use_records_queryset, workbook_response)
from .face_photo import FacePhotoError, process_and_validate_profile_photo
from .forms import (STAFF_PERM_FIELDS, USER_FIELDS, AdminSetPasswordForm, AdminUserCreateForm, ExcelUploadForm,
                    OpenRoundForm, SystemSettingForm, UserForm)
from .logo_utils import convert_existing_setting_logos_to_transparent, process_logo_upload
from .models import (ArchiveRecord, AuditLog, DASHBOARD_SUMMARY_CHOICES, DASHBOARD_WIDGET_CHOICES,
                     DEFAULT_DASHBOARD_WIDGETS, MAX_ACADEMIC_YEAR, MIN_ACADEMIC_YEAR,
                     PROFILE_EDITABLE_CHOICES, RightsRound,
                     SchoolChangeRequest, SystemSetting, UseRecord, User)


def _lock_helpers():
    from food.models import (
        ScoreApproval, approved_years_set, camp2_code_year, user_approved_years, year_locked_message,
    )
    return ScoreApproval, approved_years_set, camp2_code_year, user_approved_years, year_locked_message


def _paginate(request, qs, per_page=50):
    return Paginator(qs, per_page).get_page(request.GET.get('page'))


_SAFE_NEXT_NAMES = {
    'composn:admin_rounds', 'composn:admin_users', 'composn:admin_admins',
    'composn:admin_use_records', 'composn:admin_dashboard',
}


def _safe_redirect_next(request, fallback, **kwargs):
    nxt = (request.POST.get('next') or '').strip()
    if nxt:
        if nxt in _SAFE_NEXT_NAMES:
            return redirect(nxt)
        if url_has_allowed_host_and_scheme(nxt, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
            return redirect(nxt)
    return redirect(fallback, **kwargs)


# =============================================================================
# Dashboard
# =============================================================================
def _handle_dashboard_post(request, setting, year_int, dash_url):
    """จัดการคำสั่ง POST ทั้งการปรับแต่งและปุ่มเปิด/ปิดระบบบนแดชบอร์ด"""
    from food.exam_seating_service import check_plan_ready_to_open
    from food.models import (
        CampSurvey, CertificateYearStatus, ExamSeatingPlan, FoodOrder,
        ScoreApproval, ScoreCategory, approve_year_scores, unapprove_year_scores,
    )
    _, _, _, _, year_locked_message = _lock_helpers()
    post_action = request.POST.get('action', '')

    if post_action == 'save_dashboard_config':
        widgets = request.POST.getlist('dashboard_widgets')
        selected_dash = request.POST.getlist('dashboard_summary_fields')
        setting.set_dashboard_widgets(widgets)
        setting.set_dashboard_summary_fields(selected_dash)
        setting.save(update_fields=['dashboard_widgets', 'dashboard_summary_fields', 'updated_at'])
        log_action(request.user, 'update_settings', target=setting, new_value={
            'dashboard_widgets': setting.get_dashboard_widgets(),
            'dashboard_summary_fields': setting.get_dashboard_summary_fields(),
        }, request=request)
        messages.success(request, 'บันทึกการปรับแต่งแดชบอร์ดเรียบร้อยแล้ว')
        return redirect(dash_url)

    if post_action == 'reset_dashboard_config':
        default_summary = ['is_muslim', 'food_allergy', 'chronic_disease', 'student_phone', 'parent_phone', 'email']
        setting.set_dashboard_widgets(DEFAULT_DASHBOARD_WIDGETS)
        setting.set_dashboard_summary_fields(default_summary)
        setting.save(update_fields=['dashboard_widgets', 'dashboard_summary_fields', 'updated_at'])
        log_action(request.user, 'update_settings', target=setting, new_value={'reset_dashboard': True}, request=request)
        messages.success(request, 'คืนค่าเริ่มต้นการแสดงผลแดชบอร์ดแล้ว')
        return redirect(dash_url)

    if post_action == 'save_dashboard_summary':
        selected_dash = request.POST.getlist('dashboard_summary_fields')
        setting.set_dashboard_summary_fields(selected_dash)
        setting.save(update_fields=['dashboard_summary_fields', 'updated_at'])
        log_action(request.user, 'update_settings', target=setting,
                   new_value={'dashboard_summary_fields': selected_dash}, request=request)
        messages.success(request, 'บันทึกการตั้งค่าการแสดงผลสรุปข้อมูลนักเรียนบน Dashboard เรียบร้อย')
        return redirect(dash_url)

    if post_action == 'dash_toggle_cert_publish':
        cert_status, _ = CertificateYearStatus.objects.get_or_create(academic_year=year_int)
        old_pub = cert_status.is_published
        cert_status.is_published = not old_pub
        cert_status.updated_by = request.user
        cert_status.save(update_fields=['is_published', 'updated_by', 'updated_at'])
        log_action(
            request.user, 'toggle_cert_publish', target=cert_status,
            old_value={'academic_year': year_int, 'is_published': old_pub},
            new_value={'academic_year': year_int, 'is_published': cert_status.is_published, 'from_dashboard': True},
            request=request,
        )
        if cert_status.is_published:
            messages.success(request, f'เปิดแสดงเกียรติบัตรให้นักเรียนประจำปีการศึกษา {year_int} เรียบร้อยแล้ว')
        else:
            messages.info(request, f'ปิดการแสดงเกียรติบัตรให้นักเรียนประจำปีการศึกษา {year_int} เรียบร้อยแล้ว')
        return redirect(dash_url)

    if post_action == 'dash_approve_scores':
        if ScoreApproval.is_year_approved(year_int):
            messages.info(request, f'ปีการศึกษา {year_int} อนุมัติประกาศผลแล้ว')
        else:
            appr, ranked = approve_year_scores(year_int, request.user)
            passed_n = sum(1 for r in ranked if r['passed'])
            log_action(
                request.user, 'update_settings', target=appr,
                new_value={'approve_scores': year_int, 'total': len(ranked), 'passed': passed_n, 'from_dashboard': True},
                request=request,
            )
            messages.success(
                request,
                f'อนุมัติประกาศผลคะแนนปีการศึกษา {year_int} แล้ว '
                f'(ทั้งหมด {len(ranked)} คน ผ่านการคัดเลือกเข้าค่าย 2 จำนวน {passed_n} คน)',
            )
        return redirect(dash_url)

    if post_action == 'dash_unapprove_scores':
        if not ScoreApproval.is_year_approved(year_int):
            messages.info(request, f'ปีการศึกษา {year_int} ยังไม่ได้อนุมัติประกาศผล')
        else:
            confirm_password = request.POST.get('confirm_password') or ''
            if not confirm_password:
                messages.error(request, 'กรุณากรอกรหัสผ่านปัจจุบันของผู้ดูแลระบบเพื่อยืนยันการยกเลิกประกาศผล')
            elif not request.user.check_password(confirm_password):
                log_action(
                    request.user, 'update_settings', target_model='ScoreApproval',
                    new_value={'unapprove_scores_failed': year_int, 'reason': 'invalid_password', 'from_dashboard': True},
                    request=request,
                )
                messages.error(request, 'รหัสผ่านปัจจุบันของผู้ดูแลระบบไม่ถูกต้อง ไม่สามารถยกเลิกประกาศผลได้')
            else:
                unapprove_year_scores(year_int)
                log_action(
                    request.user, 'update_settings', target_model='ScoreApproval',
                    new_value={'unapprove_scores': year_int, 'password_verified': True, 'from_dashboard': True},
                    request=request,
                )
                messages.success(request, f'ยกเลิกการอนุมัติประกาศผลคะแนนปีการศึกษา {year_int} แล้ว (ปลดล็อกระบบเรียบร้อย)')
        return redirect(dash_url)

    if post_action == 'dash_toggle_exam_plan':
        plan = get_object_or_404(ExamSeatingPlan, pk=request.POST.get('plan_id'))
        if not plan.is_open:
            readiness = check_plan_ready_to_open(plan)
            if not readiness['is_ready']:
                messages.error(
                    request,
                    f'ไม่สามารถเปิดให้ผู้เข้าสอบดูที่นั่งสอบในวันนี้ได้ เนื่องจากยังจัดห้องไม่เรียบร้อย: {readiness["reason_text"]}',
                )
                return redirect(dash_url)
        plan.is_open = not plan.is_open
        plan.save(update_fields=['is_open', 'updated_at'])
        log_action(
            request.user, 'exam_plan_toggle', target=plan,
            new_value={'is_open': plan.is_open, 'from_dashboard': True}, request=request,
        )
        state_lbl = 'เปิดให้ผู้เข้าสอบดูที่นั่งสอบในวันดังกล่าวแล้ว' if plan.is_open else 'ปิดการแสดงที่นั่งสอบในวันดังกล่าวแล้ว'
        messages.success(request, state_lbl)
        return redirect(dash_url)

    if post_action == 'dash_toggle_food_system':
        setting.food_system_open = not setting.food_system_open
        setting.save(update_fields=['food_system_open', 'updated_at'])
        log_action(
            request.user, 'update_settings', target=setting,
            new_value={'food_system_open': setting.food_system_open, 'from_dashboard': True}, request=request,
        )
        messages.success(request, f'{"เปิด" if setting.food_system_open else "ปิด"}ระบบเลือกอาหารสำหรับสมาชิกแล้ว')
        return redirect(dash_url)

    if post_action == 'dash_toggle_food_order':
        order = get_object_or_404(FoodOrder, pk=request.POST.get('order_id'))
        order.edit_toggle_open = not order.edit_toggle_open
        order.save(update_fields=['edit_toggle_open', 'updated_at'])
        log_action(
            request.user, 'food_order', target=order,
            new_value={'action': 'toggle', 'open': order.edit_toggle_open, 'from_dashboard': True}, request=request,
        )
        messages.success(request, f'{"เปิด" if order.edit_toggle_open else "ปิด"}ให้เลือกรายการอาหาร "{order.restaurant_name}" แล้ว')
        return redirect(dash_url)

    if post_action == 'dash_toggle_survey':
        survey = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'))
        survey.is_open = not survey.is_open
        survey.save(update_fields=['is_open', 'updated_at'])
        log_action(
            request.user, 'update_settings', target=survey,
            new_value={'survey_id': survey.pk, 'is_open': survey.is_open, 'from_dashboard': True}, request=request,
        )
        messages.success(request, f'{"เปิด" if survey.is_open else "ปิด"}แบบสอบถาม "{survey.title}" แล้ว')
        return redirect(dash_url)

    if post_action == 'dash_toggle_rights_round':
        round_id = request.POST.get('round_id')
        camp_level = request.POST.get('camp_level')
        try:
            if round_id:
                rr = get_object_or_404(RightsRound, pk=round_id)
                if rr.is_open:
                    waived, _ = rights.close_round(rr, request.user, mark_waived=True, request=request)
                    messages.success(request, f'ปิดรอบใช้สิทธิ์ค่าย {rr.camp_level} ปี {rr.academic_year} แล้ว (สละสิทธิ์ {waived} คน)')
                else:
                    rights.reopen_round(rr, request.user, request=request)
                    messages.success(request, f'เปิดรอบใช้สิทธิ์ค่าย {rr.camp_level} ปี {rr.academic_year} ใหม่อีกครั้งแล้ว')
            elif camp_level in ('1', '2'):
                rr = rights.open_round(year_int, int(camp_level), request.user, title='', request=request)
                messages.success(request, f'เปิดรอบใช้สิทธิ์ค่าย {rr.camp_level} ปีการศึกษา {rr.academic_year} แล้ว')
        except rights.RightsError as e:
            messages.error(request, str(e))
        return redirect(dash_url)

    if post_action == 'dash_toggle_score_category':
        cat = get_object_or_404(ScoreCategory, pk=request.POST.get('cat_id'))
        field = (request.POST.get('field') or '').strip()
        if field in ('is_open', 'is_visible'):
            if ScoreApproval.is_year_approved(cat.academic_year):
                messages.error(request, year_locked_message(cat.academic_year))
            else:
                setattr(cat, field, not getattr(cat, field))
                cat.save(update_fields=[field])
                log_action(
                    request.user, 'update_settings', target=cat,
                    new_value={'category': cat.name, field: getattr(cat, field), 'from_dashboard': True},
                    request=request,
                )
                lbl = 'เปิดให้กรอกคะแนน' if field == 'is_open' else 'แสดงผลคะแนน'
                st = 'เปิด' if getattr(cat, field) else 'ปิด'
                messages.success(request, f'{st}{lbl}หมวด "{cat.name}" แล้ว')
        return redirect(dash_url)

    if post_action == 'dash_toggle_profile_edit_all':
        current_fields = setting.get_editable_fields()
        if current_fields:
            setting.set_editable_fields([])
            msg = 'ปิดการแก้ไขข้อมูลส่วนตัวสมาชิกทั้งหมดแล้ว'
        else:
            setting.set_editable_fields(['student_phone', 'parent_phone', 'email', 'chronic_disease', 'food_allergy', 'is_muslim'])
            msg = 'เปิดให้สมาชิกแก้ไขข้อมูลส่วนตัวพื้นฐานแล้ว'
        setting.save(update_fields=['editable_profile_fields', 'updated_at'])
        log_action(
            request.user, 'update_settings', target=setting,
            new_value={'editable_profile_fields': setting.get_editable_fields(), 'from_dashboard': True},
            request=request,
        )
        messages.success(request, msg)
        return redirect(dash_url)

    return redirect(dash_url)


def _build_dashboard_quick_toggles(setting, year_int, is_approved, cert_published):
    """รวบรวมสถานะรายการเปิด/ปิดทั้งหมดสำหรับแสดงในศูนย์ควบคุมบนแดชบอร์ด"""
    from food.exam_seating_service import check_plan_ready_to_open, get_exam_examinees
    from food.models import CampSurvey, FoodOrder, ScoreCategory, TrainingDate

    exam_dates = list(
        TrainingDate.objects.filter(academic_year=year_int, event_type=TrainingDate.TYPE_EXAM)
        .prefetch_related('seating_plans')
        .order_by('date', 'id')
    )
    exam_examinees = get_exam_examinees(year_int) if exam_dates else []
    exam_plan_items = []
    for td in exam_dates:
        plans = sorted(td.seating_plans.all(), key=lambda p: p.id)
        if plans:
            for p in plans:
                exam_plan_items.append({
                    'training_date': td,
                    'plan': p,
                    'readiness': check_plan_ready_to_open(p, examinees=exam_examinees),
                    'total_examinees': len(exam_examinees),
                })
        else:
            exam_plan_items.append({
                'training_date': td,
                'plan': None,
                'readiness': None,
                'total_examinees': len(exam_examinees),
            })

    rights_camp_items = []
    for camp_lv in (1, 2):
        rr = (
            RightsRound.objects.filter(academic_year=year_int, camp_level=camp_lv)
            .annotate(
                confirmed=Count('use_records', filter=Q(use_records__is_used=True, use_records__is_deleted=False)),
            )
            .order_by('-opened_at')
            .first()
        )
        rights_camp_items.append({
            'camp_level': camp_lv,
            'round': rr,
            'confirmed': rr.confirmed if rr else 0,
        })

    return {
        'is_approved': is_approved,
        'cert_published': cert_published,
        'food_system_open': setting.food_system_open,
        'editable_fields_count': len(setting.get_editable_fields()),
        'total_editable_fields_count': len(PROFILE_EDITABLE_CHOICES),
        'exam_plan_items': exam_plan_items,
        'rights_camp_items': rights_camp_items,
        'food_order_items': list(FoodOrder.objects.filter(academic_year=year_int).order_by('-food_date', '-id')[:8]),
        'survey_items': list(CampSurvey.objects.filter(academic_year=year_int).order_by('-created_at', '-id')[:8]),
        'score_cat_items': list(ScoreCategory.objects.filter(academic_year=year_int, parent__isnull=True).order_by('order', 'id')),
    }


@admin_required
def dashboard(request):
    from decimal import Decimal
    from django.urls import reverse
    from food.certificate_pdf import get_certificate_data_for_students
    from food.models import (
        Attendance, CampSurvey, CertificateRecipient, CertificateYearStatus, FoodOrder, QualificationResult,
        ScoreApproval, ScoreCategory, TrainingDate, build_survey_summary, compute_year_ranking,
        non_training_dates_for_year,
    )
    ScoreApproval, _, _, _, _ = _lock_helpers()
    setting = SystemSetting.get()
    year_req = (request.GET.get('year') or request.POST.get('year') or '').strip()
    try:
        year_int = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year_int = int(setting.current_academic_year)
    year = str(year_int)

    if request.method == 'POST':
        dash_url = reverse('composn:admin_dashboard') + (f'?year={year}' if year_req else '')
        return _handle_dashboard_post(request, setting, year_int, dash_url)

    active_widgets = setting.get_dashboard_widgets()
    active_widget_set = set(active_widgets)
    widget_labels = dict(DASHBOARD_WIDGET_CHOICES)
    dashboard_widget_config = [
        {'key': k, 'label': widget_labels[k], 'enabled': True}
        for k in active_widgets
        if k in widget_labels
    ] + [
        {'key': k, 'label': lbl, 'enabled': False}
        for k, lbl in DASHBOARD_WIDGET_CHOICES
        if k not in active_widget_set
    ]

    yy = f'{year_int % 100:02d}'
    members = User.objects.filter(role=User.ROLE_MEMBER).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=year_int)
    ).order_by('camp1_code')

    selected_summary_keys = setting.get_dashboard_summary_fields()
    all_summary_defs = {
        'is_muslim': {
            'key': 'is_muslim',
            'label': 'ศาสนาอิสลาม (อาหารฮาลาล)',
            'icon': 'bi-moon-stars-fill',
            'color': 'success',
            'qs': members.filter(is_muslim=True),
            'value_attr': None,
        },
        'food_allergy': {
            'key': 'food_allergy',
            'label': 'คนที่แพ้อาหาร',
            'icon': 'bi-exclamation-triangle-fill',
            'color': 'danger',
            'qs': members.exclude(food_allergy=''),
            'value_attr': 'food_allergy',
        },
        'chronic_disease': {
            'key': 'chronic_disease',
            'label': 'คนที่มีโรคประจำตัว',
            'icon': 'bi-heart-pulse-fill',
            'color': 'warning',
            'qs': members.exclude(chronic_disease=''),
            'value_attr': 'chronic_disease',
        },
        'student_phone': {
            'key': 'student_phone',
            'label': 'กรอกเบอร์โทรนักเรียนแล้ว',
            'icon': 'bi-telephone-fill',
            'color': 'primary',
            'qs': members.exclude(student_phone=''),
            'value_attr': 'student_phone',
        },
        'parent_phone': {
            'key': 'parent_phone',
            'label': 'กรอกเบอร์โทรผู้ปกครองแล้ว',
            'icon': 'bi-telephone-plus-fill',
            'color': 'info',
            'qs': members.exclude(parent_phone=''),
            'value_attr': 'parent_phone',
        },
        'email': {
            'key': 'email',
            'label': 'กรอกอีเมลแล้ว',
            'icon': 'bi-envelope-check-fill',
            'color': 'secondary',
            'qs': members.exclude(email=''),
            'value_attr': 'email',
        },
        'profile_photo': {
            'key': 'profile_photo',
            'label': 'อัปโหลดรูปประจำตัวแล้ว',
            'icon': 'bi-person-bounding-box',
            'color': 'primary',
            'qs': members.exclude(profile_photo='').exclude(profile_photo__isnull=True),
            'value_attr': None,
            'default_detail': 'มีรูปประจำตัว',
        },
        'camp2_code': {
            'key': 'camp2_code',
            'label': 'ได้รับรหัสค่าย 2 แล้ว',
            'icon': 'bi-award-fill',
            'color': 'danger',
            'qs': members.exclude(camp2_code__isnull=True).exclude(camp2_code=''),
            'value_attr': 'camp2_code',
        },
    }
    dashboard_summary_items = []
    if 'student_summary' in active_widget_set:
        for k in selected_summary_keys:
            d = all_summary_defs.get(k)
            if not d:
                continue
            u_list = list(d['qs'][:50])
            details = [
                {
                    'code': u.camp1_code,
                    'name': u.get_full_name(),
                    'detail': getattr(u, d['value_attr'], '') if d['value_attr'] else d.get('default_detail', 'อิสลาม'),
                }
                for u in u_list
            ]
            dashboard_summary_items.append({
                'key': k,
                'label': d['label'],
                'icon': d['icon'],
                'color': d['color'],
                'count': d['qs'].count(),
                'details': details,
            })

    att_stats = {}
    if 'attendance_summary' in active_widget_set:
        td_qs = TrainingDate.objects.filter(academic_year=year_int)
        non_train = non_training_dates_for_year(year_int)
        train_att_qs = Attendance.objects.filter(academic_year=year_int)
        if non_train:
            train_att_qs = train_att_qs.exclude(date__in=non_train)
        over_rows = list(
            train_att_qs.filter(status__in=Attendance.ABSENT_STATUSES)
            .values('user_id').annotate(cnt=Count('id')).filter(cnt__gt=setting.max_absences)
            .order_by('-cnt')[:20]
        )
        over_users = User.objects.in_bulk([r['user_id'] for r in over_rows]) if over_rows else {}
        over_list = [
            {'user': over_users[r['user_id']], 'count': r['cnt']}
            for r in over_rows if r['user_id'] in over_users
        ]
        att_stats = {
            'training_days': td_qs.filter(event_type=TrainingDate.TYPE_TRAINING).count(),
            'exam_days': td_qs.filter(event_type=TrainingDate.TYPE_EXAM).count(),
            'holiday_days': td_qs.filter(event_type=TrainingDate.TYPE_HOLIDAY).count(),
            'present': train_att_qs.filter(status=Attendance.STATUS_PRESENT).count(),
            'absent': train_att_qs.filter(status=Attendance.STATUS_ABSENT).count(),
            'sick': train_att_qs.filter(status=Attendance.STATUS_SICK).count(),
            'personal': train_att_qs.filter(status=Attendance.STATUS_PERSONAL).count(),
            'max_absences': setting.max_absences,
            'over_absence_list': over_list,
        }

    need_approval_status = bool({'quick_toggles', 'scores_certificates'} & active_widget_set)
    is_year_approved = ScoreApproval.is_year_approved(year_int) if need_approval_status else False
    cert_published = CertificateYearStatus.is_year_published(year_int) if need_approval_status else False

    quick_toggles = {}
    if 'quick_toggles' in active_widget_set:
        quick_toggles = _build_dashboard_quick_toggles(setting, year_int, is_year_approved, cert_published)

    score_cert_stats = {}
    if 'scores_certificates' in active_widget_set:
        main_cats = list(ScoreCategory.objects.filter(academic_year=year_int, parent__isnull=True))
        if is_year_approved:
            passed_cnt = QualificationResult.objects.filter(academic_year=year_int, passed=True).count()
        else:
            ranked_preview = compute_year_ranking(year_int, only_visible=False)
            passed_cnt = sum(1 for r in ranked_preview if r['passed'])
        cert_data = get_certificate_data_for_students(year_int)
        score_cert_stats = {
            'cat_count': len(main_cats),
            'total_weight': sum((c.weight for c in main_cats), Decimal('0')),
            'is_approved': is_year_approved,
            'passed_count': passed_cnt,
            'max_qualified_count': setting.max_qualified_count,
            'passing_score': setting.passing_score,
            'cert_eligible_count': sum(1 for s in cert_data if s['is_current'] and s['is_eligible']),
            'cert_ineligible_count': sum(1 for s in cert_data if s['is_current'] and not s['is_eligible']),
            'cert_rights_count': sum(1 for s in cert_data if not s['is_current']),
            'cert_published': cert_published,
            'lecturer_count': CertificateRecipient.objects.filter(
                academic_year=year_int, recipient_type=CertificateRecipient.TYPE_LECTURER
            ).count(),
            'committee_count': CertificateRecipient.objects.filter(
                academic_year=year_int, recipient_type=CertificateRecipient.TYPE_COMMITTEE
            ).count(),
        }

    survey_summaries = []
    if 'survey_summary' in active_widget_set:
        surveys = list(CampSurvey.objects.filter(academic_year=year).prefetch_related('questions', 'answers__user'))
        survey_summaries = [build_survey_summary(s) for s in surveys]

    recent_orders = list(
        FoodOrder.objects.filter(academic_year=year)
        .annotate(selected_users_count=Count('selections__user', distinct=True))[:6]
    ) if 'food_orders' in active_widget_set else []

    pending_school_req_list = list(
        SchoolChangeRequest.objects.filter(status='pending').select_related('user').order_by('-created_at')[:10]
    ) if 'school_requests' in active_widget_set else []

    open_rounds = list(RightsRound.objects.filter(is_open=True, academic_year=year))
    for rr in open_rounds:
        rr.confirmed = rr.use_records.filter(is_used=True).count()

    used_qs = UseRecord.objects.filter(academic_year=year, is_used=True)
    ctx = {
        'setting': setting,
        'member_count': members.count(),
        'member_active': members.filter(is_active=True).count(),
        'camp2_count': members.filter(camp2_code__isnull=False).count(),
        'ds_count': members.filter(camp_center='DS').count(),
        'km_count': members.filter(camp_center='KM').count(),
        'staff_count': User.objects.filter(role__in=[User.ROLE_STAFF, User.ROLE_ADMIN]).count(),
        'open_rounds': open_rounds,
        'used_this_year': used_qs.count(),
        'used_camp1': used_qs.filter(camp_level=1).count(),
        'used_camp2': used_qs.filter(camp_level=2).count(),
        'waived_this_year': UseRecord.objects.filter(academic_year=year, is_used=False).count(),
        'pending_requests': SchoolChangeRequest.objects.filter(status='pending').count(),
        'pending_school_req_list': pending_school_req_list,
        'absent_this_year': Attendance.objects.filter(academic_year=year, status='absent').count(),
        'food_active': FoodOrder.objects.filter(academic_year=year, edit_toggle_open=True,
                                                food_date__gte=timezone.localdate()).count(),
        'food_total_count': FoodOrder.objects.filter(academic_year=year).count(),
        'recent_orders': recent_orders,
        'recent_logs': AuditLog.objects.select_related('user')[:10] if 'recent_logs' in active_widget_set else [],
        'action_labels': ACTION_LABELS,
        'year': year,
        'active_dashboard_widgets': active_widgets,
        'dashboard_widget_config': dashboard_widget_config,
        'dashboard_summary_choices': DASHBOARD_SUMMARY_CHOICES,
        'selected_summary_keys': selected_summary_keys,
        'dashboard_summary_items': dashboard_summary_items,
        'att_stats': att_stats,
        'score_cert_stats': score_cert_stats,
        'survey_summaries': survey_summaries,
        'quick_toggles': quick_toggles,
        'is_year_approved': is_year_approved,
    }
    return render(request, 'admin/dashboard.html', ctx)


# =============================================================================
# จัดการผู้ใช้ / ผู้ดูแลและเจ้าหน้าที่ / ตั้งค่าหน้าผู้ใช้
# =============================================================================
@admin_required
def users_settings(request):
    """ตั้งค่าหน้าผู้ใช้ (กำหนดฟิลด์ที่สมาชิกแก้ไขเองได้ + ฟิลด์ที่สรุปบน Dashboard + คำขอแก้ไขชื่อโรงเรียน)"""
    setting = SystemSetting.get()
    if request.method == 'POST':
        fields = request.POST.getlist('editable_fields')
        setting.set_editable_fields(fields)
        if 'save_dashboard_fields' in request.POST:
            dash_fields = request.POST.getlist('dashboard_summary_fields')
            setting.set_dashboard_summary_fields(dash_fields)
        setting.save()
        log_action(request.user, 'update_settings', target=setting,
                   new_value={'editable_profile_fields': fields}, request=request)
        messages.success(request, 'บันทึกการตั้งค่าหน้าผู้ใช้เรียบร้อย')
        return redirect('composn:admin_users_settings')
    return render(request, 'admin/users_settings.html', {
        'setting': setting,
        'choices': PROFILE_EDITABLE_CHOICES,
        'selected': setting.get_editable_fields(),
        'dashboard_choices': DASHBOARD_SUMMARY_CHOICES,
        'dashboard_selected': setting.get_dashboard_summary_fields(),
        'pending_requests': SchoolChangeRequest.objects.filter(status='pending').count(),
    })


@admin_required
def admins_list(request):
    """จัดการผู้ดูแลและเจ้าหน้าที่ (เพิ่ม ลบ แก้ไข ค้นหา)"""
    qs = User.objects.filter(role__in=[User.ROLE_ADMIN, User.ROLE_STAFF]).order_by('role', 'camp1_code')
    q = request.GET.get('q', '').strip()
    role = request.GET.get('role', '')
    if q:
        qs = qs.filter(Q(camp1_code__icontains=q) | Q(first_name__icontains=q) | Q(last_name__icontains=q))
    if role in (User.ROLE_ADMIN, User.ROLE_STAFF):
        qs = qs.filter(role=role)
    return render(request, 'admin/admins.html', {
        'admins': qs, 'q': q, 'role': role,
    })


@admin_required
def admin_create(request):
    """เพิ่มผู้ดูแลระบบหรือเจ้าหน้าที่ใหม่ (มีเฉพาะ รหัสผู้ใช้ บทบาท และรหัสผ่าน)"""
    form = AdminUserCreateForm(request.POST or None, initial={'role': User.ROLE_STAFF})
    if request.method == 'POST' and form.is_valid():
        code = form.cleaned_data['camp1_code']
        role = form.cleaned_data['role']
        pw = form.cleaned_data['password']
        is_adm = (role == User.ROLE_ADMIN)
        user = User(
            camp1_code=code,
            role=role,
            is_active=True,
            is_first_login=True,
            is_superuser=is_adm,
            staff_can_view_users=True if is_adm else bool(form.cleaned_data.get('staff_can_view_users')),
            staff_can_attendance=True if is_adm else bool(form.cleaned_data.get('staff_can_attendance')),
            staff_can_signin_sheet=True if is_adm else bool(form.cleaned_data.get('staff_can_signin_sheet')),
            staff_can_manage_food=True if is_adm else bool(form.cleaned_data.get('staff_can_manage_food')),
            staff_can_scores=True if is_adm else bool(form.cleaned_data.get('staff_can_scores')),
            staff_can_exam_seating=True if is_adm else bool(form.cleaned_data.get('staff_can_exam_seating')),
        )
        user.set_password(pw)
        user.save()
        log_action(request.user, 'create_user', target=user,
                   new_value={'camp1_code': user.camp1_code, 'role': user.role,
                              'permissions': user.staff_permission_labels()}, request=request)
        messages.success(request, f'เพิ่มบัญชี {user.get_role_display()} "{user.camp1_code}" เรียบร้อย')
        return redirect('composn:admin_admins')
    return render(request, 'admin/user_form.html', {
        'form': form, 'title': 'เพิ่มผู้ดูแลระบบ / เจ้าหน้าที่',
        'cancel_url': 'composn:admin_admins',
        'staff_perm_fields': STAFF_PERM_FIELDS,
    })


@staff_perm_required('staff_can_view_users')
def users(request):
    from django.db.models import Case, IntegerField, Value, When

    is_admin_viewer = request.user.is_admin_role
    setting = SystemSetting.get()
    year_req = request.GET.get('year', '').strip()
    try:
        year_int = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year_int = int(setting.current_academic_year)
    yy = f'{year_int % 100:02d}'
    year = str(year_int)

    qs = User.objects.all().annotate(
        used_count=Count('archive_records', filter=Q(archive_records__is_deleted=False))
    )
    q = request.GET.get('q', '').strip()
    role = request.GET.get('role', '') if is_admin_viewer else ''
    camp = request.GET.get('camp', '')
    active = request.GET.get('active', '')
    pw_status = request.GET.get('pw_status', '')
    health = request.GET.get('health', '')
    tab = request.GET.get('tab', 'current')  # 'current', 'rights', 'all'
    if tab not in ('current', 'rights', 'all'):
        tab = 'current'

    # แยกกลุ่มนักเรียนปัจจุบัน vs นักเรียนใช้สิทธิ์ในปีการศึกษานี้
    used_ids = set(UseRecord.objects.filter(academic_year=year_int, is_used=True).values_list('user_id', flat=True))
    is_current_q = Q(camp1_code__istartswith=yy) | Q(academic_year=year_int)
    is_rights_q = Q(pk__in=used_ids) & ~is_current_q

    if is_admin_viewer and role in (User.ROLE_ADMIN, User.ROLE_STAFF):
        qs = qs.filter(role=role)
        current_count = 0
        rights_count = 0
        all_count = qs.count()
    else:
        member_base = qs.filter(role=User.ROLE_MEMBER).filter(is_current_q | Q(pk__in=used_ids))
        current_count = member_base.filter(is_current_q).count()
        rights_count = member_base.filter(is_rights_q).count()
        all_count = member_base.count()

        if tab == 'current':
            qs = member_base.filter(is_current_q)
        elif tab == 'rights':
            qs = member_base.filter(is_rights_q)
        else:
            qs = member_base

    if q:
        qs = qs.filter(Q(camp1_code__icontains=q) | Q(camp2_code__icontains=q) | Q(first_name__icontains=q) |
                       Q(last_name__icontains=q) | Q(nickname__icontains=q) | Q(school_name__icontains=q))
    if camp == '2':
        qs = qs.filter(camp2_code__isnull=False).exclude(camp2_code='')
    elif camp == '1':
        qs = qs.filter(Q(camp2_code__isnull=True) | Q(camp2_code=''))
    if active in ('1', '0'):
        qs = qs.filter(is_active=active == '1')
    if pw_status == 'first':
        qs = qs.filter(is_first_login=True)
    elif pw_status == 'changed':
        qs = qs.filter(is_first_login=False)
    if health == 'disease':
        qs = qs.exclude(chronic_disease='')
    elif health == 'allergy':
        qs = qs.exclude(food_allergy='')
    elif health == 'muslim':
        qs = qs.filter(is_muslim=True)

    # เรียงลำดับให้นักเรียนปัจจุบันขึ้นก่อน แล้วตามด้วยนักเรียนใช้สิทธิ์ แล้วเรียงตามรหัสค่าย
    qs = qs.annotate(
        cohort_order=Case(
            When(is_current_q, then=Value(0)),  # นักเรียนปัจจุบัน = 0
            default=Value(1),                   # นักเรียนใช้สิทธิ์ = 1
            output_field=IntegerField(),
        )
    ).order_by('cohort_order', 'camp1_code')

    if request.GET.get('export') == 'xlsx':
        if not is_admin_viewer:
            raise PermissionDenied('คุณไม่มีสิทธิ์ Export ข้อมูลผู้ใช้')
        log_action(request.user, 'export', target_model='User', new_value='รายชื่อผู้ใช้', request=request)
        return workbook_response(students_workbook(qs), f'รายชื่อนักเรียน_{year}_{tab}.xlsx')

    page = _paginate(request, qs)
    for u in page:
        u.display_level = u.level_for_year(year_int) or u.current_level
        u.is_current_student = bool(u.camp1_code.upper().startswith(yy) or (u.academic_year == year_int))

    params = request.GET.copy()
    params.pop('page', None)

    def tab_url(t):
        p = params.copy()
        p['tab'] = t
        return '?' + p.urlencode()

    return render(request, 'admin/users.html', {
        'page': page, 'q': q, 'role': role, 'year': year,
        'camp': camp, 'active': active, 'pw_status': pw_status, 'health': health,
        'tab': tab,
        'current_count': current_count,
        'rights_count': rights_count,
        'all_count': all_count,
        'tab_url_current': tab_url('current'),
        'tab_url_rights': tab_url('rights'),
        'tab_url_all': tab_url('all'),
        'role_choices': User.ROLE_CHOICES, 'querystring': params.urlencode(),
        'is_admin_viewer': is_admin_viewer,
    })


@admin_required
def user_create(request):
    form = UserForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        photo_cf = None
        if 'profile_photo' in request.FILES:
            try:
                photo_cf, _ = process_and_validate_profile_photo(
                    request.FILES['profile_photo'], user_code=form.cleaned_data.get('camp1_code') or 'user'
                )
            except FacePhotoError as e:
                messages.error(request, str(e))
                return render(request, 'admin/user_form.html', {
                    'form': form, 'title': 'เพิ่มผู้ใช้', 'staff_perm_fields': STAFF_PERM_FIELDS,
                    'allow_photo_upload': True,
                })
        user = form.save(commit=False)
        if user.role == User.ROLE_MEMBER:
            ScoreApproval, _, _, _, year_locked_message = _lock_helpers()
            base_y = user.base_academic_year
            if base_y and ScoreApproval.is_year_approved(base_y):
                messages.error(request, year_locked_message(base_y))
                return render(request, 'admin/user_form.html', {
                    'form': form, 'title': 'เพิ่มผู้ใช้', 'staff_perm_fields': STAFF_PERM_FIELDS,
                    'allow_photo_upload': True,
                })
        if form.cleaned_data.get('auto_levels'):
            levels.apply_levels(user)
        pw = form.cleaned_data.get('password')
        if not pw:
            pw = user.camp1_code if user.role == User.ROLE_MEMBER else '1234'
        user.set_password(pw)
        user.is_first_login = True
        user.is_superuser = user.role == User.ROLE_ADMIN
        user.save()
        if photo_cf is not None:
            user.profile_photo.save(photo_cf.name, photo_cf, save=True)
        log_action(request.user, 'create_user', target=user,
                   new_value=model_snapshot(user, ['camp1_code', 'role', 'first_name', 'last_name']),
                   request=request)
        messages.success(request, f'เพิ่มผู้ใช้ {user.camp1_code} เรียบร้อย')
        return redirect('composn:admin_user_detail', pk=user.pk)
    return render(request, 'admin/user_form.html', {
        'form': form, 'title': 'เพิ่มผู้ใช้', 'staff_perm_fields': STAFF_PERM_FIELDS,
        'allow_photo_upload': True,
    })


@admin_required
def user_edit(request, pk):
    user = get_object_or_404(User, pk=pk)
    before = model_snapshot(user, USER_FIELDS)
    orig_camp1 = user.camp1_code
    orig_camp2 = user.camp2_code
    orig_year = user.academic_year
    orig_active = user.is_active
    orig_role = user.role
    ScoreApproval, _, camp2_code_year, user_approved_years, _ = _lock_helpers()
    orig_locked_years = user_approved_years(user)
    orig_c2_year = camp2_code_year(orig_camp2)
    cancel_url = 'composn:admin_users' if user.role == User.ROLE_MEMBER else 'composn:admin_admins'
    form = UserForm(request.POST or None, instance=user, initial={'auto_levels': False})
    if request.method == 'POST' and form.is_valid():
        new_c2 = (form.cleaned_data.get('camp2_code') or '').strip().upper() or None
        new_c2_yr = camp2_code_year(new_c2)
        if new_c2 != orig_camp2:
            if ((orig_c2_year and ScoreApproval.is_year_approved(orig_c2_year))
                    or (new_c2_yr and ScoreApproval.is_year_approved(new_c2_yr))
                    or orig_locked_years):
                bad_yr = (
                    (orig_c2_year if (orig_c2_year and ScoreApproval.is_year_approved(orig_c2_year)) else None)
                    or (new_c2_yr if (new_c2_yr and ScoreApproval.is_year_approved(new_c2_yr)) else None)
                    or orig_locked_years[0]
                )
                messages.error(request, f'ปีการศึกษา {bad_yr} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนแปลงรหัสค่าย 2 ได้ กรุณายกเลิกประกาศผลก่อน')
                return redirect('composn:admin_user_edit', pk=pk)
        if orig_locked_years:
            if (form.cleaned_data.get('camp1_code') != orig_camp1
                    or form.cleaned_data.get('academic_year') != orig_year
                    or (orig_active and not form.cleaned_data.get('is_active'))
                    or form.cleaned_data.get('role') != orig_role):
                yr_txt = ', '.join(str(y) for y in orig_locked_years)
                messages.error(
                    request,
                    f'ผู้ใช้มีข้อมูลในปีการศึกษาที่อนุมัติประกาศผลแล้ว ({yr_txt}) '
                    f'ไม่สามารถแก้ไขรหัสค่าย 1 ปีการศึกษา บทบาท หรือปิดใช้งานบัญชีได้',
                )
                return redirect('composn:admin_user_edit', pk=pk)
        if user.pk == request.user.pk and (form.cleaned_data['role'] != User.ROLE_ADMIN or
                                           not form.cleaned_data['is_active']):
            messages.error(request, 'ไม่สามารถลดสิทธิ์หรือปิดการใช้งานบัญชีของตนเองได้')
            return redirect('composn:admin_user_edit', pk=pk)
        photo_cf = None
        if 'profile_photo' in request.FILES:
            try:
                photo_cf, _ = process_and_validate_profile_photo(
                    request.FILES['profile_photo'], user_code=form.cleaned_data.get('camp1_code') or user.camp1_code
                )
            except FacePhotoError as e:
                messages.error(request, str(e))
                return render(request, 'admin/user_form.html', {
                    'form': form, 'title': f'แก้ไขผู้ใช้ {user.camp1_code}', 'obj': user,
                    'cancel_url': cancel_url, 'staff_perm_fields': STAFF_PERM_FIELDS,
                    'allow_photo_upload': True,
                })
        user = form.save(commit=False)
        if form.cleaned_data.get('auto_levels'):
            levels.apply_levels(user)
        user.is_superuser = user.role == User.ROLE_ADMIN
        if request.POST.get('clear_profile_photo') and user.profile_photo:
            try:
                user.profile_photo.delete(save=False)
            except Exception:
                pass
            user.profile_photo = None
        user.save()
        if photo_cf is not None:
            if user.profile_photo:
                try:
                    user.profile_photo.delete(save=False)
                except Exception:
                    pass
            user.profile_photo.save(photo_cf.name, photo_cf, save=True)
        old, new = diff_snapshot(before, model_snapshot(user, USER_FIELDS))
        if new or photo_cf is not None or request.POST.get('clear_profile_photo'):
            log_action(request.user, 'update_user', target=user, old_value=old, new_value=new, request=request)
        messages.success(request, 'บันทึกข้อมูลเรียบร้อย')
        return redirect('composn:admin_user_detail', pk=user.pk)
    return render(request, 'admin/user_form.html', {
        'form': form, 'title': f'แก้ไขผู้ใช้ {user.camp1_code}', 'obj': user,
        'cancel_url': cancel_url, 'staff_perm_fields': STAFF_PERM_FIELDS,
        'allow_photo_upload': True,
    })


@staff_perm_required('staff_can_view_users')
def user_detail(request, pk):
    is_admin_viewer = request.user.is_admin_role
    user = get_object_or_404(User, pk=pk)
    if not is_admin_viewer and user.role != User.ROLE_MEMBER:
        raise PermissionDenied('เจ้าหน้าที่สามารถดูรายละเอียดได้เฉพาะนักเรียน/ผู้ใช้สิทธิ์เท่านั้น')
    if request.GET.get('export') == 'xlsx':
        if not is_admin_viewer:
            raise PermissionDenied('คุณไม่มีสิทธิ์ Export ข้อมูลผู้ใช้')
        log_action(request.user, 'export', target=user, new_value='ประวัติการใช้สิทธิ์รายบุคคล', request=request)
        return workbook_response(history_workbook(user), f'ประวัติใช้สิทธิ์_{user.camp1_code}.xlsx')
    from food.models import Attendance, absence_status, total_score
    total, breakdown = total_score(user)
    ctx = {
        'obj': user,
        'is_admin_viewer': is_admin_viewer,
        'summary': rights.member_rights_summary(user) if user.is_member_role else None,
        'school_requests': user.school_requests.all() if is_admin_viewer else [],
        'absence': absence_status(user),
        'attendances': Attendance.objects.filter(user=user)[:30],
        'total_score': total, 'score_breakdown': breakdown,
        'logs': AuditLog.objects.filter(Q(user=user) | Q(target_model='User', target_id=user.pk))[:30] if is_admin_viewer else [],
        'action_labels': ACTION_LABELS,
        'password_form': AdminSetPasswordForm() if is_admin_viewer else None,
    }
    return render(request, 'admin/user_detail.html', ctx)


@admin_required
@require_POST
def user_delete(request, pk):
    user = get_object_or_404(User, pk=pk)
    fallback_url = 'composn:admin_users' if user.role == User.ROLE_MEMBER else 'composn:admin_admins'
    if user.pk == request.user.pk:
        messages.error(request, 'ไม่สามารถลบบัญชีของตนเองได้')
        return _safe_redirect_next(request, fallback_url)
    _, _, _, user_approved_years, _ = _lock_helpers()
    locked_years = user_approved_years(user)
    if locked_years:
        yr_txt = ', '.join(str(y) for y in locked_years)
        messages.error(
            request,
            f'ไม่สามารถลบผู้ใช้ {user.camp1_code} ได้ เนื่องจากมีข้อมูลในปีการศึกษาที่อนุมัติประกาศผลแล้ว ({yr_txt})',
        )
        return _safe_redirect_next(request, fallback_url)
    snap = model_snapshot(user, ['camp1_code', 'first_name', 'last_name', 'role'])
    try:
        with transaction.atomic():
            uid = user.pk
            user.delete()
        log_action(request.user, 'delete_user', target_model='User', target_id=uid, old_value=snap, request=request)
        messages.success(request, f'ลบผู้ใช้ {snap["camp1_code"]} เรียบร้อย')
    except ProtectedError:
        user.is_active = False
        user.save(update_fields=['is_active', 'updated_at'])
        log_action(request.user, 'toggle_user', target=user, old_value={'is_active': True},
                   new_value={'is_active': False, 'reason': 'ลบไม่ได้เพราะมีประวัติ'}, request=request)
        messages.warning(request, f'ผู้ใช้ {snap["camp1_code"]} มีประวัติการใช้สิทธิ์/เวลาเรียน/คะแนน '
                                  f'จึงไม่สามารถลบได้ ระบบได้ปิดการใช้งานบัญชีแทน')
    return _safe_redirect_next(request, fallback_url)


@admin_required
@require_POST
def user_toggle(request, pk):
    user = get_object_or_404(User, pk=pk)
    _, _, _, user_approved_years, _ = _lock_helpers()
    locked_years = user_approved_years(user) if user.is_active else []
    if user.pk == request.user.pk:
        messages.error(request, 'ไม่สามารถปิดการใช้งานบัญชีของตนเองได้')
    elif locked_years:
        yr_txt = ', '.join(str(y) for y in locked_years)
        messages.error(
            request,
            f'ไม่สามารถปิดการใช้งาน {user.camp1_code} ได้ เนื่องจากมีข้อมูลในปีการศึกษาที่อนุมัติประกาศผลแล้ว ({yr_txt})',
        )
    else:
        user.is_active = not user.is_active
        if not user.is_active:
            user.current_session_key = ''
        user.save(update_fields=['is_active', 'current_session_key', 'updated_at'])
        log_action(request.user, 'toggle_user', target=user, old_value={'is_active': not user.is_active},
                   new_value={'is_active': user.is_active}, request=request)
        messages.success(request, f'{"เปิด" if user.is_active else "ปิด"}การใช้งาน {user.camp1_code} แล้ว')
    return _safe_redirect_next(request, 'composn:admin_users')


@admin_required
@require_POST
def user_reset_password(request, pk):
    """Reset รหัสผ่านกลับเป็นรหัสค่าย 1 (Admin/Staff = 1234)"""
    user = get_object_or_404(User, pk=pk)
    pw = user.camp1_code if user.role == User.ROLE_MEMBER else '1234'
    user.set_password(pw)
    user.is_first_login = True
    user.save()
    log_action(request.user, 'reset_password', target=user, new_value='reset เป็นรหัสเริ่มต้น', request=request)
    messages.success(request, f'Reset รหัสผ่านของ {user.camp1_code} เป็น "{pw}" แล้ว '
                              f'(ผู้ใช้ต้องเปลี่ยนรหัสผ่านเมื่อ Login ครั้งถัดไป)')
    return _safe_redirect_next(request, 'composn:admin_users')


@admin_required
@require_POST
def users_bulk_action(request):
    """จัดการผู้ใช้หลายรายการพร้อมกัน (เปิด/ปิดการใช้งาน, Reset รหัสผ่าน, กำหนดศูนย์, ลบ)"""
    ids = request.POST.getlist('user_ids')
    action = request.POST.get('bulk_action', '')
    qs = User.objects.filter(pk__in=ids).exclude(pk=request.user.pk)
    count = qs.count()
    if not count:
        messages.warning(request, 'กรุณาเลือกรายการผู้ใช้ก่อนดำเนินการ')
        return redirect('composn:admin_users')

    if action == 'activate':
        qs.update(is_active=True)
        log_action(request.user, 'toggle_user', target_model='User', new_value={'bulk_activate': count}, request=request)
        messages.success(request, f'เปิดการใช้งาน {count} บัญชีเรียบร้อย')
    elif action == 'deactivate':
        _, _, _, user_approved_years, _ = _lock_helpers()
        allowed_pks, locked_codes = [], []
        for u in qs:
            if user_approved_years(u):
                locked_codes.append(u.camp1_code)
            else:
                allowed_pks.append(u.pk)
        if allowed_pks:
            User.objects.filter(pk__in=allowed_pks).update(is_active=False, current_session_key='')
            log_action(request.user, 'toggle_user', target_model='User', new_value={'bulk_deactivate': len(allowed_pks)}, request=request)
            messages.success(request, f'ปิดการใช้งาน {len(allowed_pks)} บัญชีเรียบร้อย')
        if locked_codes:
            messages.error(request, f'ไม่สามารถปิดการใช้งานบัญชีในปีที่อนุมัติประกาศผลแล้วได้: {", ".join(locked_codes[:10])}')
    elif action == 'assign_center':
        center = request.POST.get('assign_center', '').upper()
        if center in ('DS', 'KM'):
            qs.update(camp_center=center)
            log_action(request.user, 'update_user', target_model='User', new_value={'bulk_center': center, 'count': count}, request=request)
            messages.success(request, f'กำหนดศูนย์ {center} ให้ {count} บัญชีเรียบร้อย')
    elif action == 'reset_password':
        for u in qs:
            pw = u.camp1_code if u.role == User.ROLE_MEMBER else '1234'
            u.set_password(pw)
            u.is_first_login = True
            u.save(update_fields=['password', 'is_first_login', 'updated_at'])
        log_action(request.user, 'reset_password', target_model='User', new_value={'bulk_reset': count}, request=request)
        messages.success(request, f'Reset รหัสผ่าน {count} บัญชีเรียบร้อย')
    elif action == 'delete':
        _, _, _, user_approved_years, _ = _lock_helpers()
        deleted, deactivated, locked_codes = 0, 0, []
        for u in qs:
            if user_approved_years(u):
                locked_codes.append(u.camp1_code)
                continue
            try:
                with transaction.atomic():
                    u.delete()
                deleted += 1
            except ProtectedError:
                u.is_active = False
                u.save(update_fields=['is_active', 'updated_at'])
                deactivated += 1
        if deleted or deactivated:
            log_action(request.user, 'delete_user', target_model='User',
                       new_value={'bulk_deleted': deleted, 'deactivated': deactivated}, request=request)
            msg = f'ลบสำเร็จ {deleted} บัญชี'
            if deactivated:
                msg += f' (ปิดการใช้งานแทน {deactivated} บัญชีที่มีประวัติ)'
            messages.success(request, msg)
        if locked_codes:
            messages.error(request, f'ไม่สามารถลบหรือปิดใช้งานบัญชีในปีที่อนุมัติประกาศผลแล้วได้: {", ".join(locked_codes[:10])}')
    return redirect('composn:admin_users')


@admin_required
@require_POST
def user_set_password(request, pk):
    user = get_object_or_404(User, pk=pk)
    form = AdminSetPasswordForm(request.POST)
    if form.is_valid():
        user.set_password(form.cleaned_data['new_password'])
        user.is_first_login = form.cleaned_data['force_change']
        user.save()
        log_action(request.user, 'set_password', target=user, request=request)
        messages.success(request, f'เปลี่ยนรหัสผ่านให้ {user.camp1_code} เรียบร้อย')
    else:
        messages.error(request, 'รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร')
    return redirect('composn:admin_user_detail', pk=pk)


# =============================================================================
# Import / คำนวณระดับชั้น
# =============================================================================
@admin_required
def import_students_view(request):
    form = ExcelUploadForm(request.POST or None, request.FILES or None)
    result = None
    preview = None
    if request.method == 'POST':
        action = request.POST.get('action', '')
        # กรณี 1: กดยืนยันนำเข้าจากหน้า Preview
        if action == 'confirm_import' and request.session.get('import_students_preview'):
            pdata = request.session.pop('import_students_preview')
            result = import_students_from_preview(pdata)
            log_action(request.user, 'import_students', target_model='User', new_value={
                'file': pdata.get('filename', ''), 'created': result['created'], 'updated': result['updated'],
                'errors': len(result['errors'])}, request=request)
            messages.success(request, f'Import สำเร็จ: เพิ่มใหม่ {result["created"]} คน, ทับข้อมูลเดิม {result["updated"]} คน')
        elif form.is_valid():
            uploaded = form.cleaned_data['file']
            if action == 'preview':
                try:
                    preview = preview_students(uploaded)
                    request.session['import_students_preview'] = {
                        'filename': uploaded.name,
                        'rows': preview['rows'],
                        'present_fields': preview.get('present_fields', []),
                    }
                except ValueError as e:
                    messages.error(request, str(e))
            else:
                try:
                    result = import_students(uploaded)
                except ValueError as e:
                    messages.error(request, str(e))
                else:
                    log_action(request.user, 'import_students', target_model='User', new_value={
                        'file': uploaded.name, 'created': result['created'],
                        'updated': result['updated'], 'errors': len(result['errors'])}, request=request)
                    messages.success(request, f'Import สำเร็จ: เพิ่มใหม่ {result["created"]} คน, '
                                              f'ทับข้อมูลเดิม {result["updated"]} คน')
    return render(request, 'admin/import.html', {
        'form': form, 'result': result, 'preview': preview,
        'title': 'Import ข้อมูลนักเรียน', 'columns': STUDENT_COLUMNS_TH, 'columns_en': STUDENT_COLUMNS,
        'template_url': 'composn:admin_template_students', 'supports_preview': True,
        'notes': [
            'Template รวมทุกฟิลด์รายละเอียดผู้ใช้ที่จัดเก็บในระบบ ยกเว้น รหัสค่าย 2 (14 คอลัมน์)',
            'คอลัมน์ในไฟล์ที่นำเข้าสามารถเรียงลำดับแบบใดก็ได้ ระบบจะจับคู่ข้อมูลลงฐานข้อมูลตามชื่อหัวคอลัมน์ให้อัตโนมัติ (และข้ามคอลัมน์ที่ไม่เกี่ยวข้องหรือรหัสค่าย 2)',
            'รองรับทั้งคอลัมน์ "ชื่อ-นามสกุล" ช่องเดียว และแบบแยก "ชื่อ" + "นามสกุล" รวมถึงไฟล์จากระบบเดิมที่ไม่มีแถวหัวตาราง',
            'หากรหัสค่าย 1 ซ้ำกับผู้ใช้เดิม ระบบจะอัปเดตเฉพาะคอลัมน์ที่มีอยู่ในไฟล์ (คอลัมน์ที่ไม่ได้ใส่มาจะคงค่าเดิมไว้) · หากเป็นรหัสใหม่จะสร้างผู้ใช้ใหม่ (รหัสผ่านเริ่มต้น = รหัสค่าย 1)',
            'ระบบคำนวณระดับชั้น +1 ถึง +5 ให้อัตโนมัติ และรองรับวันเกิดรูปแบบ วว/ดด/ปปปป (พ.ศ./ค.ศ.) หรือวันที่ของ Excel',
        ],
    })


@admin_required
def import_camp2_view(request):
    form = ExcelUploadForm(request.POST or None, request.FILES or None)
    result = None
    if request.method == 'POST' and form.is_valid():
        try:
            result = import_camp2(form.cleaned_data['file'])
        except ValueError as e:
            messages.error(request, str(e))
        else:
            log_action(request.user, 'import_camp2', target_model='User', new_value={
                'file': form.cleaned_data['file'].name, 'updated': result['updated'],
                'errors': len(result['errors'])}, request=request)
            messages.success(request, f'อัปเดตรหัสค่าย 2 สำเร็จ {result["updated"]} คน')
    return render(request, 'admin/import.html', {
        'form': form, 'result': result, 'title': 'Import รหัสค่าย 2',
        'columns': CAMP2_COLUMNS_TH, 'columns_en': CAMP2_COLUMNS,
        'template_url': 'composn:admin_template_camp2',
        'notes': ['รองรับหัวตารางทั้งภาษาไทย (รหัสค่าย 1, รหัสค่าย 2) และภาษาอังกฤษ (camp1_code, camp2_code)',
                  'ระบบค้นหาผู้ใช้จาก รหัสค่าย 1 แล้วอัปเดต รหัสค่าย 2',
                  'ผู้ที่มีรหัสค่าย 2 จะใช้สิทธิ์ค่าย 2 ได้ และไม่สามารถใช้สิทธิ์ค่าย 1 ได้อีก'],
    })


@admin_required
def delete_camp2_view(request):
    """ลบรหัสค่าย 2 โดยการพิมพ์รหัสค่าย 2 ทั้งหมด และ Preview ก่อนการลบ พร้อมยืนยันด้วยรหัสผ่านของ Admin"""
    ScoreApproval, approved_years_set, camp2_code_year, user_approved_years, _ = _lock_helpers()
    approved_set = approved_years_set()
    preview = None
    not_found = []
    codes_text = ''
    matched_users = []
    locked_codes = []

    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'preview':
            codes_text = request.POST.get('camp2_codes_text', '').strip()
            if not codes_text:
                messages.error(request, 'กรุณากรอกรหัสค่าย 2 ที่ต้องการลบ')
            else:
                raw_codes = [c.strip().upper() for c in re.split(r'[\r\n,\s]+', codes_text) if c.strip()]
                seen = set()
                codes = []
                for c in raw_codes:
                    if c not in seen:
                        seen.add(c)
                        codes.append(c)

                users_qs = list(User.objects.filter(camp2_code__in=codes))
                matched_map = {u.camp2_code.upper(): u for u in users_qs if u.camp2_code}
                matched_users = [matched_map[c] for c in codes if c in matched_map]
                not_found = [c for c in codes if c not in matched_map]
                locked_codes = [
                    u.camp2_code for u in matched_users
                    if (camp2_code_year(u.camp2_code) in approved_set) or user_approved_years(u)
                ]
                if locked_codes:
                    messages.warning(
                        request,
                        f'พบรหัสค่าย 2 ที่อยู่ในปีการศึกษาที่อนุมัติประกาศผลแล้ว ({", ".join(locked_codes[:10])}) '
                        f'ไม่สามารถลบได้ กรุณายกเลิกประกาศผลก่อน',
                    )
                preview = {
                    'total_input': len(codes),
                    'matched_count': len(matched_users),
                    'not_found_count': len(not_found),
                    'codes_str': ','.join(matched_map.keys()),
                }
        elif action == 'confirm_delete':
            admin_password = request.POST.get('admin_password', '')
            if not request.user.check_password(admin_password):
                messages.error(request, 'รหัสผ่านของผู้ดูแลระบบไม่ถูกต้อง ไม่สามารถลบรหัสค่าย 2 ได้')
            else:
                confirmed_codes = [c.strip().upper() for c in request.POST.get('confirmed_codes', '').split(',') if c.strip()]
                if not confirmed_codes:
                    messages.error(request, 'ไม่พบรหัสค่าย 2 ที่ต้องการลบ')
                    return redirect('composn:admin_delete_camp2')
                users_to_clear = list(User.objects.filter(camp2_code__in=confirmed_codes))
                blocked = [
                    c for c in confirmed_codes if camp2_code_year(c) in approved_set
                ] + [
                    u.camp2_code for u in users_to_clear if user_approved_years(u) and u.camp2_code
                ]
                if blocked:
                    uniq_blocked = list(dict.fromkeys(blocked))
                    messages.error(
                        request,
                        f'ไม่สามารถลบรหัสค่าย 2 ของปีการศึกษาที่อนุมัติประกาศผลแล้วได้ ({", ".join(uniq_blocked[:10])}) '
                        f'กรุณายกเลิกประกาศผลก่อน',
                    )
                    return redirect('composn:admin_delete_camp2')
                count = 0
                with transaction.atomic():
                    for u in users_to_clear:
                        u.camp2_code = None
                        u.save(update_fields=['camp2_code', 'updated_at'])
                        count += 1
                log_action(request.user, 'delete_camp2', target_model='User',
                           new_value={'cleared_count': count, 'codes': confirmed_codes}, request=request)
                messages.success(request, f'ลบรหัสค่าย 2 เรียบร้อยแล้ว {count} คน')
                return redirect('composn:admin_users')

    return render(request, 'admin/delete_camp2.html', {
        'preview': preview,
        'matched_users': matched_users,
        'not_found': not_found,
        'codes_text': codes_text,
        'locked_codes': locked_codes,
    })


@admin_required
def template_students(request):
    lang = 'en' if request.GET.get('lang') == 'en' else 'th'
    wb = students_template_workbook(lang=lang)
    fname = 'template_students_en.xlsx' if lang == 'en' else 'ตัวอย่างไฟล์นำเข้านักเรียน.xlsx'
    return workbook_response(wb, fname)


@admin_required
def template_camp2(request):
    lang = 'en' if request.GET.get('lang') == 'en' else 'th'
    wb = camp2_template_workbook(lang=lang)
    fname = 'template_camp2_en.xlsx' if lang == 'en' else 'ตัวอย่างไฟล์นำเข้ารหัสค่าย2.xlsx'
    return workbook_response(wb, fname)


@admin_required
def template_rights(request):
    lang = 'en' if request.GET.get('lang') == 'en' else 'th'
    wb = rights_template_workbook(lang=lang)
    fname = 'template_rights_en.xlsx' if lang == 'en' else 'ตัวอย่างไฟล์นำเข้าผู้ใช้สิทธิ์ค่าย1.xlsx'
    return workbook_response(wb, fname)


@admin_required
def import_rights_users_view(request):
    """Admin นำเข้าผู้ใช้สิทธิ์ในปีการศึกษานั้น (ตรวจสอบตามค่ายที่เปิดให้ใช้สิทธิ์อยู่ ณ เวลานั้น และแสดง Preview ก่อนนำเข้า)"""
    ScoreApproval, _, _, _, year_locked_message = _lock_helpers()
    setting = SystemSetting.get()
    year = int(setting.current_academic_year)
    is_year_approved = ScoreApproval.is_year_approved(year)
    active_camp, active_rr = rights.get_active_camp_for_year(year)
    form = ExcelUploadForm(request.POST or None, request.FILES or None)
    preview = None
    result = None
    raw_codes = request.POST.get('raw_codes', '').strip() if request.method == 'POST' else ''

    if request.method == 'POST':
        if is_year_approved:
            messages.error(request, year_locked_message(year))
            return redirect('composn:admin_import_rights')
        action = request.POST.get('action', 'preview')
        if action == 'confirm_import':
            valid_codes = request.POST.getlist('valid_codes')
            if not valid_codes and request.session.get('import_rights_preview'):
                pdata = request.session.pop('import_rights_preview', {})
                valid_codes = pdata.get('codes', [])
            else:
                request.session.pop('import_rights_preview', None)
            if valid_codes:
                try:
                    result = rights.import_camp1_rights(valid_codes, request.user, academic_year=year, request=request)
                except rights.RightsError as e:
                    messages.error(request, str(e))
                else:
                    messages.success(
                        request,
                        f'นำเข้าผู้ใช้สิทธิ์ (ค่ายที่ {result["camp_level"]}) ปีการศึกษา {year} สำเร็จ {result["imported"]} คน',
                    )
            else:
                messages.warning(request, 'ไม่มีรายการรหัสค่าย 1 ที่ผ่านการตรวจสอบให้นำเข้า')
        else:
            code_items = []
            uploaded = request.FILES.get('file')
            try:
                if uploaded:
                    code_items = read_rights_codes(uploaded)
                elif raw_codes:
                    tokens = [t.strip() for t in re.split(r'[\r\n,;\t\s]+', raw_codes) if t.strip()]
                    code_items = [(idx, t) for idx, t in enumerate(tokens, start=1)]
                else:
                    raise ValueError('กรุณาเลือกไฟล์ Excel/CSV หรือวางรหัสค่าย 1 ที่ต้องการตรวจสอบ')
                preview = rights.preview_camp1_rights_import(code_items, academic_year=year)
                request.session['import_rights_preview'] = {
                    'codes': [r['camp1_code'] for r in preview['rows'] if r['valid']],
                    'academic_year': year,
                }
            except ValueError as e:
                messages.error(request, str(e))

    return render(request, 'admin/import_rights.html', {
        'form': form,
        'preview': preview,
        'result': result,
        'raw_codes': raw_codes,
        'year': year,
        'active_camp': active_camp,
        'active_rr': active_rr,
        'setting': setting,
        'is_year_approved': is_year_approved,
    })


@admin_required
def calculate_level(request):
    if request.method == 'POST':
        qs = User.objects.filter(role=User.ROLE_MEMBER)
        if request.POST.get('only_empty'):
            qs = qs.filter(year_plus_1='')
        ok, unknown = 0, []
        with transaction.atomic():
            for u in qs:
                if levels.apply_levels(u):
                    ok += 1
                else:
                    unknown.append(u.camp1_code)
                u.save(update_fields=['current_level', 'year_plus_1', 'year_plus_2', 'year_plus_3',
                                      'year_plus_4', 'year_plus_5', 'updated_at'])
        log_action(request.user, 'calculate_level', target_model='User',
                   new_value={'calculated': ok, 'unknown': len(unknown)}, request=request)
        messages.success(request, f'คำนวณระดับชั้นเรียบร้อย {ok} คน')
        if unknown:
            messages.warning(request, f'ไม่สามารถคำนวณได้ {len(unknown)} คน (ระดับชั้นปัจจุบันไม่ถูกต้อง/ว่าง): '
                                      + ', '.join(unknown[:30]) + (' ...' if len(unknown) > 30 else ''))
        return redirect('composn:admin_calculate_level')
    sample = User.objects.filter(role=User.ROLE_MEMBER).order_by('camp1_code')[:20]
    return render(request, 'admin/calculate_level.html', {'sample': sample})


# =============================================================================
# รอบใช้สิทธิ์ & การใช้สิทธิ์
# =============================================================================
@admin_required
def rights_settings(request):
    """ตั้งค่าการใช้สิทธิ์ (รวมกับเปิด-ปิดการใช้สิทธิ์ โดยนำจำนวนครั้งสูงสุดออก)"""
    setting = SystemSetting.get()
    if request.method == 'POST':
        setting.rights_warning_text = request.POST.get('rights_warning_text', '').strip()
        setting.exam_date_camp1 = request.POST.get('exam_date_camp1', '').strip()
        setting.exam_date_camp2 = request.POST.get('exam_date_camp2', '').strip()
        setting.display_title = request.POST.get('display_title', setting.display_title).strip()
        setting.save()
        log_action(request.user, 'update_settings', target=setting,
                   new_value='ตั้งค่าการใช้สิทธิ์', request=request)
        messages.success(request, 'บันทึกการตั้งค่าการใช้สิทธิ์เรียบร้อย')
        return redirect('composn:admin_rounds')
    return rounds(request)


@admin_required
def rounds(request):
    setting = SystemSetting.get()
    if request.method == 'POST' and request.POST.get('action') == 'save_rights_settings':
        setting.rights_warning_text = request.POST.get('rights_warning_text', '').strip()
        setting.exam_date_camp1 = request.POST.get('exam_date_camp1', '').strip()
        setting.exam_date_camp2 = request.POST.get('exam_date_camp2', '').strip()
        setting.display_title = request.POST.get('display_title', setting.display_title).strip()
        setting.save()
        log_action(request.user, 'update_settings', target=setting,
                   new_value='ตั้งค่าการใช้สิทธิ์', request=request)
        messages.success(request, 'บันทึกการตั้งค่าการใช้สิทธิ์เรียบร้อย')
        return redirect('composn:admin_rounds')
    year = setting.current_academic_year
    auto_exam_camp1 = setting.get_auto_exam_date_camp1(year)
    if not setting.exam_date_camp1 and auto_exam_camp1:
        setting.exam_date_camp1 = auto_exam_camp1
        setting.save(update_fields=['exam_date_camp1', 'updated_at'])
    form = OpenRoundForm(initial={'academic_year': year, 'camp_level': 1})
    qs = RightsRound.objects.filter(academic_year=year).annotate(
        confirmed=Count('use_records', filter=Q(use_records__is_used=True, use_records__is_deleted=False)),
        waived=Count('use_records', filter=Q(use_records__is_used=False, use_records__is_deleted=False)),
    ).order_by('-opened_at')
    ScoreApproval, approved_years_set, _, _, _ = _lock_helpers()
    return render(request, 'admin/rounds.html', {
        'form': form, 'page': _paginate(request, qs, 30), 'setting': setting,
        'auto_exam_camp1': auto_exam_camp1,
        'is_year_approved': ScoreApproval.is_year_approved(year),
        'approved_years': approved_years_set(),
    })


@admin_required
@require_POST
def open_round(request):
    setting = SystemSetting.get()
    post_data = request.POST.copy()
    post_data['academic_year'] = str(setting.current_academic_year)
    form = OpenRoundForm(post_data)
    if form.is_valid():
        try:
            rr = rights.open_round(
                form.cleaned_data['academic_year'], form.cleaned_data['camp_level'],
                request.user, form.cleaned_data['title'], request=request,
                scheduled_open_at=form.cleaned_data.get('scheduled_open_at'),
                scheduled_close_at=form.cleaned_data.get('scheduled_close_at'),
            )
        except rights.RightsError as e:
            messages.error(request, str(e))
        else:
            messages.success(request, f'เปิดรอบใช้สิทธิ์ค่าย {rr.camp_level} ปีการศึกษา {rr.academic_year} แล้ว')
            return redirect('composn:admin_round_detail', pk=rr.pk)
    else:
        messages.error(request, 'ข้อมูลไม่ถูกต้อง')
    return redirect('composn:admin_rounds')


@admin_required
@require_POST
def close_round(request, pk):
    rr = get_object_or_404(RightsRound, pk=pk)
    try:
        waived, path = rights.close_round(rr, request.user, mark_waived=bool(request.POST.get('mark_waived')),
                                          request=request)
    except rights.RightsError as e:
        messages.error(request, str(e))
    else:
        messages.success(request, f'ปิดรอบเรียบร้อย บันทึกสละสิทธิ์ {waived} คน และสำรองข้อมูลไว้ที่ {path.name}')
    return _safe_redirect_next(request, 'composn:admin_round_detail', pk=pk)


@admin_required
@require_POST
def reopen_round(request, pk):
    """เปิดรอบที่เคยปิดไปแล้วขึ้นมาใหม่"""
    rr = get_object_or_404(RightsRound, pk=pk)
    try:
        rights.reopen_round(rr, request.user, request=request)
    except rights.RightsError as e:
        messages.error(request, str(e))
    else:
        messages.success(request, f'เปิดรอบใช้สิทธิ์ค่าย {rr.camp_level} ปีการศึกษา {rr.academic_year} ใหม่อีกครั้งแล้ว')
    return _safe_redirect_next(request, 'composn:admin_round_detail', pk=pk)


def _parse_dt_local(val):
    if not val:
        return None
    try:
        naive = dt.datetime.fromisoformat(val.strip())
        return timezone.make_aware(naive) if timezone.is_naive(naive) else naive
    except ValueError:
        return None


@admin_required
@require_POST
def round_edit_schedule(request, pk):
    """แก้ไขชื่อรอบ และวันเวลาเปิด-ปิดการใช้สิทธิ์"""
    rr = get_object_or_404(RightsRound, pk=pk)
    rr.title = request.POST.get('title', rr.title).strip()
    rr.scheduled_open_at = _parse_dt_local(request.POST.get('scheduled_open_at'))
    rr.scheduled_close_at = _parse_dt_local(request.POST.get('scheduled_close_at'))
    if request.POST.get('opened_at'):
        dt_open = _parse_dt_local(request.POST.get('opened_at'))
        if dt_open:
            rr.opened_at = dt_open
    if request.POST.get('closed_at'):
        dt_close = _parse_dt_local(request.POST.get('closed_at'))
        if dt_close:
            rr.closed_at = dt_close
    rr.save()
    log_action(request.user, 'update_settings', target=rr,
               new_value={'scheduled_open_at': str(rr.scheduled_open_at),
                          'scheduled_close_at': str(rr.scheduled_close_at)}, request=request)
    messages.success(request, 'บันทึกการแก้ไขวันเวลาเปิด-ปิดรอบเรียบร้อย')
    return _safe_redirect_next(request, 'composn:admin_round_detail', pk=pk)


@admin_required
def use_records_list(request):
    """หน้าแสดงรายชื่อผู้ใช้สิทธิ์ (รวมหน้าแสดงผลไว้ในหน้าเดียว แสดงตามปีการศึกษาที่ตั้งค่าไว้)"""
    setting = SystemSetting.get()
    year_req = request.GET.get('year', '').strip()
    try:
        year_int = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year_int = int(setting.current_academic_year)
    year = str(year_int)
    camp = request.GET.get('camp', '')
    center = request.GET.get('center', '')
    status = request.GET.get('status', 'used')
    q = request.GET.get('q', '').strip()

    base_qs = UseRecord.objects.filter(academic_year=year_int)
    if center in ('DS', 'KM'):
        base_qs = base_qs.filter(center_snapshot=center)
    stats = base_qs.aggregate(
        total=Count('id', filter=Q(is_used=True)),
        camp1=Count('id', filter=Q(is_used=True, camp_level=1)),
        camp2=Count('id', filter=Q(is_used=True, camp_level=2)),
        waived=Count('id', filter=Q(is_used=False)),
    )
    yy = f'{year_int % 100:02d}'
    members_qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=year_int)
    )
    if center in ('DS', 'KM'):
        members_qs = members_qs.filter(camp_center=center)
    open_rounds = list(RightsRound.objects.filter(is_open=True, academic_year=year_int))

    qs = UseRecord.objects.select_related('user', 'rights_round').filter(academic_year=year_int)
    if camp in ('1', '2'):
        qs = qs.filter(camp_level=int(camp))
    if center in ('DS', 'KM'):
        qs = qs.filter(center_snapshot=center)
    if status == 'used':
        qs = qs.filter(is_used=True)
    elif status == 'waived':
        qs = qs.filter(is_used=False)
    if q:
        qs = qs.filter(Q(rights_code__icontains=q) | Q(user__camp1_code__icontains=q) |
                       Q(user__camp2_code__icontains=q) | Q(full_name_snapshot__icontains=q) |
                       Q(school_snapshot__icontains=q))
    qs = qs.order_by('-academic_year', 'camp_level', '-is_used', 'rights_code')

    if request.GET.get('export') == 'xlsx':
        wb = build_workbook([(f'ใช้สิทธิ์_{year or "ทั้งหมด"}', USE_RECORD_HEADERS, use_record_rows(qs))])
        return workbook_response(wb, f'รายชื่อผู้ใช้สิทธิ์_{year or "ทั้งหมด"}.xlsx')

    return render(request, 'admin/use_records.html', {
        'page': _paginate(request, qs, 50),
        'year': year, 'camp': camp, 'center': center, 'status': status, 'q': q,
        'stats': stats, 'member_total': members_qs.count(), 'open_rounds': open_rounds,
        'setting': setting,
    })


@admin_required
@require_POST
def archive_record_delete(request, pk):
    """ลบประวัติการใช้สิทธิ์แต่ละครั้ง (ARCHIVE_RECORD) และ USE_RECORD ที่ผูกกัน"""
    arc = get_object_or_404(ArchiveRecord, pk=pk)
    uid = arc.user_id
    ScoreApproval, _, _, _, year_locked_message = _lock_helpers()
    if ScoreApproval.is_year_approved(arc.academic_year):
        messages.error(request, year_locked_message(arc.academic_year))
        return _safe_redirect_next(request, 'composn:admin_user_detail', pk=uid)
    with transaction.atomic():
        if arc.use_record_id:
            UseRecord.objects.filter(pk=arc.use_record_id).delete(user=request.user)
        arc.delete(user=request.user)
    log_action(request.user, 'delete_use_record', target=arc,
               old_value={'rights_code': arc.rights_code, 'year': arc.academic_year}, request=request)
    messages.success(request, f'ลบประวัติการใช้สิทธิ์ {arc.rights_code} (ปี {arc.academic_year}) เรียบร้อย')
    return _safe_redirect_next(request, 'composn:admin_user_detail', pk=uid)


@admin_required
def round_detail(request, pk):
    rr = get_object_or_404(RightsRound, pk=pk)
    qs = rr.use_records.select_related('user').order_by('-is_used', 'rights_code', 'user__camp1_code')
    if request.GET.get('export') == 'xlsx':
        log_action(request.user, 'export', target=rr, new_value='รายงานรอบใช้สิทธิ์', request=request)
        wb = build_workbook([(f'ค่าย{rr.camp_level}_{rr.academic_year}', USE_RECORD_HEADERS, use_record_rows(qs))])
        return workbook_response(wb, f'รายงานใช้สิทธิ์ค่าย{rr.camp_level}_ปีการศึกษา{rr.academic_year}.xlsx')
    pending = rights.eligible_unconfirmed_users(rr) if rr.is_open else []
    return render(request, 'admin/round_detail.html', {'rr': rr, 'records': qs, 'pending': pending})


@admin_required
def year_report(request):
    """ดาวน์โหลดรายงาน Excel ทั้งปีการศึกษา"""
    try:
        year = int(request.GET.get('year') or SystemSetting.get().current_academic_year)
    except (TypeError, ValueError):
        year = SystemSetting.get().current_academic_year
    camp_raw = request.GET.get('camp')
    if camp_raw in ('1', '2'):
        camp = int(camp_raw)
        qs = use_records_queryset(year, camp)
        sheet_title = f'ค่าย{camp}_{year}'
        fname = f'ใช้สิทธิ์ค่าย{camp}_ปีการศึกษา{year}.xlsx'
    else:
        qs = UseRecord.objects.filter(academic_year=year).select_related('user').order_by(
            'camp_level', '-is_used', 'rights_code', 'user__camp1_code'
        )
        sheet_title = f'ใช้สิทธิ์_{year}'
        fname = f'ใช้สิทธิ์_ปีการศึกษา{year}.xlsx'
    log_action(request.user, 'export', target_model='UseRecord', new_value={'year': year}, request=request)
    wb = build_workbook([(sheet_title, USE_RECORD_HEADERS, use_record_rows(qs))])
    return workbook_response(wb, fname)


@admin_required
@require_POST
def use_record_delete(request, pk):
    rec = get_object_or_404(UseRecord, pk=pk)
    ScoreApproval, _, _, _, year_locked_message = _lock_helpers()
    if ScoreApproval.is_year_approved(rec.academic_year):
        messages.error(request, year_locked_message(rec.academic_year))
        return _safe_redirect_next(request, 'composn:admin_rounds')
    with transaction.atomic():
        ArchiveRecord.objects.filter(use_record=rec).delete(user=request.user)
        rec.delete(user=request.user)
    log_action(request.user, 'delete_use_record', target=rec, old_value={
        'rights_code': rec.rights_code, 'user': rec.user.camp1_code, 'year': rec.academic_year}, request=request)
    messages.success(request, f'ลบ (Soft Delete) บันทึก {rec.rights_code or "สละสิทธิ์"} ของ '
                              f'{rec.user.camp1_code} แล้ว')
    return _safe_redirect_next(request, 'composn:admin_rounds')


# =============================================================================
# คำขอแก้ไขชื่อโรงเรียน
# =============================================================================
@admin_required
def school_requests(request):
    if request.method == 'POST':
        req = get_object_or_404(SchoolChangeRequest, pk=request.POST.get('id'))
        action = request.POST.get('action')
        note = request.POST.get('admin_note', '').strip()
        if action in ('approve', 'reject') and req.status == SchoolChangeRequest.STATUS_PENDING:
            req.reviewed_by = request.user
            req.reviewed_at = timezone.now()
            req.admin_note = note
            if action == 'approve':
                user = req.user
                old = user.school_name
                user.school_name = req.new_school_name
                user.save(update_fields=['school_name', 'updated_at'])
                req.status = SchoolChangeRequest.STATUS_APPROVED
                log_action(request.user, 'school_approve', target=user, old_value={'school_name': old},
                           new_value={'school_name': user.school_name}, request=request)
                messages.success(request, f'อนุมัติชื่อโรงเรียนของ {user.camp1_code} แล้ว')
            else:
                req.status = SchoolChangeRequest.STATUS_REJECTED
                log_action(request.user, 'school_reject', target=req, new_value=note, request=request)
                messages.success(request, 'บันทึกไม่อนุมัติแล้ว')
            req.save()
        elif action == 'clear':
            req.is_cleared = True
            req.save(update_fields=['is_cleared'])
            log_action(request.user, 'school_clear', target=req, request=request)
            messages.success(request, 'เคลียร์สถานะแล้ว ผู้ใช้สามารถส่งคำขอใหม่ได้')
        return redirect(request.get_full_path())
    status = request.GET.get('status', 'open')
    qs = SchoolChangeRequest.objects.select_related('user', 'reviewed_by')
    if status == 'open':
        qs = qs.filter(is_cleared=False)
    elif status in ('pending', 'approved', 'rejected'):
        qs = qs.filter(status=status)
    return render(request, 'admin/school_requests.html', {'page': _paginate(request, qs), 'status': status})


# =============================================================================
# ตั้งค่าระบบ
# =============================================================================
def handle_add_academic_year_post(request, setting):
    """ตรวจสอบและเพิ่มปีการศึกษาใหม่เข้าระบบ ใช้ร่วมกันระหว่างหน้าตั้งค่าปีการศึกษาและ API เดิม"""
    raw_year = (request.POST.get('new_academic_year') or '').strip()
    set_as_current = request.POST.get('set_as_current') in ('1', 'true', 'on', 'yes')
    try:
        new_year = int(raw_year)
    except (TypeError, ValueError):
        messages.error(request, 'กรุณาระบุปีการศึกษา (พ.ศ.) เป็นตัวเลข 4 หลัก')
        return False
    if not (MIN_ACADEMIC_YEAR <= new_year <= MAX_ACADEMIC_YEAR):
        messages.error(request, f'ปีการศึกษาต้องอยู่ระหว่าง พ.ศ. {MIN_ACADEMIC_YEAR} ถึง {MAX_ACADEMIC_YEAR}')
        return False
    existing_years = set(get_academic_year_choices(setting.current_academic_year, setting=setting))
    if new_year in existing_years:
        messages.warning(request, f'ปีการศึกษา {new_year} มีอยู่ในระบบแล้ว กรุณาเลือกจากรายการด้านล่าง')
        return False
    before_year = setting.current_academic_year
    setting.add_academic_year(new_year)
    if set_as_current:
        setting.current_academic_year = new_year
    setting.save()
    log_action(
        request.user, 'update_settings', target=setting,
        old_value={'current_academic_year': before_year},
        new_value={
            'add_academic_year': new_year,
            'current_academic_year': setting.current_academic_year,
            'set_as_current': set_as_current,
        },
        request=request,
    )
    if set_as_current:
        messages.success(
            request,
            f'เพิ่มปีการศึกษา {new_year} และสลับเป็นปีการศึกษาปัจจุบันเรียบร้อยแล้ว '
            f'สามารถเริ่มนำเข้านักเรียน ตั้งวันอบรม และกำหนดคะแนนของปี {new_year} ได้ทันที',
        )
    else:
        messages.success(
            request,
            f'เพิ่มปีการศึกษา {new_year} เข้าระบบเรียบร้อยแล้ว สามารถเลือกใช้งานจากตัวเลือกปีการศึกษาได้ทันที',
        )
    return True


@admin_required
def system_settings(request):
    from food.models import AbsenceReason, WarningColor
    ScoreApproval, _, _, _, _ = _lock_helpers()
    setting = SystemSetting.get()
    old_year = setting.current_academic_year
    old_max_abs = setting.max_absences
    old_max_qual = setting.max_qualified_count
    current_year_approved = ScoreApproval.is_year_approved(old_year)
    if request.method == 'POST' and request.POST.get('action') == 'add_academic_year':
        handle_add_academic_year_post(request, setting)
        return redirect('composn:admin_settings')
    if request.method == 'POST' and request.POST.get('action') == 'make_logos_transparent':
        converted = convert_existing_setting_logos_to_transparent(setting)
        log_action(request.user, 'update_settings', target=setting,
                   new_value={'make_logos_transparent': converted}, request=request)
        if converted:
            messages.success(request, f'แปลงพื้นหลังรูปตราสัญลักษณ์ให้โปร่งใสเรียบร้อยแล้ว ({len(converted)} รูป)')
        else:
            messages.success(request, 'รูปตราสัญลักษณ์ทั้งหมดมีพื้นหลังโปร่งใสอยู่แล้ว')
        return redirect('composn:admin_settings')
    if request.method == 'POST' and request.POST.get('action') == 'save_all_settings':
        # 1. บันทึกตั้งค่าหลัก (SystemSettingForm + fallback ฟิลด์พื้นฐาน)
        form = SystemSettingForm(request.POST, instance=setting)
        if form.is_valid():
            setting = form.save(commit=False)
        else:
            if 'site_title' in request.POST and request.POST.get('site_title', '').strip():
                setting.site_title = request.POST.get('site_title').strip()
            if 'site_title_line2' in request.POST:
                setting.site_title_line2 = request.POST.get('site_title_line2', '').strip()
            if 'display_title' in request.POST and request.POST.get('display_title', '').strip():
                setting.display_title = request.POST.get('display_title').strip()
            try:
                if 'auto_logout_minutes' in request.POST and request.POST.get('auto_logout_minutes', '').strip() != '':
                    setting.auto_logout_minutes = max(0, min(1440, int(request.POST.get('auto_logout_minutes'))))
            except ValueError:
                pass
            if 'editable_fields' in request.POST:
                setting.set_editable_fields(request.POST.getlist('editable_fields'))
        # รองรับพารามิเตอร์ปีการศึกษา/สิทธิ์/จำนวนขาดสูงสุด หากมีการส่งมาโดยตรงเพื่อความเข้ากันได้ย้อนหลัง
        try:
            if request.POST.get('current_academic_year'):
                req_year = int(request.POST.get('current_academic_year'))
                allowed_years = set(get_academic_year_choices(setting.current_academic_year, setting=setting))
                if req_year in allowed_years:
                    setting.current_academic_year = req_year
                else:
                    messages.warning(
                        request,
                        f'ปีการศึกษา {req_year} ยังไม่มีในระบบ กรุณากดปุ่ม "เพิ่มปีการศึกษา" ก่อนเลือกใช้งาน',
                    )
        except ValueError:
            pass
        try:
            if request.POST.get('max_rights'):
                setting.max_rights = max(1, int(request.POST.get('max_rights')))
        except ValueError:
            pass
        try:
            if request.POST.get('max_absences'):
                setting.max_absences = max(1, int(request.POST.get('max_absences')))
        except ValueError:
            pass
        # 2. บันทึกสีและหน้า Login (สีแถบข้าง Staff / Member, ตราสัญลักษณ์, หน้าเข้าสู่ระบบ, เอกสารประกาศผล)
        extra_text_fields = [
            'color_monday', 'color_tuesday', 'color_wednesday', 'color_thursday',
            'color_friday', 'color_saturday', 'color_sunday',
            'member_theme_primary', 'member_theme_gradient',
            'staff_theme_primary', 'staff_theme_gradient',
            'login_subtitle', 'login_bg_color1', 'login_bg_color2', 'login_bg_color3',
            'login_btn_color1', 'login_btn_color2',
            'announcement_school_name', 'announcement_center_name',
            'announcement_director_name', 'announcement_director_position',
            'announcement_date_text', 'announcement_title', 'announcement_note',
        ]
        for f in extra_text_fields:
            if f in request.POST:
                setattr(setting, f, request.POST.get(f, '').strip())
        if 'max_qualified_count' in request.POST:
            try:
                setting.max_qualified_count = max(0, int(request.POST.get('max_qualified_count')))
            except (ValueError, TypeError):
                pass
        if 'editable_fields' in request.POST:
            setting.set_editable_fields(request.POST.getlist('editable_fields'))
        for img_field, clear_key in (
            ('login_logo', 'clear_login_logo'),
            ('logo_school', 'clear_logo_school'),
            ('logo_posn', 'clear_logo_posn'),
            ('logo_center', 'clear_logo_center'),
            ('logo_obec', 'clear_logo_obec'),
            ('logo_garuda', 'clear_logo_garuda'),
            ('site_favicon', 'clear_site_favicon'),
        ):
            if img_field in request.FILES:
                setattr(setting, img_field, process_logo_upload(request.FILES[img_field]))
            elif request.POST.get(clear_key):
                setattr(setting, img_field, None)
        if current_year_approved and int(setting.current_academic_year) == int(old_year):
            if int(setting.max_absences) != int(old_max_abs) or int(setting.max_qualified_count) != int(old_max_qual):
                setting.max_absences = old_max_abs
                setting.max_qualified_count = old_max_qual
                messages.warning(
                    request,
                    f'ปีการศึกษา {old_year} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนจำนวนครั้งขาดอบรมสูงสุดหรือจำนวนผู้ผ่านการคัดเลือกได้ (คงค่าเดิมไว้)',
                )
        setting.save()

        # 3. บันทึกสีเตือนภัยทั้งหมด
        for c in WarningColor.objects.all():
            cid = str(c.pk)
            if f'color_name_{cid}' in request.POST:
                cname = request.POST.get(f'color_name_{cid}', '').strip()
                ccode = request.POST.get(f'color_code_{cid}', c.color).strip()
                cmin = max(1, int(request.POST.get(f'color_min_{cid}') or c.min_absences))
                if cname:
                    c.name = cname
                if re.match(r'^#[0-9a-fA-F]{6}$', ccode):
                    c.color = ccode
                c.min_absences = cmin
                c.save()

        log_action(request.user, 'update_settings', target=setting, new_value='บันทึกการตั้งค่าระบบทั้งหมด',
                   request=request)
        messages.success(request, 'บันทึกการตั้งค่าระบบทั้งหมดเรียบร้อย')
        return redirect('composn:admin_settings')

    if request.method == 'POST' and request.POST.get('action') == 'save_theme':
        color_fields = [
            'color_monday', 'color_tuesday', 'color_wednesday', 'color_thursday',
            'color_friday', 'color_saturday', 'color_sunday',
            'member_theme_primary', 'member_theme_gradient',
            'staff_theme_primary', 'staff_theme_gradient',
            'login_subtitle', 'login_bg_color1', 'login_bg_color2', 'login_bg_color3',
            'login_btn_color1', 'login_btn_color2',
        ]
        for f in color_fields:
            if f in request.POST:
                setattr(setting, f, request.POST.get(f, '').strip())
        for img_field, clear_key in (
            ('login_logo', 'clear_login_logo'),
            ('logo_school', 'clear_logo_school'),
            ('logo_posn', 'clear_logo_posn'),
            ('logo_center', 'clear_logo_center'),
            ('logo_obec', 'clear_logo_obec'),
            ('logo_garuda', 'clear_logo_garuda'),
            ('site_favicon', 'clear_site_favicon'),
        ):
            if img_field in request.FILES:
                setattr(setting, img_field, process_logo_upload(request.FILES[img_field]))
            elif request.POST.get(clear_key):
                setattr(setting, img_field, None)
        setting.save()
        log_action(request.user, 'update_settings', target=setting, new_value='อัปเดตโทนสีและหน้า Login',
                   request=request)
        messages.success(request, 'บันทึกการตั้งค่าโทนสีและหน้าเข้าสู่ระบบเรียบร้อย')
        return redirect('composn:admin_settings')

    form = SystemSettingForm(request.POST or None, instance=setting)
    if request.method == 'POST' and form.is_valid():
        before = model_snapshot(SystemSetting.get(), ['current_academic_year', 'max_rights', 'max_absences',
                                                      'editable_profile_fields', 'site_title', 'site_title_line2',
                                                      'display_title'])
        obj = form.save(commit=False)
        if current_year_approved and int(obj.current_academic_year) == int(old_year):
            if int(obj.max_absences) != int(old_max_abs):
                obj.max_absences = old_max_abs
                messages.warning(
                    request,
                    f'ปีการศึกษา {old_year} อนุมัติประกาศผลแล้ว ไม่สามารถเปลี่ยนจำนวนครั้งขาดอบรมสูงสุดได้ (คงค่าเดิมไว้)',
                )
        obj.save()
        old, new = diff_snapshot(before, model_snapshot(obj, list(before)))
        log_action(request.user, 'update_settings', target=obj, old_value=old, new_value=new, request=request)
        messages.success(request, 'บันทึกการตั้งค่าเรียบร้อย')
        return redirect('composn:admin_settings')
    available_years_desc = get_academic_year_choices(setting.current_academic_year, setting=setting)
    available_years_asc = sorted(available_years_desc)
    max_academic_year = max(available_years_asc) if available_years_asc else MIN_ACADEMIC_YEAR
    return render(request, 'admin/settings.html', {
        'form': form, 'setting': setting,
        'reasons': AbsenceReason.objects.all(), 'colors': WarningColor.objects.all(),
        'available_years_asc': available_years_asc,
        'min_academic_year': MIN_ACADEMIC_YEAR,
        'max_academic_year': max_academic_year,
        'next_academic_year': min(MAX_ACADEMIC_YEAR, max_academic_year + 1),
        'current_year_approved': current_year_approved,
    })


# =============================================================================
# Backup / Export
# =============================================================================
@admin_required
def backups(request):
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'delete':
            p = safe_backup_path(request.POST.get('name', ''))
            if p:
                p.unlink()
                log_action(request.user, 'delete_backup', target_model='Backup', old_value=p.name, request=request)
                messages.success(request, f'ลบไฟล์ {p.name} แล้ว')
            else:
                messages.error(request, 'ไม่พบไฟล์')
        elif action == 'backup':
            try:
                year = int(request.POST.get('year') or SystemSetting.get().current_academic_year)
                camp_raw = request.POST.get('camp')
                camp = int(camp_raw) if camp_raw in ('1', '2') else rights.get_active_camp_for_year(year)[0]
                assert camp in (1, 2)
            except (TypeError, ValueError, AssertionError):
                messages.error(request, 'ข้อมูลไม่ถูกต้อง')
            else:
                p = backup_use_records(year, camp)
                log_action(request.user, 'backup', target_model='UseRecord', new_value={'file': p.name},
                           request=request)
                messages.success(request, f'สำรองข้อมูลเป็น {p.name} แล้ว')
        return redirect('composn:admin_backups')
    return render(request, 'admin/backups.html', {
        'files': list_backups(), 'year': SystemSetting.get().current_academic_year,
    })


@admin_required
def backup_download(request, name):
    p = safe_backup_path(name)
    if not p:
        raise Http404
    return FileResponse(open(p, 'rb'), as_attachment=True, filename=p.name)


@admin_required
def exports(request):
    from food.models import Attendance
    kind = request.GET.get('kind')
    if kind == 'students':
        qs = User.objects.filter(role=User.ROLE_MEMBER).annotate(
            used_count=Count('archive_records', filter=Q(archive_records__is_deleted=False)))
        log_action(request.user, 'export', target_model='User', new_value='รายชื่อนักเรียนทั้งหมด', request=request)
        return workbook_response(students_workbook(qs), 'รายชื่อนักเรียนทั้งหมด.xlsx')
    if kind == 'history':
        code = request.GET.get('code', '').strip()
        user = User.objects.filter(Q(camp1_code__iexact=code) | Q(camp2_code__iexact=code)).first()
        if not user:
            messages.error(request, f'ไม่พบรหัส {code}')
            return redirect('composn:admin_exports')
        log_action(request.user, 'export', target=user, new_value='ประวัติการใช้สิทธิ์รายบุคคล', request=request)
        return workbook_response(history_workbook(user), f'ประวัติใช้สิทธิ์_{user.camp1_code}.xlsx')
    if kind == 'all_history':
        rows = [[a.user.camp1_code, a.user.get_full_name(), a.academic_year, a.camp_level, a.rights_code, a.used_at]
                for a in ArchiveRecord.objects.select_related('user').order_by('user__camp1_code', 'used_at')]
        wb = build_workbook([('ARCHIVE_RECORD', ['รหัสค่าย 1', 'ชื่อ-นามสกุล', 'ปีการศึกษา', 'ค่าย',
                                                 'รหัสใช้สิทธิ์', 'ใช้สิทธิ์เมื่อ'], rows)])
        log_action(request.user, 'export', target_model='ArchiveRecord', new_value='ประวัติทั้งหมด', request=request)
        return workbook_response(wb, 'ประวัติการใช้สิทธิ์ทั้งหมด.xlsx')
    if kind == 'absence_monthly':
        month = request.GET.get('month') or timezone.localdate().strftime('%Y-%m')
        try:
            y, m = [int(x) for x in month.split('-')]
            start = dt.date(y, m, 1)
        except ValueError:
            messages.error(request, 'เดือนไม่ถูกต้อง')
            return redirect('composn:admin_exports')
        end = dt.date(y + (m == 12), m % 12 + 1, 1)
        qs = Attendance.objects.filter(date__gte=start, date__lt=end).select_related('user')
        summary = {}
        for a in qs:
            s = summary.setdefault(a.user_id, {
                'user': a.user, 'present': 0, 'absent': 0, 'sick': 0, 'personal': 0, 'total_absent': 0, 'reasons': {},
            })
            if a.status in s:
                s[a.status] += 1
            if a.status in Attendance.ABSENT_STATUSES:
                s['total_absent'] += 1
                if a.reason:
                    s['reasons'][a.reason] = s['reasons'].get(a.reason, 0) + 1
        rows = [[s['user'].camp1_code, s['user'].get_full_name(), s['present'], s['absent'],
                 s['sick'], s['personal'], s['total_absent'],
                 ', '.join(f'{k} {v}' for k, v in s['reasons'].items())]
                for s in sorted(summary.values(), key=lambda x: x['user'].camp1_code)]
        detail = [[a.date, a.get_session_period_display(), a.user.camp1_code, a.user.get_full_name(),
                   a.get_status_display(), a.reason, a.note] for a in qs.order_by('date', 'session_period')]
        wb = build_workbook([
            (f'สรุป {month}', ['รหัสค่าย 1', 'ชื่อ-นามสกุล', 'เข้าอบรม', 'ขาดอบรม', 'ลาป่วย', 'ลากิจ', 'รวมขาดอบรม', 'เหตุผล'], rows),
            ('รายละเอียด', ['วันที่', 'ช่วง', 'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'สถานะ', 'เหตุผล', 'หมายเหตุ'], detail),
        ])
        log_action(request.user, 'export', target_model='Attendance', new_value={'absence_monthly': month},
                   request=request)
        return workbook_response(wb, f'สรุปการขาดเรียน_{month}.xlsx')
    return render(request, 'admin/exports.html', {
        'month': timezone.localdate().strftime('%Y-%m'), 'year': SystemSetting.get().current_academic_year,
    })


# =============================================================================
# Audit Log / ถังขยะ
# =============================================================================
@admin_required
def audit_logs(request):
    from .security import get_active_lockouts, unlock_login
    if request.method == 'POST' and request.POST.get('action') == 'unlock_login':
        ltype = request.POST.get('lock_type', '').strip()
        lval = request.POST.get('lock_value', '').strip()
        if unlock_login(ltype, lval, by=request.user, request=request):
            messages.success(request, f'ปลดล็อกการเข้าสู่ระบบสำหรับ "{lval}" เรียบร้อยแล้ว')
        return redirect('composn:admin_audit_logs')
    qs = AuditLog.objects.select_related('user')
    q = request.GET.get('q', '').strip()
    action = request.GET.get('action', '')
    if q:
        qs = qs.filter(Q(user__camp1_code__icontains=q) | Q(new_value__icontains=q) | Q(old_value__icontains=q) |
                       Q(target_model__icontains=q) | Q(ip_address__icontains=q))
    if action:
        qs = qs.filter(action=action)
    params = request.GET.copy()
    params.pop('page', None)
    return render(request, 'admin/audit_logs.html', {
        'page': _paginate(request, qs, 100), 'q': q, 'action': action, 'action_labels': ACTION_LABELS,
        'querystring': params.urlencode(), 'active_lockouts': get_active_lockouts(),
    })


def _trash_models():
    from food.models import Attendance, Score
    return {'use_record': UseRecord, 'archive_record': ArchiveRecord, 'attendance': Attendance, 'score': Score}


@admin_required
def trash(request):
    ScoreApproval, approved_years_set, _, _, year_locked_message = _lock_helpers()
    models_map = _trash_models()
    if request.method == 'POST':
        model = models_map.get(request.POST.get('model'))
        if model:
            obj = get_object_or_404(model.all_objects, pk=request.POST.get('id'), is_deleted=True)
            obj_year = obj.category.academic_year if model.__name__ == 'Score' else getattr(obj, 'academic_year', None)
            if obj_year and ScoreApproval.is_year_approved(obj_year):
                messages.error(request, year_locked_message(obj_year))
                return redirect(request.get_full_path())
            try:
                with transaction.atomic():
                    obj.restore()
                    if model is UseRecord:
                        ArchiveRecord.all_objects.filter(use_record=obj, is_deleted=True).update(
                            is_deleted=False, deleted_at=None, deleted_by=None)
            except Exception as e:  # เช่น ชน unique (มีรายการใหม่แทนแล้ว)
                messages.error(request, f'ไม่สามารถกู้คืนได้: มีข้อมูลที่ซ้ำกันอยู่แล้ว ({e.__class__.__name__})')
            else:
                log_action(request.user, 'restore_record', target=obj, request=request)
                messages.success(request, 'กู้คืนข้อมูลเรียบร้อย')
        return redirect(request.get_full_path())
    approved_set = approved_years_set()
    sections = []
    labels = {'use_record': 'USE_RECORD', 'archive_record': 'ARCHIVE_RECORD', 'attendance': 'เวลาเรียน',
              'score': 'คะแนน'}
    for key, model in models_map.items():
        qs = model.all_objects.filter(is_deleted=True).select_related('deleted_by')
        if key == 'score':
            qs = qs.select_related('category')
        items = list(qs.order_by('-deleted_at')[:200])
        for item in items:
            iy = item.category.academic_year if key == 'score' else getattr(item, 'academic_year', None)
            item.is_year_locked = bool(iy in approved_set)
        sections.append({'key': key, 'label': labels[key], 'items': items})
    return render(request, 'admin/trash.html', {'sections': sections})
