"""
Views: Login / Logout / เปลี่ยนรหัสผ่าน / Member / Display สาธารณะ
"""
from django.contrib import messages
from django.contrib.auth import login, logout, update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.contrib.sessions.models import Session
from django.db import transaction
from django.db.models import Count, Q
from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from . import rights
from .audit import diff_snapshot, get_client_ip, log_action, model_snapshot
from .decorators import admin_required, member_required
from .excel import (USE_RECORD_HEADERS, build_workbook, history_workbook, use_record_rows,
                    workbook_response)
from .face_photo import FacePhotoError, process_and_validate_profile_photo
from .forms import ChangePasswordForm, LoginForm, ProfileForm, SchoolRequestForm
from .models import RightsRound, SchoolChangeRequest, SystemSetting, UseRecord, User
from .security import check_login_lockout, normalize_username, should_log_login_blocked


# =============================================================================
# Auth
# =============================================================================
def login_view(request):
    if request.user.is_authenticated:
        return redirect('composn:home')
    if request.method == 'POST':
        raw_username = request.POST.get('username', '')
        uname_norm = normalize_username(raw_username)
        client_ip = get_client_ip(request)
        lock = check_login_lockout(uname_norm, client_ip)
        if lock['locked']:
            if should_log_login_blocked(uname_norm, client_ip):
                log_action(
                    None,
                    'login_blocked',
                    target_model='LoginLock',
                    new_value={
                        'username': uname_norm,
                        'reason': lock['reason'],
                        'fail_count': lock['fail_count'],
                        'remaining_minutes': lock['remaining_minutes'],
                    },
                    request=request,
                )
            messages.error(
                request,
                f'มีการพยายามเข้าสู่ระบบไม่สำเร็จเกินกำหนด กรุณารอ {lock["remaining_minutes"]} นาที แล้วลองใหม่อีกครั้ง หรือติดต่อผู้ดูแลระบบ',
            )
            form = LoginForm(request, initial={'username': raw_username})
            return render(request, 'auth/login.html', {'form': form}, status=429)

        form = LoginForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            if not request.session.session_key:
                request.session.save()
            new_key = request.session.session_key
            kicked_old = False
            with transaction.atomic():
                locked_user = User.objects.select_for_update().get(pk=user.pk)
                old_key = locked_user.current_session_key
                kicked_old = bool(old_key and old_key != new_key)
                if kicked_old:
                    Session.objects.filter(session_key=old_key).delete()
                # ลบ session อื่นทั้งหมดของผู้ใช้คนนี้ที่อาจค้างอยู่ ให้เหลือได้เพียง 1 session เท่านั้น
                for sess in Session.objects.filter(expire_date__gte=timezone.now()).exclude(session_key=new_key):
                    try:
                        if str(sess.get_decoded().get('_auth_user_id', '')) == str(user.pk):
                            sess.delete()
                            kicked_old = True
                    except Exception:
                        pass
                User.objects.filter(pk=user.pk).update(current_session_key=new_key)
            user.current_session_key = new_key
            log_action(user, 'login', target=user, request=request,
                       new_value={'kicked_old_session': kicked_old})
            if user.is_first_login:
                messages.info(request, 'กรุณาเปลี่ยนรหัสผ่านก่อนใช้งานครั้งแรก')
                return redirect('composn:change_password')
            nxt = request.GET.get('next')
            if nxt:
                from django.utils.http import url_has_allowed_host_and_scheme
                if url_has_allowed_host_and_scheme(nxt, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
                    return redirect(nxt)
            return redirect('composn:home')
        log_action(None, 'login_failed', target_model='User',
                   new_value={'username': uname_norm}, request=request)
        after_lock = check_login_lockout(uname_norm, client_ip)
        if after_lock['locked']:
            messages.warning(
                request,
                f'กรอกรหัสผ่านผิดเกินกำหนด ระบบระงับการเข้าสู่ระบบชั่วคราว {after_lock["remaining_minutes"]} นาที',
            )
    else:
        form = LoginForm(request)
    if request.GET.get('timeout'):
        messages.warning(request, 'ออกจากระบบอัตโนมัติ เนื่องจากไม่มีการใช้งานเกินเวลาที่กำหนด')
    return render(request, 'auth/login.html', {'form': form})


def logout_view(request):
    if request.method != 'POST':
        # GET: แสดงหน้า login (ใช้สำหรับ auto logout ด้วย ?timeout=1)
        if request.user.is_authenticated and request.GET.get('timeout'):
            _do_logout(request)
            return redirect('/login/?timeout=1')
        return redirect('composn:home' if request.user.is_authenticated else 'composn:login')
    _do_logout(request)
    messages.success(request, 'ออกจากระบบเรียบร้อย')
    return redirect('composn:login')


def _do_logout(request):
    user = request.user
    if user.is_authenticated:
        log_action(user, 'logout', target=user, request=request)
        if user.current_session_key == request.session.session_key:
            User.objects.filter(pk=user.pk).update(current_session_key='')
    logout(request)


@login_required
def change_password(request):
    form = ChangePasswordForm(request.user, request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = request.user
        user.set_password(form.cleaned_data['new_password1'])
        user.is_first_login = False
        user.save()
        update_session_auth_hash(request, user)
        # session key อาจเปลี่ยน -> อัปเดตให้ตรง
        User.objects.filter(pk=user.pk).update(current_session_key=request.session.session_key or '')
        log_action(user, 'change_password', target=user, request=request)
        messages.success(request, 'เปลี่ยนรหัสผ่านเรียบร้อยแล้ว')
        return redirect('composn:home')
    return render(request, 'auth/change_password.html', {'form': form})


def session_ping(request):
    """ใช้ตรวจสอบสถานะ session จากฝั่ง client (JS)"""
    if not request.user.is_authenticated:
        return JsonResponse({'ok': False}, status=401)
    return JsonResponse({'ok': True})


@login_required
def home(request):
    u = request.user
    if u.is_admin_role:
        return redirect('composn:admin_dashboard')
    if u.role == User.ROLE_STAFF:
        if u.staff_can_attendance:
            return redirect('food:staff_attendance')
        if u.staff_can_scores:
            return redirect('food:staff_scores')
        if u.staff_can_manage_food:
            return redirect('food:orders')
        if u.staff_can_signin_sheet:
            return redirect('food:signin_sheet')
        if u.staff_can_view_users:
            return redirect('composn:admin_users')
        if u.staff_can_exam_seating:
            return redirect('food:exam_seating')
        messages.info(request, 'บัญชีเจ้าหน้าที่ของท่านยังไม่ได้รับมอบหมายสิทธิ์เมนู กรุณาติดต่อผู้ดูแลระบบ')
        return redirect('composn:display')
    return redirect('composn:profile')


# =============================================================================
# Member
# =============================================================================
@member_required
def profile(request):
    user = request.user
    setting = SystemSetting.get()
    allowed = setting.get_editable_fields()
    form = ProfileForm(request.POST or None, instance=user, allowed_fields=allowed) if allowed else None
    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'upload_photo':
            uploaded = request.FILES.get('profile_photo')
            try:
                cf, info = process_and_validate_profile_photo(uploaded, user_code=user.camp1_code)
            except FacePhotoError as e:
                messages.error(request, str(e))
            else:
                if user.profile_photo:
                    try:
                        user.profile_photo.delete(save=False)
                    except Exception:
                        pass
                user.profile_photo.save(cf.name, cf, save=True)
                log_action(user, 'update_profile', target=user,
                           new_value={'profile_photo': user.profile_photo.name, 'size_bytes': info['size_bytes']},
                           request=request)
                size_kb = max(1, info['size_bytes'] // 1024)
                messages.success(
                    request,
                    f'อัปโหลดรูปภาพประจำตัวเรียบร้อย (ครอปเฉพาะใบหน้า {info["width"]}×{info["height"]} px, ขนาด {size_kb} KB)'
                )
            return redirect('composn:profile')
        if action == 'delete_photo':
            if user.profile_photo:
                try:
                    user.profile_photo.delete(save=False)
                except Exception:
                    pass
                user.profile_photo = None
                user.save(update_fields=['profile_photo', 'updated_at'])
                log_action(user, 'update_profile', target=user, new_value={'profile_photo': 'deleted'}, request=request)
                messages.success(request, 'ลบรูปภาพประจำตัวแล้ว')
            return redirect('composn:profile')
        if action == 'profile':
            if not form:
                messages.error(request, 'ไม่มีฟิลด์ที่อนุญาตให้แก้ไข')
                return redirect('composn:profile')
            before = model_snapshot(User.objects.get(pk=user.pk), allowed)
            if form.is_valid():
                form.save()
                old, new = diff_snapshot(before, model_snapshot(user, allowed))
                if new:
                    log_action(user, 'update_profile', target=user, old_value=old, new_value=new, request=request)
                messages.success(request, 'บันทึกข้อมูลเรียบร้อย')
                return redirect('composn:profile')

    from food.certificate_pdf import check_certificate_eligibility
    from food.models import CertificateYearStatus, ScoreApproval
    summary = rights.member_rights_summary(user)
    open_request = SchoolChangeRequest.objects.filter(user=user, is_cleared=False).first()
    raw_levels = [
        (y, '-' if (not lv or lv == 'หมดสิทธิ์') else lv)
        for y, lv in user.future_levels()
    ]
    # แสดง 5 ปีการศึกษาตามรูปแบบมาตรฐาน (หากปีที่ 6 ยังไม่จบ ม.6/ปวช.3 ให้แสดงครบ 6 ปี)
    selected_levels = raw_levels if (len(raw_levels) > 5 and raw_levels[5][1] != '-') else raw_levels[:5]
    use_map = {}
    for rec in UseRecord.objects.filter(user=user).only('academic_year', 'is_used'):
        ry = int(rec.academic_year)
        if ry not in use_map or rec.is_used:
            use_map[ry] = bool(rec.is_used)
    base_y = int(user.base_academic_year) if user.base_academic_year else None
    profile_levels = []
    for y, lv in selected_levels:
        y_int = int(y) if y else None
        if base_y is not None and y_int == base_y:
            st_kind, st_label = 'training', 'อบรม'
        elif y_int is not None and y_int in use_map:
            if use_map[y_int]:
                st_kind, st_label = 'rights', 'ใช้สิทธิ์'
            else:
                st_kind, st_label = 'waived', 'สละสิทธิ์'
        else:
            st_kind, st_label = 'none', '-'
        profile_levels.append({
            'year': y,
            'level': lv,
            'status_kind': st_kind,
            'status_label': st_label,
        })
    cert_year = int(user.base_academic_year or setting.current_academic_year)
    cert_eligibility = check_certificate_eligibility(user, cert_year)
    cert_year_published = CertificateYearStatus.is_year_published(cert_year)
    cohort_approved = ScoreApproval.is_year_approved(cert_year)
    show_profile_exam_status = (not summary['is_current_student']) or cohort_approved or bool(user.camp2_code)
    open_exam_seats = []
    try:
        from food.models import ExamSeatAssignment
        open_exam_seats = list(
            ExamSeatAssignment.objects.filter(
                user=user,
                status=ExamSeatAssignment.STATUS_ASSIGNED,
                plan__is_open=True,
                plan_room__is_open=True,
            ).select_related('plan__training_date', 'plan_room__room', 'seat')
            .order_by('plan__training_date__date', 'plan__session_period')
        )
    except Exception:
        open_exam_seats = []
    ctx = {
        'summary': summary,
        'profile_levels': profile_levels,
        'form': form,
        'editable_field_names': set(form.fields.keys()) if form else set(),
        'edit_mode_open': bool(form and form.errors),
        'school_form': SchoolRequestForm(),
        'open_request': open_request,
        'cert_year': cert_year,
        'cert_eligibility': cert_eligibility,
        'cert_year_published': cert_year_published,
        'show_profile_exam_status': show_profile_exam_status,
        'open_exam_seats': open_exam_seats,
    }
    return render(request, 'member/profile.html', ctx)


@member_required
def training_scores(request, year=None):
    """หน้าข้อมูลการอบรม แยกตามปีการศึกษา (เฉพาะปีที่เข้าอบรมและปีที่ใช้สิทธิ์)"""
    from food.models import member_year_results
    all_results = member_year_results(request.user)
    if year is not None:
        year_int = int(year)
        matched = next((r for r in all_results if int(r['year']) == year_int), None)
        if matched is None:
            messages.warning(request, f'ไม่พบข้อมูลการอบรมหรือการใช้สิทธิ์ในปีการศึกษา {year_int}')
            if all_results:
                return redirect('composn:training_scores_year', year=all_results[0]['year'])
            return render(request, 'member/training_scores.html', {
                'year_results': [],
                'selected_year': None,
                'all_year_results': [],
            })
        selected_list = [matched]
        selected_year = matched['year']
    elif all_results:
        selected_list = [all_results[0]]
        selected_year = all_results[0]['year']
    else:
        selected_list = []
        selected_year = None
    return render(request, 'member/training_scores.html', {
        'year_results': selected_list,
        'selected_year': selected_year,
        'all_year_results': all_results,
    })


@member_required
@require_POST
def confirm_rights(request):
    try:
        rec = rights.confirm_rights(request.user, request=request)
    except rights.RightsError as e:
        messages.error(request, str(e))
    else:
        messages.success(request, f'ยืนยันใช้สิทธิ์เรียบร้อย รหัสใช้สิทธิ์ของท่านคือ {rec.rights_code}')
    return redirect('composn:profile')


@member_required
@require_POST
def cancel_rights(request):
    try:
        old_code = rights.cancel_rights(request.user, request=request)
    except rights.RightsError as e:
        messages.error(request, str(e))
    else:
        messages.success(request, f'ยกเลิกการใช้สิทธิ์ (รหัส {old_code}) เรียบร้อยแล้ว ท่านสามารถกดใช้สิทธิ์ใหม่ได้ก่อนหมดเวลา')
    return redirect('composn:profile')


@member_required
@require_POST
def school_request(request):
    user = request.user
    if SchoolChangeRequest.objects.filter(user=user, is_cleared=False).exists():
        messages.error(request, 'ท่านส่งคำขอไปแล้ว ต้องรอให้ผู้ดูแลระบบเคลียร์สถานะก่อนจึงจะส่งใหม่ได้')
        return redirect('composn:profile')
    form = SchoolRequestForm(request.POST)
    if form.is_valid():
        req = SchoolChangeRequest.objects.create(
            user=user, old_school_name=user.school_name,
            new_school_name=form.cleaned_data['new_school_name'].strip(),
            reason=form.cleaned_data['reason'],
        )
        log_action(user, 'school_request', target=req, old_value=user.school_name,
                   new_value=req.new_school_name, request=request)
        messages.success(request, 'ส่งคำขอแก้ไขชื่อโรงเรียนแล้ว กรุณารอผู้ดูแลระบบอนุมัติ')
    else:
        messages.error(request, 'กรุณากรอกชื่อโรงเรียนใหม่')
    return redirect('composn:profile')


@member_required
def my_history_export(request):
    log_action(request.user, 'export', target_model='ArchiveRecord', new_value='ประวัติการใช้สิทธิ์ของตนเอง',
               request=request)
    return workbook_response(history_workbook(request.user), f'ประวัติใช้สิทธิ์_{request.user.camp1_code}.xlsx')


# =============================================================================
# Display สาธารณะ (ไม่ต้อง Login / ไม่แสดงข้อมูลส่วนตัวที่ละเอียดอ่อน)
# =============================================================================
def mask_name(first, last):
    first = (first or '').strip()
    last = (last or '').strip()
    return f'{first} {last[:1]}.' if last else first


def _display_data(center=None, year=None):
    setting = SystemSetting.get()
    if not year:
        year = setting.current_academic_year
    else:
        try:
            year = str(int(year))
        except (ValueError, TypeError):
            year = setting.current_academic_year
    open_rounds = list(RightsRound.objects.filter(is_open=True, academic_year=year))
    records = UseRecord.objects.filter(academic_year=year, is_used=True).select_related('user')
    yy = f'{int(year) % 100:02d}'
    members = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=int(year))
    )
    if center:
        records = records.filter(center_snapshot=center)
        members = members.filter(camp_center=center)
    records = records.order_by('camp_level', 'confirmed_at')
    rows = [{
        'rights_code': r.rights_code,
        'name': mask_name(r.user.first_name, r.user.last_name),
        'center': r.center_snapshot,
        'camp_level': r.camp_level,
        'grade_level': r.grade_level,
        'confirmed_at': r.confirmed_at,
    } for r in records]
    stats = records.aggregate(
        total=Count('id'), camp1=Count('id', filter=Q(camp_level=1)), camp2=Count('id', filter=Q(camp_level=2)))
    return {
        'year': year, 'rows': rows, 'stats': stats, 'open_rounds': open_rounds,
        'member_total': members.count(), 'center': center, 'setting': setting,
    }


@login_required
def display(request, center=None):
    center = center.upper() if center else None
    year = request.GET.get('year')
    ctx = _display_data(center, year=year)
    ctx['updated_at'] = timezone.now()
    return render(request, 'display/display.html', ctx)


@admin_required
def display_download(request):
    center = (request.GET.get('center') or '').upper() or None
    year = request.GET.get('year')
    data = _display_data(center if center in ('DS', 'KM') else None, year=year)
    qs = UseRecord.objects.filter(academic_year=data['year']).select_related('user')
    if data['center']:
        qs = qs.filter(center_snapshot=data['center'])
    qs = qs.order_by('camp_level', '-is_used', 'rights_code')
    wb = build_workbook([(f'สรุป {data["year"]}', USE_RECORD_HEADERS, use_record_rows(qs))])
    log_action(request.user, 'export', target_model='UseRecord', new_value={'display_report': data['year'],
                                                                             'center': center}, request=request)
    return workbook_response(wb, f'รายงานใช้สิทธิ์_{data["year"]}{"_" + center if center else ""}.xlsx')
