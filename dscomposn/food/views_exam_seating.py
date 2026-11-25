"""
Views สำหรับระบบจัดที่นั่งสอบ (Exam Seating Management)
"""
from urllib.parse import urlencode

from django.contrib import messages
from django.db import transaction
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from composn.audit import log_action
from composn.decorators import member_required, staff_perm_required
from composn.models import User

from .exam_seating_export import build_exam_seating_pdf, build_exam_seating_xlsx
from .exam_seating_service import (
    SeatingError,
    add_bench_block,
    add_single_seat,
    apply_default_reserves,
    assign_single_seat,
    auto_assign_multi_rooms,
    auto_assign_room,
    auto_close_plan_if_incomplete,
    build_room_layout_data,
    cancel_seat_assignment,
    check_plan_ready_to_open,
    delete_room_block,
    delete_seat,
    ensure_default_exam_rooms,
    estimate_room_assignable_count,
    generate_desk_grid_seats,
    get_exam_examinees,
    get_or_create_plan_for_date,
    move_or_swap_seat,
    save_room_proctors,
    set_seat_status,
    sync_plan_rooms,
    update_seat_meta,
)
from .models import (
    CertificateRecipient,
    ExamRoom,
    ExamRoomProctor,
    ExamSeat,
    ExamSeatAssignment,
    ExamSeatingPlan,
    ExamSeatingPlanRoom,
    TrainingDate,
    current_year,
)


def _seating_redirect(request, **overrides):
    params = {}
    for key in ('year', 'date_id', 'plan_id', 'room_id', 'view'):
        val = overrides.get(key)
        if val is None:
            val = request.POST.get(key) or request.GET.get(key)
        if val not in (None, ''):
            params[key] = val
    url = reverse('food:exam_seating')
    if params:
        url = f'{url}?{urlencode(params)}'
    return redirect(url)


def _safe_int(val, default=None):
    try:
        return int(val)
    except (TypeError, ValueError):
        return default


def _compute_default_examinee_range(examinees, all_asgns, active_plan_room, active_layout, active_pattern, active_alternate):
    total_ex = len(examinees)
    first_unseated_idx = next(
        (idx for idx, ex in enumerate(examinees, start=1) if ex['assignment'] is None),
        None,
    )
    if first_unseated_idx is None:
        return 1, (total_ex if total_ex > 0 else 1)

    other_room_stats = {}
    active_pr_id = active_plan_room.pk if active_plan_room else None
    for a in all_asgns:
        if a.status != ExamSeatAssignment.STATUS_ASSIGNED or not a.user_id:
            continue
        if active_pr_id is not None and a.plan_room_id == active_pr_id:
            continue
        st = other_room_stats.setdefault(a.plan_room_id, {'count': 0, 'latest': None})
        st['count'] += 1
        key = (a.assigned_at, a.pk)
        if st['latest'] is None or key > st['latest']:
            st['latest'] = key

    if other_room_stats:
        last_pr = max(other_room_stats.values(), key=lambda x: x['latest'])
        estimated_span = last_pr['count']
    elif active_plan_room is not None:
        estimated_span = estimate_room_assignable_count(
            active_plan_room,
            pattern=active_pattern,
            alternate=active_alternate,
            skip_adjacent_to_reserve=True,
        )
        if estimated_span <= 0:
            free_cnt = (active_layout or {}).get('free_count', 0)
            estimated_span = free_cnt if free_cnt > 0 else 25
    else:
        estimated_span = 25

    range_to = min(total_ex, first_unseated_idx + max(1, estimated_span) - 1)
    return first_unseated_idx, max(first_unseated_idx, range_to)


# =============================================================================
# หน้าหลักจัดการที่นั่งสอบ
# =============================================================================
@staff_perm_required('staff_can_exam_seating')
def exam_seating_index(request):
    ensure_default_exam_rooms()
    year = _safe_int(request.GET.get('year'), current_year())

    exam_dates = list(
        TrainingDate.objects.filter(
            academic_year=year, event_type=TrainingDate.TYPE_EXAM,
        ).order_by('date', 'id')
    )
    selected_date = None
    if exam_dates:
        req_date_id = _safe_int(request.GET.get('date_id'))
        if req_date_id:
            selected_date = next((d for d in exam_dates if d.pk == req_date_id), None)
        if selected_date is None:
            selected_date = exam_dates[0]

    # สร้างแผนหลักให้ครบทุกวันสอบเพื่อให้ตั้งเปิด/ปิดให้ผู้เข้าสอบดูได้รายวัน
    exam_date_items = []
    plans_by_td_id = {}
    for ed in exam_dates:
        p = get_or_create_plan_for_date(ed, actor=request.user)
        plans_by_td_id[ed.pk] = p

    all_plan_ids = [p.pk for p in plans_by_td_id.values()]
    seated_counts_by_plan = {}
    open_rooms_by_plan = {}
    if all_plan_ids:
        for pid in all_plan_ids:
            seated_counts_by_plan[pid] = 0
            open_rooms_by_plan[pid] = 0
        for row in ExamSeatAssignment.objects.filter(
            plan_id__in=all_plan_ids, status=ExamSeatAssignment.STATUS_ASSIGNED,
        ).values_list('plan_id', flat=True):
            seated_counts_by_plan[row] = seated_counts_by_plan.get(row, 0) + 1
        for row in ExamSeatingPlanRoom.objects.filter(
            plan_id__in=all_plan_ids, is_open=True,
        ).values_list('plan_id', flat=True):
            open_rooms_by_plan[row] = open_rooms_by_plan.get(row, 0) + 1

    for ed in exam_dates:
        p = plans_by_td_id[ed.pk]
        exam_date_items.append({
            'training_date': ed,
            'plan': p,
            'seated_count': seated_counts_by_plan.get(p.pk, 0),
            'open_rooms_count': open_rooms_by_plan.get(p.pk, 0),
        })

    selected_plan = None
    plan_rooms = []
    active_plan_room = None
    view_mode = request.GET.get('view') or 'room'
    if view_mode not in ('room', 'overview'):
        view_mode = 'room'

    if selected_date is not None:
        req_plan_id = _safe_int(request.GET.get('plan_id'))
        if req_plan_id:
            selected_plan = ExamSeatingPlan.objects.filter(pk=req_plan_id, training_date=selected_date).first()
        if selected_plan is None:
            selected_plan = plans_by_td_id.get(selected_date.pk) or get_or_create_plan_for_date(selected_date, actor=request.user)
        plan_rooms = sync_plan_rooms(selected_plan, actor=request.user)
        if view_mode == 'room' and plan_rooms:
            req_room_id = _safe_int(request.GET.get('room_id'))
            if req_room_id:
                active_plan_room = next((pr for pr in plan_rooms if pr.pk == req_room_id), None)
            if active_plan_room is None:
                active_plan_room = next((pr for pr in plan_rooms if pr.is_open), plan_rooms[0])

    examinees = get_exam_examinees(year)
    examinee_map = {ex['user_id']: ex for ex in examinees}
    code_by_user_id = {ex['user_id']: ex['display_code'] for ex in examinees}
    rights_user_ids = {ex['user_id'] for ex in examinees if ex['is_rights']}

    plan_assignments_by_user = {}
    room_summaries = []
    active_layout = None
    active_proctors = []
    active_room_seats = []
    other_plan_rooms_for_move = []
    proctor_other_rooms = {}
    all_asgns = []

    if selected_plan is not None:
        all_asgns = list(
            ExamSeatAssignment.objects.filter(plan=selected_plan)
            .select_related('user', 'seat', 'plan_room__room')
        )
        by_pr_seat = {}
        for a in all_asgns:
            by_pr_seat.setdefault(a.plan_room_id, {})[a.seat_id] = a
            if a.status == ExamSeatAssignment.STATUS_ASSIGNED and a.user_id:
                plan_assignments_by_user[a.user_id] = a

        # สรุปข้อมูลแต่ละห้องในแผน
        all_proctors = list(
            ExamRoomProctor.objects.filter(plan_room__plan=selected_plan)
            .select_related('plan_room__room', 'recipient')
            .order_by('sort_order', 'id')
        )
        proctors_by_pr = {}
        for p in all_proctors:
            proctors_by_pr.setdefault(p.plan_room_id, []).append(p)
            if p.recipient_id and active_plan_room and p.plan_room_id != active_plan_room.pk:
                proctor_other_rooms.setdefault(p.recipient_id, []).append(p.plan_room.room.name)

        for pr in plan_rooms:
            pr_asgns = by_pr_seat.get(pr.pk, {})
            layout = build_room_layout_data(
                pr.room,
                assignment_by_seat_id=pr_asgns,
                code_by_user_id=code_by_user_id,
                rights_user_ids=rights_user_ids,
            )
            pr_proctors = proctors_by_pr.get(pr.pk, [])
            room_summaries.append({
                'plan_room': pr,
                'layout': layout,
                'proctors': pr_proctors,
            })
            if active_plan_room and pr.pk == active_plan_room.pk:
                active_layout = layout
                active_proctors = pr_proctors
                active_room_seats = list(
                    pr.room.seats.filter(is_active=True).order_by('seat_number', 'block_index', 'row_index', 'col_index')
                )
            else:
                other_plan_rooms_for_move.append({
                    'plan_room': pr,
                    'seats': list(
                        pr.room.seats.filter(is_active=True).order_by('seat_number', 'seat_label')
                    ),
                    'assigned_seat_ids': {
                        sid: a for sid, a in pr_asgns.items()
                    },
                })

    # เติมข้อมูลที่นั่งปัจจุบันของผู้เข้าสอบแต่ละคน
    seated_count = 0
    unseated_count = 0
    for ex in examinees:
        asgn = plan_assignments_by_user.get(ex['user_id'])
        ex['assignment'] = asgn
        if asgn:
            seated_count += 1
        else:
            unseated_count += 1

    unseated_examinees = [ex for ex in examinees if ex['assignment'] is None]
    lecturers = list(
        CertificateRecipient.objects.filter(
            academic_year=year, recipient_type=CertificateRecipient.TYPE_LECTURER,
        ).order_by('sort_order', 'full_name')
    )
    committees = list(
        CertificateRecipient.objects.filter(
            academic_year=year, recipient_type=CertificateRecipient.TYPE_COMMITTEE,
        ).order_by('sort_order', 'full_name')
    )

    active_proctor_recipient_ids = {p.recipient_id for p in active_proctors if p.recipient_id}
    active_proctor_custom_names = '\n'.join(
        p.display_name for p in active_proctors if not p.recipient_id
    )

    # สร้างข้อความแผนการเรียงของห้องปัจจุบัน
    active_pattern = ''
    active_alternate = True
    if active_plan_room:
        active_pattern = active_plan_room.effective_fill_pattern
        active_alternate = active_plan_room.alternate_seats

    selected_plan_readiness = None
    for item in exam_date_items:
        item['readiness'] = check_plan_ready_to_open(item['plan'], examinees=examinees)
        if selected_plan is not None and item['plan'].pk == selected_plan.pk:
            selected_plan_readiness = item['readiness']
    if selected_plan is not None and selected_plan_readiness is None:
        selected_plan_readiness = check_plan_ready_to_open(selected_plan, examinees=examinees)

    default_range_from, default_range_to = _compute_default_examinee_range(
        examinees, all_asgns, active_plan_room, active_layout, active_pattern, active_alternate,
    )

    ctx = {
        'year': year,
        'exam_dates': exam_dates,
        'exam_date_items': exam_date_items,
        'selected_date': selected_date,
        'selected_plan': selected_plan,
        'selected_plan_readiness': selected_plan_readiness,
        'plan_rooms': plan_rooms,
        'room_summaries': room_summaries,
        'active_plan_room': active_plan_room,
        'active_layout': active_layout,
        'active_proctors': active_proctors,
        'active_proctor_recipient_ids': active_proctor_recipient_ids,
        'active_proctor_custom_names': active_proctor_custom_names,
        'active_room_seats': active_room_seats,
        'active_pattern': active_pattern,
        'active_alternate': active_alternate,
        'other_plan_rooms_for_move': other_plan_rooms_for_move,
        'view_mode': view_mode,
        'examinees': examinees,
        'unseated_examinees': unseated_examinees,
        'total_examinees': len(examinees),
        'seated_count': seated_count,
        'unseated_count': unseated_count,
        'default_range_from': default_range_from,
        'default_range_to': default_range_to,
        'open_rooms_count': sum(1 for pr in plan_rooms if pr.is_open),
        'lecturers': lecturers,
        'committees': committees,
        'proctor_other_rooms': proctor_other_rooms,
        'fill_pattern_choices': ExamRoom.FILL_PATTERN_CHOICES,
        'examinee_map': examinee_map,
    }
    return render(request, 'food/exam_seating.html', ctx)


def _auto_close_warning_if_needed(request, plan):
    closed, readiness = auto_close_plan_if_incomplete(plan)
    if closed:
        reason = readiness.get('reason_text') or 'ข้อมูลการจัดที่นั่งยังไม่ครบถ้วน'
        messages.warning(
            request,
            f'ระบบได้ปิดการแสดงที่นั่งสอบของวันนี้อัตโนมัติ ({reason})',
        )
    return closed, readiness


# =============================================================================
# Actions ระดับแผนสอบ / ห้องในแผน
# =============================================================================
@staff_perm_required('staff_can_exam_seating')
@require_POST
def exam_plan_action(request):
    action = (request.POST.get('action') or '').strip()
    year = _safe_int(request.POST.get('year'), current_year())

    plan = get_object_or_404(ExamSeatingPlan, pk=request.POST.get('plan_id'))

    if action == 'toggle_plan_open':
        return_date_id = _safe_int(request.POST.get('date_id')) or plan.training_date_id
        return_plan_id = _safe_int(request.POST.get('return_plan_id')) or ''
        if not plan.is_open:
            readiness = check_plan_ready_to_open(plan)
            if not readiness['is_ready']:
                messages.error(
                    request,
                    f'ไม่สามารถเปิดให้ผู้เข้าสอบดูที่นั่งสอบในวันนี้ได้ เนื่องจากยังจัดห้องไม่เรียบร้อย: {readiness["reason_text"]}',
                )
                return _seating_redirect(request, year=year, date_id=return_date_id, plan_id=return_plan_id)

        plan.is_open = not plan.is_open
        plan.save(update_fields=['is_open', 'updated_at'])
        log_action(request.user, 'exam_plan_toggle', target=plan, new_value={'is_open': plan.is_open}, request=request)
        state_lbl = 'เปิดให้ผู้เข้าสอบดูที่นั่งสอบในวันดังกล่าวแล้ว' if plan.is_open else 'ปิดการแสดงที่นั่งสอบในวันดังกล่าวแล้ว'
        messages.success(request, state_lbl)
        return _seating_redirect(request, year=year, date_id=return_date_id, plan_id=return_plan_id)

    if action == 'delete_plan':
        td_id = plan.training_date_id
        if ExamSeatingPlan.objects.filter(training_date_id=td_id).count() <= 1:
            messages.warning(request, 'ต้องมีแผนจัดสอบอย่างน้อย 1 รอบในวันสอบนี้')
            return _seating_redirect(request, year=year, date_id=td_id, plan_id=plan.pk)
        log_action(request.user, 'exam_plan_delete', target=plan, old_value={'id': plan.pk}, request=request)
        plan.delete()
        messages.success(request, 'ลบรอบจัดที่นั่งสอบเรียบร้อย')
        return _seating_redirect(request, year=year, date_id=td_id, plan_id='')

    if action == 'clear_plan':
        deleted_cnt, _ = ExamSeatAssignment.objects.filter(
            plan=plan, status=ExamSeatAssignment.STATUS_ASSIGNED,
        ).delete()
        log_action(request.user, 'exam_plan_clear', target=plan, new_value={'deleted': deleted_cnt}, request=request)
        messages.success(request, f'ล้างการจัดผู้เข้าสอบทั้งหมดในวันสอบนี้แล้ว ({deleted_cnt} ที่นั่ง)')
        _auto_close_warning_if_needed(request, plan)
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk)

    if action == 'auto_assign_multi':
        pr_ids = request.POST.getlist('plan_room_ids')
        user_ids = request.POST.getlist('user_ids')
        if not pr_ids:
            messages.warning(request, 'กรุณาเลือกอย่างน้อย 1 ห้องสอบ')
            return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk)
        if not user_ids:
            # ใช้รายชื่อผู้เข้าสอบที่ยังไม่ถูกจัดทั้งหมดตามลำดับ
            examinees = get_exam_examinees(plan.academic_year)
            seated_ids = set(
                ExamSeatAssignment.objects.filter(plan=plan, status=ExamSeatAssignment.STATUS_ASSIGNED, user__isnull=False)
                .values_list('user_id', flat=True)
            )
            user_ids = [ex['user_id'] for ex in examinees if ex['user_id'] not in seated_ids]
        alternate = request.POST.get('alternate') == '1'
        skip_adj = request.POST.get('skip_adjacent_to_reserve', '1') == '1'
        res = auto_assign_multi_rooms(
            plan, pr_ids, user_ids, alternate=alternate, skip_adjacent_to_reserve=skip_adj, actor=request.user,
        )
        detail = ', '.join(f"{item['plan_room'].room.name}: {item['count']} คน" for item in res['per_room_counts'])
        log_action(request.user, 'exam_auto_assign_multi', target=plan, new_value={'total': res['total_assigned'], 'detail': detail}, request=request)
        msg = f'จัดที่นั่งกระจายหลายห้องเรียบร้อย {res["total_assigned"]} คน ({detail})'
        if res['remaining_count']:
            msg += f' · คงเหลือที่ยังไม่มีที่นั่ง {res["remaining_count"]} คน'
        messages.success(request, msg)
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk)

    plan_room = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('plan_room_id'), plan=plan)

    if action == 'toggle_room_open':
        plan_room.is_open = not plan_room.is_open
        plan_room.save(update_fields=['is_open'])
        log_action(request.user, 'exam_room_plan_toggle', target=plan_room, new_value={'is_open': plan_room.is_open}, request=request)
        messages.success(
            request,
            f'{"เปิด" if plan_room.is_open else "ปิด"}ห้อง {plan_room.room.name} ในวันสอบนี้เรียบร้อย',
        )
        if not plan_room.is_open:
            _auto_close_warning_if_needed(request, plan)
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    if action == 'update_room_meta':
        plan_room.exam_subtitle = (request.POST.get('exam_subtitle') or '').strip()
        plan_room.note = (request.POST.get('note') or '').strip()
        plan_room.save(update_fields=['exam_subtitle', 'note'])
        messages.success(request, f'บันทึกข้อมูลห้อง {plan_room.room.name} เรียบร้อย')
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    if action == 'save_proctors':
        rec_ids = request.POST.getlist('recipient_ids')
        name_style = request.POST.get('name_style') or 'short'
        custom_names = request.POST.get('custom_names') or ''
        created = save_room_proctors(plan_room, recipient_ids=rec_ids, name_style=name_style, custom_names=custom_names)
        log_action(
            request.user, 'exam_save_proctors', target=plan_room,
            new_value={'proctors': [p.display_name for p in created]}, request=request,
        )
        messages.success(request, f'บันทึกวิทยากรคุมสอบห้อง {plan_room.room.name} ({len(created)} ท่าน) เรียบร้อย')
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    if action == 'auto_assign_room':
        user_ids = request.POST.getlist('user_ids')
        if not user_ids:
            messages.warning(request, 'กรุณาเลือกผู้เข้าสอบอย่างน้อย 1 คนก่อนกดจัดที่นั่ง')
            return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)
        pattern = request.POST.get('pattern') or plan_room.effective_fill_pattern
        start_seat_id = _safe_int(request.POST.get('start_seat_id'))
        alternate = request.POST.get('alternate') == '1'
        skip_adj = request.POST.get('skip_adjacent_to_reserve', '1') == '1'
        clear_existing = request.POST.get('clear_existing_students') == '1'

        cnt, remaining = auto_assign_room(
            plan_room,
            user_ids,
            pattern=pattern,
            start_seat_id=start_seat_id,
            alternate=alternate,
            skip_adjacent_to_reserve=skip_adj,
            clear_existing_students=clear_existing,
            actor=request.user,
        )
        log_action(
            request.user, 'exam_auto_assign_room', target=plan_room,
            new_value={'assigned': cnt, 'pattern': pattern, 'alternate': alternate, 'remaining': len(remaining)},
            request=request,
        )
        msg = f'จัดที่นั่งลงห้อง {plan_room.room.name} เรียบร้อย {cnt} คน'
        if remaining:
            msg += f' (ที่นั่งว่างตามเงื่อนไขเต็มแล้ว เหลือ {len(remaining)} คน)'
        messages.success(request, msg)
        if clear_existing:
            _auto_close_warning_if_needed(request, plan)
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    if action == 'clear_room':
        include_reserves = request.POST.get('include_reserves') == '1'
        qs = ExamSeatAssignment.objects.filter(plan_room=plan_room)
        if not include_reserves:
            qs = qs.filter(status=ExamSeatAssignment.STATUS_ASSIGNED)
        cnt, _ = qs.delete()
        log_action(request.user, 'exam_clear_room', target=plan_room, new_value={'deleted': cnt}, request=request)
        messages.success(request, f'ล้างที่นั่งห้อง {plan_room.room.name} เรียบร้อย ({cnt} ตำแหน่ง)')
        _auto_close_warning_if_needed(request, plan)
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    if action == 'apply_default_reserves':
        cnt = apply_default_reserves(plan_room, actor=request.user)
        messages.success(request, f'นำเข้าตำแหน่งเครื่องสำรองเริ่มต้นของห้อง {plan_room.room.name} แล้ว ({cnt} ตำแหน่ง)')
        return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)

    messages.warning(request, 'ไม่พบคำสั่งที่ร้องขอ')
    return _seating_redirect(request, year=year, date_id=plan.training_date_id, plan_id=plan.pk, room_id=plan_room.pk)


# =============================================================================
# Actions ระดับที่นั่งสอบ (จัดรายคน, ยกเลิก, ย้าย/สลับ, ตั้งค่าสำรอง/งดใช้)
# =============================================================================
@staff_perm_required('staff_can_exam_seating')
@require_POST
def exam_seat_action(request):
    is_ajax = request.headers.get('x-requested-with') == 'XMLHttpRequest'
    action = (request.POST.get('action') or '').strip()

    def _err(msg, status=400):
        if is_ajax:
            return JsonResponse({'ok': False, 'error': msg}, status=status)
        messages.error(request, msg)
        return _seating_redirect(request)

    def _ok(msg, plan_room=None, plan_auto_closed=False, auto_close_reason=''):
        if is_ajax:
            payload = {'ok': True, 'message': msg, 'plan_auto_closed': plan_auto_closed}
            if plan_auto_closed:
                payload['message'] = f'{msg} (ระบบปิดการเปิดดูอัตโนมัติ: {auto_close_reason})'
                payload['auto_close_reason'] = auto_close_reason
            return JsonResponse(payload)
        messages.success(request, msg)
        if plan_auto_closed:
            messages.warning(
                request,
                f'ระบบได้ปิดการแสดงที่นั่งสอบของวันนี้อัตโนมัติ ({auto_close_reason})',
            )
        kw = {}
        if plan_room is not None:
            kw = {
                'year': plan_room.plan.academic_year,
                'date_id': plan_room.plan.training_date_id,
                'plan_id': plan_room.plan_id,
                'room_id': plan_room.pk,
            }
        return _seating_redirect(request, **kw)

    try:
        if action == 'assign_seat':
            pr = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('plan_room_id'))
            seat = get_object_or_404(ExamSeat, pk=request.POST.get('seat_id'))
            student = get_object_or_404(User, pk=request.POST.get('user_id'), role=User.ROLE_MEMBER)
            overwrite = request.POST.get('overwrite_reserve') == '1'
            asgn = assign_single_seat(pr, seat, student, actor=request.user, overwrite_reserve=overwrite)
            log_action(request.user, 'exam_seat_assign', target=asgn, new_value={'seat': seat.seat_label, 'user': student.camp1_code}, request=request)
            return _ok(f'จัด {student.camp1_code} ลงที่นั่ง {seat.seat_label} เรียบร้อย', pr)

        if action == 'cancel_seat':
            asgn_id = _safe_int(request.POST.get('assignment_id'))
            if asgn_id:
                asgn = get_object_or_404(ExamSeatAssignment, pk=asgn_id)
            else:
                pr = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('plan_room_id'))
                seat = get_object_or_404(ExamSeat, pk=request.POST.get('seat_id'))
                asgn = get_object_or_404(ExamSeatAssignment, plan_room=pr, seat=seat)
            pr = asgn.plan_room
            seat_lbl = asgn.seat.seat_label
            user_code = asgn.user.camp1_code if asgn.user else asgn.get_status_display()
            cancel_seat_assignment(asgn)
            log_action(request.user, 'exam_seat_cancel', target=pr, new_value={'seat': seat_lbl, 'removed': user_code}, request=request)
            closed, readiness = auto_close_plan_if_incomplete(pr.plan)
            return _ok(
                f'ยกเลิกที่นั่ง {seat_lbl} ({user_code}) เรียบร้อย',
                pr,
                plan_auto_closed=closed,
                auto_close_reason=readiness.get('reason_text', ''),
            )

        if action == 'move_or_swap':
            src_id = _safe_int(request.POST.get('source_assignment_id'))
            if src_id:
                src_asgn = get_object_or_404(ExamSeatAssignment, pk=src_id)
            else:
                src_pr = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('source_plan_room_id'))
                src_seat = get_object_or_404(ExamSeat, pk=request.POST.get('source_seat_id'))
                src_asgn = get_object_or_404(ExamSeatAssignment, plan_room=src_pr, seat=src_seat)

            tgt_pr_id = _safe_int(request.POST.get('target_plan_room_id')) or src_asgn.plan_room_id
            tgt_pr = get_object_or_404(ExamSeatingPlanRoom, pk=tgt_pr_id)
            tgt_seat = get_object_or_404(ExamSeat, pk=request.POST.get('target_seat_id'))
            src_label = f'{src_asgn.plan_room.room.name} #{src_asgn.seat.seat_label}'
            outcome = move_or_swap_seat(src_asgn, tgt_pr, tgt_seat, actor=request.user)
            log_action(
                request.user, 'exam_seat_move_swap', target=tgt_pr,
                new_value={'outcome': outcome, 'from': src_label, 'to': f'{tgt_pr.room.name} #{tgt_seat.seat_label}'},
                request=request,
            )
            closed, readiness = auto_close_plan_if_incomplete(tgt_pr.plan)
            if outcome == 'swapped':
                return _ok(
                    f'สลับที่นั่งระหว่าง {src_label} กับ {tgt_pr.room.name} #{tgt_seat.seat_label} เรียบร้อย',
                    tgt_pr,
                    plan_auto_closed=closed,
                    auto_close_reason=readiness.get('reason_text', ''),
                )
            return _ok(
                f'ย้ายไปที่นั่ง {tgt_pr.room.name} #{tgt_seat.seat_label} เรียบร้อย',
                tgt_pr,
                plan_auto_closed=closed,
                auto_close_reason=readiness.get('reason_text', ''),
            )

        if action == 'set_seat_status':
            pr = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('plan_room_id'))
            seat = get_object_or_404(ExamSeat, pk=request.POST.get('seat_id'))
            status = request.POST.get('status') or 'empty'
            set_seat_status(pr, seat, status, actor=request.user)
            lbl_map = {'reserve': 'เครื่องสำรอง', 'blocked': 'งดใช้งาน', 'empty': 'ที่นั่งว่าง'}
            closed, readiness = auto_close_plan_if_incomplete(pr.plan)
            return _ok(
                f'ตั้งค่าที่นั่ง {seat.seat_label} เป็น "{lbl_map.get(status, status)}" เรียบร้อย',
                pr,
                plan_auto_closed=closed,
                auto_close_reason=readiness.get('reason_text', ''),
            )

        if action == 'set_start_seat':
            pr = get_object_or_404(ExamSeatingPlanRoom, pk=request.POST.get('plan_room_id'))
            seat_id = _safe_int(request.POST.get('seat_id'))
            seat = get_object_or_404(ExamSeat, pk=seat_id, room=pr.room) if seat_id else None
            pr.start_seat = seat
            pr.save(update_fields=['start_seat'])
            return _ok(f'ตั้งค่าจุดเริ่มต้นที่นั่งหมายเลข {seat.seat_label} เรียบร้อย' if seat else 'ล้างค่าจุดเริ่มต้นแล้ว', pr)

        return _err('ไม่พบคำสั่งที่ร้องขอ')
    except SeatingError as exc:
        return _err(str(exc))


# =============================================================================
# จัดการห้องสอบ (สร้างห้องใหม่ / เปิด-ปิด / ออกแบบผังที่นั่งในห้อง)
# =============================================================================
@staff_perm_required('staff_can_exam_seating')
def exam_rooms_manage(request):
    ensure_default_exam_rooms()
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'create_room':
            name = (request.POST.get('name') or '').strip()
            if not name:
                messages.error(request, 'กรุณาระบุชื่อห้องสอบ')
                return redirect('food:exam_rooms_manage')
            if ExamRoom.objects.filter(name=name).exists():
                messages.error(request, f'มีห้องสอบชื่อ "{name}" อยู่ในระบบแล้ว')
                return redirect('food:exam_rooms_manage')

            has_computer = request.POST.get('has_computer') == '1'
            front_position = request.POST.get('front_position') or (ExamRoom.FRONT_LEFT if has_computer else ExamRoom.FRONT_TOP)
            if front_position not in dict(ExamRoom.FRONT_CHOICES):
                front_position = ExamRoom.FRONT_TOP
            front_label = (request.POST.get('front_label') or '').strip() or ('หน้าห้อง' if has_computer else 'หน้าเวที')
            default_fill_pattern = request.POST.get('default_fill_pattern') or (
                ExamRoom.FILL_SEQ_ASC if has_computer else ExamRoom.FILL_SNAKE_AUTO
            )
            if default_fill_pattern not in dict(ExamRoom.FILL_PATTERN_CHOICES):
                default_fill_pattern = ExamRoom.FILL_SEQ_ASC
            default_alternate = request.POST.get('default_alternate') == '1'
            sort_order = _safe_int(request.POST.get('sort_order'), ExamRoom.objects.count() + 1)

            with transaction.atomic():
                room = ExamRoom.objects.create(
                    name=name,
                    has_computer=has_computer,
                    front_position=front_position,
                    front_label=front_label,
                    default_fill_pattern=default_fill_pattern,
                    default_alternate=default_alternate,
                    is_active=True,
                    sort_order=sort_order,
                    description=(request.POST.get('description') or '').strip(),
                )
                init_mode = request.POST.get('init_mode') or 'grid'
                if init_mode == 'grid':
                    rows = _safe_int(request.POST.get('init_rows'), 5)
                    cols = _safe_int(request.POST.get('init_cols'), 5)
                    start_no = _safe_int(request.POST.get('init_start_no'), 1)
                    direction = request.POST.get('init_direction') or 'lr'
                    generate_desk_grid_seats(room, rows, cols, start_no=start_no, direction=direction, clear_existing=True)
            log_action(request.user, 'exam_room_create', target=room, new_value={'name': room.name, 'has_computer': room.has_computer}, request=request)
            messages.success(request, f'สร้างห้องสอบ "{room.name}" เรียบร้อย สามารถปรับแก้ตำแหน่งที่นั่งเพิ่มเติมได้ด้านล่าง')
            return redirect('food:exam_room_layout', room_id=room.pk)

        if action == 'toggle_room_active':
            room = get_object_or_404(ExamRoom, pk=request.POST.get('room_id'))
            room.is_active = not room.is_active
            room.save(update_fields=['is_active', 'updated_at'])
            messages.success(request, f'{"เปิด" if room.is_active else "ปิด"}ใช้งานห้อง "{room.name}" เรียบร้อย')
            return redirect('food:exam_rooms_manage')

        if action == 'delete_room':
            room = get_object_or_404(ExamRoom, pk=request.POST.get('room_id'))
            r_name = room.name
            log_action(request.user, 'exam_room_delete', target=room, old_value={'name': r_name}, request=request)
            room.delete()
            messages.success(request, f'ลบห้องสอบ "{r_name}" เรียบร้อย')
            return redirect('food:exam_rooms_manage')

    rooms = list(ExamRoom.objects.all().order_by('sort_order', 'id'))
    room_rows = []
    for r in rooms:
        room_rows.append({
            'room': r,
            'seat_count': r.seats.filter(is_active=True).count(),
            'reserve_count': r.seats.filter(is_active=True, default_status=ExamSeat.STATUS_RESERVE).count(),
        })
    return render(request, 'food/exam_rooms_manage.html', {
        'room_rows': room_rows,
        'front_choices': ExamRoom.FRONT_CHOICES,
        'fill_pattern_choices': ExamRoom.FILL_PATTERN_CHOICES,
    })


@staff_perm_required('staff_can_exam_seating')
def exam_room_layout(request, room_id):
    room = get_object_or_404(ExamRoom, pk=room_id)
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        try:
            if action == 'update_room':
                new_name = (request.POST.get('name') or '').strip()
                if not new_name:
                    raise SeatingError('ชื่อห้องสอบห้ามว่าง')
                if ExamRoom.objects.exclude(pk=room.pk).filter(name=new_name).exists():
                    raise SeatingError(f'มีห้องสอบชื่อ "{new_name}" อยู่แล้ว')
                room.name = new_name
                room.has_computer = request.POST.get('has_computer') == '1'
                fp = request.POST.get('front_position')
                if fp in dict(ExamRoom.FRONT_CHOICES):
                    room.front_position = fp
                room.front_label = (request.POST.get('front_label') or '').strip() or ('หน้าห้อง' if room.has_computer else 'หน้าเวที')
                dfp = request.POST.get('default_fill_pattern')
                if dfp in dict(ExamRoom.FILL_PATTERN_CHOICES):
                    room.default_fill_pattern = dfp
                room.default_alternate = request.POST.get('default_alternate') == '1'
                room.is_active = request.POST.get('is_active') == '1'
                room.sort_order = _safe_int(request.POST.get('sort_order'), room.sort_order)
                room.description = (request.POST.get('description') or '').strip()
                room.save()
                messages.success(request, 'บันทึกการตั้งค่าห้องสอบเรียบร้อย')

            elif action == 'generate_grid':
                has_assigned = ExamSeatAssignment.objects.filter(
                    seat__room=room, status=ExamSeatAssignment.STATUS_ASSIGNED,
                ).exists()
                if has_assigned and request.POST.get('confirm_reset') != '1':
                    raise SeatingError('ห้องนี้มีการจัดผู้เข้าสอบค้างอยู่ในบางวันสอบ หากต้องการสร้างผังใหม่ทับทั้งหมด กรุณาติ๊กยืนยันก่อน')
                rows = _safe_int(request.POST.get('rows'), 5)
                cols = _safe_int(request.POST.get('cols'), 5)
                start_no = _safe_int(request.POST.get('start_no'), 1)
                direction = request.POST.get('direction') or 'lr'
                cnt = generate_desk_grid_seats(room, rows, cols, start_no=start_no, direction=direction, clear_existing=True)
                messages.success(request, f'สร้างผังที่นั่ง {rows} แถว × {cols} คอลัมน์ ({cnt} ที่นั่ง) เรียบร้อย')

            elif action == 'add_block':
                bench_type = request.POST.get('bench_type') or 'single'
                count = _safe_int(request.POST.get('count'), 10)
                col_offset = _safe_int(request.POST.get('col_offset'), 0)
                start_no = _safe_int(request.POST.get('start_no'), 1)
                dir_top = request.POST.get('dir_top') or 'lr'
                dir_bottom = request.POST.get('dir_bottom') or 'rl'
                label_pos = request.POST.get('label_pos') or ExamSeat.LABEL_BELOW
                b_cnt = _safe_int(request.POST.get('bottom_count'))
                b_start = _safe_int(request.POST.get('bottom_start_no'))
                cnt = add_bench_block(
                    room, bench_type=bench_type, count=count, col_offset=col_offset,
                    start_no=start_no, dir_top=dir_top, dir_bottom=dir_bottom,
                    label_pos=label_pos, bottom_count=b_cnt, bottom_start_no=b_start,
                )
                messages.success(request, f'เพิ่มกลุ่มที่นั่งใหม่ {cnt} ตำแหน่งเรียบร้อย')

            elif action == 'add_seat':
                blk = _safe_int(request.POST.get('block_index'), 1)
                row = _safe_int(request.POST.get('row_index'), 1)
                col = _safe_int(request.POST.get('col_index'), 1)
                lbl = request.POST.get('seat_label') or ''
                s_num = request.POST.get('seat_number')
                lpos = request.POST.get('label_position') or ExamSeat.LABEL_BELOW
                dst = request.POST.get('default_status') or ExamSeat.STATUS_NORMAL
                seat = add_single_seat(room, blk, row, col, lbl, seat_number=s_num, label_position=lpos, default_status=dst)
                messages.success(request, f'เพิ่มที่นั่ง "{seat.seat_label}" เรียบร้อย')

            elif action == 'update_seat':
                seat = get_object_or_404(ExamSeat, pk=request.POST.get('seat_id'), room=room)
                update_seat_meta(
                    seat,
                    request.POST.get('seat_label'),
                    seat_number=request.POST.get('seat_number'),
                    default_status=request.POST.get('default_status'),
                    label_position=request.POST.get('label_position'),
                )
                messages.success(request, f'แก้ไขที่นั่ง "{seat.seat_label}" เรียบร้อย')

            elif action == 'delete_seat':
                seat = get_object_or_404(ExamSeat, pk=request.POST.get('seat_id'), room=room)
                lbl = seat.seat_label
                delete_seat(seat)
                messages.success(request, f'ลบที่นั่ง "{lbl}" เรียบร้อย')

            elif action == 'delete_block':
                blk = _safe_int(request.POST.get('block_index'))
                if blk:
                    cnt, _ = delete_room_block(room, blk)
                    messages.success(request, f'ลบกลุ่มแถวที่ {blk} ({cnt} ที่นั่ง) เรียบร้อย')
        except SeatingError as exc:
            messages.error(request, str(exc))
        return redirect('food:exam_room_layout', room_id=room.pk)

    layout = build_room_layout_data(room)
    seats = list(room.seats.filter(is_active=True).order_by('block_index', 'row_index', 'col_index'))
    max_num = max((s.seat_number for s in seats), default=0)
    next_block = max((s.block_index for s in seats), default=0) + 1
    next_row = max((s.row_index for s in seats), default=0) + 1

    return render(request, 'food/exam_room_layout.html', {
        'room': room,
        'layout': layout,
        'seats': seats,
        'next_seat_no': max_num + 1,
        'next_block': next_block,
        'next_row': next_row,
        'front_choices': ExamRoom.FRONT_CHOICES,
        'fill_pattern_choices': ExamRoom.FILL_PATTERN_CHOICES,
        'default_status_choices': ExamSeat.DEFAULT_STATUS_CHOICES,
        'label_pos_choices': ExamSeat.LABEL_POS_CHOICES,
    })


# =============================================================================
# พิมพ์ & ส่งออก (PDF / Excel)
# =============================================================================
def _resolve_export_rooms(plan):
    sync_plan_rooms(plan)
    qs = plan.plan_rooms.select_related('room').prefetch_related('proctors').order_by('sort_order', 'room__sort_order', 'id')
    open_rooms = list(qs.filter(is_open=True))
    with_seats = [pr for pr in open_rooms if pr.room.seats.filter(is_active=True).exists()]
    return with_seats or open_rooms


@staff_perm_required('staff_can_exam_seating')
def exam_seating_print(request, plan_id):
    plan = get_object_or_404(ExamSeatingPlan.objects.select_related('training_date'), pk=plan_id)
    plan_rooms = _resolve_export_rooms(plan)
    examinees = get_exam_examinees(plan.academic_year)
    code_by_user_id = {ex['user_id']: ex['display_code'] for ex in examinees}
    rights_user_ids = {ex['user_id'] for ex in examinees if ex['is_rights']}

    pages = []
    for pr in plan_rooms:
        asgn_by_seat = {
            a.seat_id: a
            for a in ExamSeatAssignment.objects.filter(plan_room=pr).select_related('user', 'seat')
        }
        layout = build_room_layout_data(
            pr.room,
            assignment_by_seat_id=asgn_by_seat,
            code_by_user_id=code_by_user_id,
            rights_user_ids=rights_user_ids,
        )
        proctors = list(pr.proctors.all().order_by('sort_order', 'id'))
        pages.append({
            'plan_room': pr,
            'layout': layout,
            'proctors': proctors,
        })

    return render(request, 'food/exam_seating_print.html', {
        'plan': plan,
        'pages': pages,
    })


@staff_perm_required('staff_can_exam_seating')
def exam_seating_export_pdf(request, plan_id):
    plan = get_object_or_404(ExamSeatingPlan.objects.select_related('training_date'), pk=plan_id)
    plan_rooms = _resolve_export_rooms(plan)
    examinee_map = {ex['user_id']: ex for ex in get_exam_examinees(plan.academic_year)}
    pdf_bytes = build_exam_seating_pdf(plan, plan_rooms, examinee_map)
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    fname = f'exam_seating_{plan.academic_year}_{plan.training_date.date}.pdf'
    resp['Content-Disposition'] = f'inline; filename="{fname}"'
    return resp


@staff_perm_required('staff_can_exam_seating')
def exam_seating_export_xlsx(request, plan_id):
    plan = get_object_or_404(ExamSeatingPlan.objects.select_related('training_date'), pk=plan_id)
    plan_rooms = _resolve_export_rooms(plan)
    examinee_map = {ex['user_id']: ex for ex in get_exam_examinees(plan.academic_year)}
    xlsx_bytes = build_exam_seating_xlsx(plan, plan_rooms, examinee_map)
    resp = HttpResponse(
        xlsx_bytes,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    fname = f'exam_seating_{plan.academic_year}_{plan.training_date.date}.xlsx'
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


# =============================================================================
# มุมมองของนักเรียน (Member): ดูที่นั่งสอบของตนเองที่เปิดประกาศแล้ว
# =============================================================================
@member_required
def member_exam_seating(request):
    from django.core.cache import cache

    user = request.user
    assignments = list(
        ExamSeatAssignment.objects.filter(
            user=user,
            status=ExamSeatAssignment.STATUS_ASSIGNED,
            plan__is_open=True,
            plan_room__is_open=True,
        ).select_related('plan__training_date', 'plan_room__room', 'seat')
        .prefetch_related('plan_room__proctors')
        .order_by('plan__training_date__date', 'plan__session_period')
    )

    items = []
    examinees_by_year = {}
    for asgn in assignments:
        pr = asgn.plan_room
        plan = asgn.plan
        if plan.academic_year not in examinees_by_year:
            ex_cache_key = f'dscomposn:exam_examinees:{plan.academic_year}'
            cached_ex = cache.get(ex_cache_key)
            if cached_ex is None:
                cached_ex = get_exam_examinees(plan.academic_year)
                cache.set(ex_cache_key, cached_ex, 30)
            examinees_by_year[plan.academic_year] = cached_ex
        examinees = examinees_by_year[plan.academic_year]
        code_by_user_id = {ex['user_id']: ex['display_code'] for ex in examinees}
        rights_user_ids = {ex['user_id'] for ex in examinees if ex['is_rights']}
        asgn_by_seat = {
            a.seat_id: a
            for a in ExamSeatAssignment.objects.filter(plan_room=pr).select_related('user', 'seat')
        }
        layout = build_room_layout_data(
            pr.room,
            assignment_by_seat_id=asgn_by_seat,
            code_by_user_id=code_by_user_id,
            rights_user_ids=rights_user_ids,
        )
        proctors = sorted(pr.proctors.all(), key=lambda p: (p.sort_order, p.id))
        items.append({
            'assignment': asgn,
            'plan': plan,
            'plan_room': pr,
            'seat': asgn.seat,
            'display_code': code_by_user_id.get(user.pk) or user.camp1_code,
            'layout': layout,
            'proctors': proctors,
        })

    return render(request, 'food/member_exam_seating.html', {
        'seating_items': items,
    })
