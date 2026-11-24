"""
Service Layer สำหรับระบบจัดที่นั่งสอบ (Exam Seating Management)
- จัดการข้อมูลห้องสอบเริ่มต้น 4 ห้อง (134, 135, 136, ห้องประชุมเทิดพระเกียรติ ชั้น 2)
- สร้าง/แก้ไขผังห้องสอบและที่นั่งสอบ
- จัดที่นั่งสอบอัตโนมัติ (แบบเครื่องเว้นเครื่อง, เรียงซ้าย→ขวา / ขวา→ซ้าย / สลับฟันปลา, จัดกระจายหลายห้อง)
- จัดวางรายคน, ยกเลิก, ย้าย, สลับที่นั่ง (ในห้องและข้ามห้อง)
- จัดการวิทยากรคุมสอบประจำห้อง
"""
import re

from django.db import transaction

from composn.models import UseRecord
from .models import (
    CertificateRecipient,
    ExamRoom,
    ExamRoomProctor,
    ExamSeat,
    ExamSeatAssignment,
    ExamSeatingPlan,
    ExamSeatingPlanRoom,
    score_members_for_year,
)


class SeatingError(Exception):
    """ข้อผิดพลาดทางธุรกิจของการจัดที่นั่งสอบ"""


def _build_default_rooms_spec():
    return [
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 134',
            'has_computer': True,
            'front_position': ExamRoom.FRONT_RIGHT,
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': ExamRoom.FILL_SEQ_ASC,
            'default_alternate': True,
            'sort_order': 1,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 134 (50 เครื่อง)',
            'reserves': {1, 16, 41, 49},
            'rows': [
                (1, 1, ExamSeat.LABEL_BELOW, 1, [15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]),
                (2, 2, ExamSeat.LABEL_ABOVE, 2, [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27]),
                (2, 3, ExamSeat.LABEL_BELOW, 2, [39, 38, 37, 36, 35, 34, 33, 32, 31, 30, 29, 28]),
                (3, 4, ExamSeat.LABEL_ABOVE, 3, [40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]),
            ],
        },
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 135',
            'has_computer': True,
            'front_position': ExamRoom.FRONT_LEFT,
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': ExamRoom.FILL_SEQ_ASC,
            'default_alternate': True,
            'sort_order': 2,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 135 (55 เครื่อง)',
            'reserves': {16, 22, 30, 54},
            'rows': [
                (1, 1, ExamSeat.LABEL_BELOW, 1, list(range(1, 16))),
                (2, 2, ExamSeat.LABEL_ABOVE, 1, list(range(29, 15, -1))),
                (2, 3, ExamSeat.LABEL_BELOW, 1, list(range(30, 44))),
                (3, 4, ExamSeat.LABEL_ABOVE, 1, list(range(55, 43, -1))),
            ],
        },
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 136',
            'has_computer': True,
            'front_position': ExamRoom.FRONT_LEFT,
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': ExamRoom.FILL_SEQ_ASC,
            'default_alternate': True,
            'sort_order': 3,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 136 (55 เครื่อง)',
            'reserves': {11, 30, 40, 42, 50, 55, 57},
            'rows': [
                (1, 1, ExamSeat.LABEL_BELOW, 1, list(range(1, 9))),
                (2, 2, ExamSeat.LABEL_ABOVE, 1, list(range(20, 10, -1))),
                (2, 3, ExamSeat.LABEL_BELOW, 1, list(range(21, 31))),
                (3, 4, ExamSeat.LABEL_ABOVE, 1, list(range(40, 30, -1))),
                (3, 5, ExamSeat.LABEL_BELOW, 1, list(range(41, 51))),
                (4, 6, ExamSeat.LABEL_ABOVE, 4, list(range(57, 50, -1))),
            ],
        },
        {
            'name': 'ห้องประชุมเทิดพระเกียรติ ชั้น 2',
            'has_computer': False,
            'front_position': ExamRoom.FRONT_TOP,
            'front_label': 'หน้าเวที',
            'default_fill_pattern': ExamRoom.FILL_SNAKE_AUTO,
            'default_alternate': False,
            'sort_order': 4,
            'description': 'ห้องประชุมเทิดพระเกียรติ ชั้น 2 (โต๊ะสอบ 14 แถว 70 ที่นั่ง ไม่มีคอมพิวเตอร์)',
            'reserves': set(),
            'rows': [
                (r, r, ExamSeat.LABEL_BELOW, 1, [(r - 1) * 5 + c for c in range(1, 6)])
                for r in range(1, 15)
            ],
        },
    ]


@transaction.atomic
def ensure_default_exam_rooms():
    """สร้างห้องสอบเริ่มต้นทั้ง 4 ห้องหากยังไม่มีห้องสอบในระบบเลย"""
    if ExamRoom.objects.exists():
        return
    for spec in _build_default_rooms_spec():
        room = ExamRoom.objects.create(
            name=spec['name'],
            has_computer=spec['has_computer'],
            front_position=spec['front_position'],
            front_label=spec['front_label'],
            default_fill_pattern=spec['default_fill_pattern'],
            default_alternate=spec['default_alternate'],
            is_active=True,
            sort_order=spec['sort_order'],
            description=spec['description'],
        )
        reserves = spec['reserves']
        seats_to_create = []
        for block_idx, row_idx, label_pos, start_col, numbers in spec['rows']:
            for offset, num in enumerate(numbers):
                seats_to_create.append(
                    ExamSeat(
                        room=room,
                        seat_label=str(num),
                        seat_number=int(num),
                        block_index=block_idx,
                        row_index=row_idx,
                        col_index=start_col + offset,
                        label_position=label_pos,
                        default_status=ExamSeat.STATUS_RESERVE if num in reserves else ExamSeat.STATUS_NORMAL,
                        is_active=True,
                    )
                )
        ExamSeat.objects.bulk_create(seats_to_create)


def format_short_proctor_name(full_name):
    """ย่อชื่อวิทยากรคุมสอบเป็นรูปแบบ 'ครู<ชื่อ>' เช่น 'นายเสฏฐวุฒิ โมลานิล' -> 'ครูเสฏฐวุฒิ'"""
    s = (full_name or '').strip()
    if not s:
        return ''
    if s.startswith('ครู'):
        first_part = s.split()[0]
        return first_part
    prefixes = (
        'ว่าที่ร้อยตรี', 'ว่าที่ ร.ต.', 'ผศ.ดร.', 'รศ.ดร.', 'ศ.ดร.',
        'นางสาว', 'น.ส.', 'นาย', 'นาง', 'ดร.', 'ผศ.', 'รศ.', 'ศ.', 'อ.', 'อาจารย์',
    )
    cleaned = s
    for p in prefixes:
        if cleaned.startswith(p):
            cleaned = cleaned[len(p):].strip()
            break
    first_name = cleaned.split()[0] if cleaned else s.split()[0]
    if not first_name:
        return s
    return f'ครู{first_name}'


def get_exam_examinees(year):
    """คืนรายชื่อผู้มีสิทธิ์สอบในปีการศึกษา `year` (นักเรียนปัจจุบันก่อน ตามด้วยนักเรียนใช้สิทธิ์)"""
    year_int = int(year)
    qs, used_ids = score_members_for_year(year_int)
    use_records = {
        ur.user_id: ur.rights_code
        for ur in UseRecord.objects.filter(academic_year=year_int, is_used=True)
    }
    examinees = []
    for u in qs:
        is_rights = u.pk in used_ids
        display_code = (use_records.get(u.pk) if is_rights else None) or u.camp1_code
        examinees.append({
            'user': u,
            'user_id': u.pk,
            'display_code': display_code,
            'camp1_code': u.camp1_code,
            'full_name': u.get_full_name() or u.camp1_code,
            'school_name': u.school_name or '',
            'is_rights': is_rights,
            'group_label': 'ใช้สิทธิ์' if is_rights else 'ปัจจุบัน',
        })
    return examinees


def apply_default_reserves(plan_room, actor=None):
    """นำสถานะเครื่องสำรอง/งดใช้งานเริ่มต้นจากผังห้องมาใส่ใน ExamSeatingPlanRoom (ไม่ทับที่นั่งที่มีผู้สอบแล้ว)"""
    seats = list(
        plan_room.room.seats.filter(
            is_active=True,
            default_status__in=(ExamSeat.STATUS_RESERVE, ExamSeat.STATUS_DISABLED),
        )
    )
    if not seats:
        return 0
    existing_seat_ids = set(
        ExamSeatAssignment.objects.filter(plan_room=plan_room).values_list('seat_id', flat=True)
    )
    to_create = []
    for seat in seats:
        if seat.pk in existing_seat_ids:
            continue
        st = (
            ExamSeatAssignment.STATUS_RESERVE
            if seat.default_status == ExamSeat.STATUS_RESERVE
            else ExamSeatAssignment.STATUS_BLOCKED
        )
        to_create.append(
            ExamSeatAssignment(
                plan=plan_room.plan,
                plan_room=plan_room,
                seat=seat,
                status=st,
                user=None,
                assigned_by=actor if getattr(actor, 'is_authenticated', False) else None,
            )
        )
    if to_create:
        ExamSeatAssignment.objects.bulk_create(to_create, ignore_conflicts=True)
    return len(to_create)


@transaction.atomic
def sync_plan_rooms(plan, actor=None):
    """สร้าง ExamSeatingPlanRoom ให้ครบทุก ExamRoom ที่เปิดใช้งาน"""
    active_rooms = list(ExamRoom.objects.filter(is_active=True).order_by('sort_order', 'id'))
    existing_map = {pr.room_id: pr for pr in plan.plan_rooms.select_related('room').all()}
    result = []
    for room in active_rooms:
        pr = existing_map.get(room.pk)
        if pr is None:
            pr = ExamSeatingPlanRoom.objects.create(
                plan=plan,
                room=room,
                is_open=True,
                fill_pattern=room.default_fill_pattern,
                alternate_seats=room.default_alternate,
                sort_order=room.sort_order,
            )
            apply_default_reserves(pr, actor=actor)
        result.append(pr)
    return sorted(result, key=lambda x: (x.sort_order, x.room.sort_order, x.pk))


def get_or_create_plan_for_date(training_date, session_period=None, actor=None):
    """ดึงหรือสร้าง ExamSeatingPlan สำหรับวันสอบที่เลือก"""
    ensure_default_exam_rooms()
    if session_period is None:
        existing = ExamSeatingPlan.objects.filter(training_date=training_date).order_by('id').first()
        if existing:
            sync_plan_rooms(existing, actor=actor)
            return existing
        session_period = training_date.exam_period or ExamSeatingPlan.PERIOD_BOTH
    plan = ExamSeatingPlan.objects.filter(
        training_date=training_date, session_period=session_period,
    ).first()
    if plan is None:
        title_str = (training_date.title or '').strip()
        if training_date.exam_type in (
            ExamSeatingPlan.TYPE_THEORY, ExamSeatingPlan.TYPE_PRACTICE, ExamSeatingPlan.TYPE_GENERAL,
        ):
            default_type = training_date.exam_type
        elif 'ปฏิบัติ' in title_str or 'ปฎิบัติ' in title_str:
            default_type = ExamSeatingPlan.TYPE_PRACTICE
        else:
            default_type = ExamSeatingPlan.TYPE_THEORY
        plan = ExamSeatingPlan.objects.create(
            academic_year=training_date.academic_year,
            training_date=training_date,
            title=title_str,
            exam_type=default_type,
            session_period=session_period,
            is_open=False,
            created_by=actor if getattr(actor, 'is_authenticated', False) else None,
        )
    sync_plan_rooms(plan, actor=actor)
    return plan


def build_room_layout_data(room, assignment_by_seat_id=None, code_by_user_id=None, rights_user_ids=None):
    """สร้างโครงสร้างตารางสำหรับแสดงผังห้องสอบ (จัดกลุ่มตาม block_index -> row_index -> col_index)"""
    assignment_by_seat_id = assignment_by_seat_id or {}
    code_by_user_id = code_by_user_id or {}
    rights_user_ids = rights_user_ids or set()

    seats = list(room.seats.filter(is_active=True).order_by('block_index', 'row_index', 'col_index'))
    if not seats:
        return {
            'room': room,
            'max_col': 0,
            'blocks': [],
            'total_seats': 0,
            'assigned_count': 0,
            'reserve_count': 0,
            'blocked_count': 0,
            'free_count': 0,
        }

    max_col = max(s.col_index for s in seats)
    block_map = {}
    assigned_count = 0
    reserve_count = 0
    blocked_count = 0

    for seat in seats:
        row_map = block_map.setdefault(seat.block_index, {})
        col_map = row_map.setdefault(seat.row_index, {})
        asgn = assignment_by_seat_id.get(seat.pk)
        if asgn:
            status = asgn.status
            if status == ExamSeatAssignment.STATUS_ASSIGNED and asgn.user:
                assigned_count += 1
                u = asgn.user
                display_code = code_by_user_id.get(u.pk) or u.camp1_code
                full_name = u.get_full_name() or u.camp1_code
                school_name = u.school_name or ''
                is_rights = u.pk in rights_user_ids
                user_id = u.pk
            else:
                if status == ExamSeatAssignment.STATUS_RESERVE:
                    reserve_count += 1
                elif status == ExamSeatAssignment.STATUS_BLOCKED:
                    blocked_count += 1
                display_code = ''
                full_name = ''
                school_name = ''
                is_rights = False
                user_id = None
        else:
            status = 'empty'
            display_code = ''
            full_name = ''
            school_name = ''
            is_rights = False
            user_id = None

        col_map[seat.col_index] = {
            'seat': seat,
            'assignment': asgn,
            'status': status,
            'display_code': display_code,
            'full_name': full_name,
            'school_name': school_name,
            'is_rights': is_rights,
            'user_id': user_id,
        }

    blocks = []
    for block_idx in sorted(block_map.keys()):
        rows = []
        row_map = block_map[block_idx]
        for row_idx in sorted(row_map.keys()):
            col_map = row_map[row_idx]
            cells = [col_map.get(c) for c in range(1, max_col + 1)]
            non_empty = [cell for cell in cells if cell is not None]
            label_pos = non_empty[0]['seat'].label_position if non_empty else ExamSeat.LABEL_BELOW
            rows.append({
                'row_index': row_idx,
                'label_position': label_pos,
                'cells': cells,
            })
        blocks.append({
            'block_index': block_idx,
            'rows': rows,
        })

    total_seats = len(seats)
    free_count = max(0, total_seats - assigned_count - reserve_count - blocked_count)
    return {
        'room': room,
        'max_col': max_col,
        'blocks': blocks,
        'total_seats': total_seats,
        'assigned_count': assigned_count,
        'reserve_count': reserve_count,
        'blocked_count': blocked_count,
        'free_count': free_count,
    }


def order_room_seats(room, pattern, start_seat_id=None):
    """จัดเรียงลำดับที่นั่งในห้องตามรูปแบบและจุดเริ่มต้นที่เลือก"""
    seats = list(room.seats.filter(is_active=True))
    if not seats:
        return []

    start_seat = None
    if start_seat_id:
        start_seat = next((s for s in seats if s.pk == int(start_seat_id)), None)

    pat = pattern or room.default_fill_pattern or ExamRoom.FILL_SEQ_ASC

    if pat == ExamRoom.FILL_SEQ_ASC:
        ordered = sorted(seats, key=lambda s: (s.seat_number, s.block_index, s.row_index, s.col_index, s.seat_label))
        if start_seat:
            idx = ordered.index(start_seat)
            ordered = ordered[idx:]
        return ordered

    # จัดกลุ่มตามแถว
    rows_by_idx = {}
    for s in seats:
        rows_by_idx.setdefault(s.row_index, []).append(s)
    sorted_row_indices = sorted(rows_by_idx.keys())

    if start_seat and start_seat.row_index in rows_by_idx:
        start_row_pos = sorted_row_indices.index(start_seat.row_index)
        active_row_indices = sorted_row_indices[start_row_pos:]
    else:
        active_row_indices = sorted_row_indices

    if pat == ExamRoom.FILL_SNAKE_AUTO:
        if start_seat is None:
            pat = ExamRoom.FILL_SNAKE_LR
        else:
            row_cols = [s.col_index for s in rows_by_idx[start_seat.row_index]]
            mid_col = (min(row_cols) + max(row_cols)) / 2.0
            pat = ExamRoom.FILL_SNAKE_RL if start_seat.col_index > mid_col else ExamRoom.FILL_SNAKE_LR

    ordered = []
    for k, r_idx in enumerate(active_row_indices):
        row_seats = rows_by_idx[r_idx]
        if pat == ExamRoom.FILL_LR:
            dir_lr = True
        elif pat == ExamRoom.FILL_RL:
            dir_lr = False
        elif pat == ExamRoom.FILL_SNAKE_LR:
            dir_lr = (k % 2 == 0)
        elif pat == ExamRoom.FILL_SNAKE_RL:
            dir_lr = (k % 2 == 1)
        else:
            dir_lr = True

        row_sorted = sorted(row_seats, key=lambda s: s.col_index, reverse=(not dir_lr))
        if k == 0 and start_seat and start_seat.row_index == r_idx:
            if dir_lr:
                row_sorted = [s for s in row_sorted if s.col_index >= start_seat.col_index]
            else:
                row_sorted = [s for s in row_sorted if s.col_index <= start_seat.col_index]
        ordered.extend(row_sorted)
    return ordered


_UNSET = object()


def _select_assignable_seats(
    plan_room,
    pattern=None,
    start_seat_id=_UNSET,
    alternate=None,
    skip_adjacent_to_reserve=True,
    limit=None,
):
    """คำนวณรายชื่อที่นั่งที่จัดลงได้ตามลำดับและเงื่อนไขเว้นระยะ (อ่านอย่างเดียว)"""
    room = plan_room.room
    eff_pattern = pattern or plan_room.effective_fill_pattern
    eff_alternate = room.default_alternate if alternate is None else bool(alternate)
    eff_start_seat_id = plan_room.start_seat_id if start_seat_id is _UNSET else start_seat_id

    existing_assignments = list(
        ExamSeatAssignment.objects.filter(plan_room=plan_room).select_related('seat')
    )
    occupied_seat_ids = {a.seat_id for a in existing_assignments}
    blocked_neighbor_coords = set()
    for a in existing_assignments:
        if a.status == ExamSeatAssignment.STATUS_ASSIGNED:
            blocked_neighbor_coords.add((a.seat.row_index, a.seat.col_index))
        elif a.status == ExamSeatAssignment.STATUS_RESERVE and skip_adjacent_to_reserve:
            blocked_neighbor_coords.add((a.seat.row_index, a.seat.col_index))

    candidate_seats = order_room_seats(room, eff_pattern, start_seat_id=eff_start_seat_id)
    selected = []
    for seat in candidate_seats:
        if limit is not None and len(selected) >= limit:
            break
        if seat.pk in occupied_seat_ids:
            continue
        if eff_alternate:
            if (
                (seat.row_index, seat.col_index - 1) in blocked_neighbor_coords
                or (seat.row_index, seat.col_index + 1) in blocked_neighbor_coords
            ):
                continue
        selected.append(seat)
        occupied_seat_ids.add(seat.pk)
        blocked_neighbor_coords.add((seat.row_index, seat.col_index))
    return selected, eff_pattern, eff_alternate, eff_start_seat_id


def estimate_room_assignable_count(
    plan_room,
    pattern=None,
    start_seat_id=_UNSET,
    alternate=None,
    skip_adjacent_to_reserve=True,
):
    """ประมาณการจำนวนที่นั่งที่จัดลงได้ในห้องตามการตั้งค่าปัจจุบัน"""
    selected, _, _, _ = _select_assignable_seats(
        plan_room,
        pattern=pattern,
        start_seat_id=start_seat_id,
        alternate=alternate,
        skip_adjacent_to_reserve=skip_adjacent_to_reserve,
    )
    return len(selected)


@transaction.atomic
def auto_assign_room(
    plan_room,
    user_ids,
    pattern=None,
    start_seat_id=_UNSET,
    alternate=None,
    skip_adjacent_to_reserve=True,
    clear_existing_students=False,
    actor=None,
):
    """จัดผู้เข้าสอบลงในห้องสอบอัตโนมัติตามลำดับที่นั่งที่คำนวณได้"""
    if clear_existing_students:
        ExamSeatAssignment.objects.filter(
            plan_room=plan_room, status=ExamSeatAssignment.STATUS_ASSIGNED,
        ).delete()

    plan = plan_room.plan
    room = plan_room.room

    # ข้ามผู้เข้าสอบที่ถูกจัดที่นั่งในแผนสอบนี้แล้ว (รวมห้องอื่นในแผนเดียวกัน)
    already_seated_user_ids = set(
        ExamSeatAssignment.objects.filter(
            plan=plan, status=ExamSeatAssignment.STATUS_ASSIGNED, user__isnull=False,
        ).values_list('user_id', flat=True)
    )
    pending_user_ids = []
    seen_u = set()
    for uid in user_ids:
        uid_int = int(uid)
        if uid_int not in already_seated_user_ids and uid_int not in seen_u:
            seen_u.add(uid_int)
            pending_user_ids.append(uid_int)

    selected_seats, eff_pattern, eff_alternate, eff_start_seat_id = _select_assignable_seats(
        plan_room,
        pattern=pattern,
        start_seat_id=start_seat_id,
        alternate=alternate,
        skip_adjacent_to_reserve=skip_adjacent_to_reserve,
        limit=len(pending_user_ids),
    )

    actor_user = actor if getattr(actor, 'is_authenticated', False) else None
    to_create = [
        ExamSeatAssignment(
            plan=plan,
            plan_room=plan_room,
            seat=seat,
            status=ExamSeatAssignment.STATUS_ASSIGNED,
            user_id=uid,
            assigned_by=actor_user,
        )
        for uid, seat in zip(pending_user_ids, selected_seats)
    ]
    assigned_count = len(to_create)

    if to_create:
        ExamSeatAssignment.objects.bulk_create(to_create)

    # บันทึกค่า pattern/start_seat/alternate ล่าสุดของห้องในแผน
    plan_room.fill_pattern = eff_pattern
    plan_room.alternate_seats = eff_alternate
    if eff_start_seat_id:
        valid_start = room.seats.filter(pk=int(eff_start_seat_id), is_active=True).first()
        plan_room.start_seat = valid_start
    else:
        plan_room.start_seat = None
    plan_room.save(update_fields=['fill_pattern', 'alternate_seats', 'start_seat'])

    remaining_user_ids = pending_user_ids[assigned_count:]
    return assigned_count, remaining_user_ids


@transaction.atomic
def auto_assign_multi_rooms(
    plan,
    plan_room_ids,
    user_ids,
    alternate=True,
    skip_adjacent_to_reserve=True,
    actor=None,
):
    """จัดผู้เข้าสอบกระจายต่อเนื่องข้ามหลายห้องตามลำดับห้องที่เลือก"""
    remaining = [int(uid) for uid in user_ids]
    per_room_counts = []
    total_assigned = 0

    rooms_by_id = {
        pr.pk: pr
        for pr in ExamSeatingPlanRoom.objects.filter(plan=plan, pk__in=plan_room_ids).select_related('room', 'start_seat')
    }
    for pr_id in plan_room_ids:
        pr = rooms_by_id.get(int(pr_id))
        if not pr or not remaining:
            continue
        cnt, remaining = auto_assign_room(
            pr,
            remaining,
            pattern=pr.effective_fill_pattern,
            start_seat_id=pr.start_seat_id,
            alternate=alternate,
            skip_adjacent_to_reserve=skip_adjacent_to_reserve,
            clear_existing_students=False,
            actor=actor,
        )
        per_room_counts.append({'plan_room': pr, 'count': cnt})
        total_assigned += cnt

    return {
        'total_assigned': total_assigned,
        'per_room_counts': per_room_counts,
        'remaining_count': len(remaining),
    }


@transaction.atomic
def assign_single_seat(plan_room, seat, user, actor=None, overwrite_reserve=False):
    """จัดผู้เข้าสอบ 1 คนลงบนที่นั่งที่เลือก"""
    if seat.room_id != plan_room.room_id:
        raise SeatingError('ที่นั่งที่เลือกไม่อยู่ในห้องสอบนี้')
    if not seat.is_active:
        raise SeatingError('ที่นั่งนี้ถูกปิดใช้งาน')

    existing_user_asgn = ExamSeatAssignment.objects.filter(
        plan=plan_room.plan, user=user, status=ExamSeatAssignment.STATUS_ASSIGNED,
    ).select_related('plan_room__room', 'seat').first()
    if existing_user_asgn:
        raise SeatingError(
            f'{user.camp1_code} ถูกจัดที่นั่งแล้วใน {existing_user_asgn.plan_room.room.name} '
            f'ที่นั่ง {existing_user_asgn.seat.seat_label}'
        )

    existing_seat_asgn = ExamSeatAssignment.objects.filter(plan_room=plan_room, seat=seat).first()
    actor_user = actor if getattr(actor, 'is_authenticated', False) else None
    if existing_seat_asgn:
        if existing_seat_asgn.status == ExamSeatAssignment.STATUS_ASSIGNED:
            raise SeatingError(f'ที่นั่ง {seat.seat_label} มีผู้เข้าสอบนั่งอยู่แล้ว')
        if not overwrite_reserve:
            raise SeatingError(f'ที่นั่ง {seat.seat_label} ถูกตั้งสถานะ "{existing_seat_asgn.get_status_display()}" อยู่')
        existing_seat_asgn.status = ExamSeatAssignment.STATUS_ASSIGNED
        existing_seat_asgn.user = user
        existing_seat_asgn.assigned_by = actor_user
        existing_seat_asgn.save()
        return existing_seat_asgn

    return ExamSeatAssignment.objects.create(
        plan=plan_room.plan,
        plan_room=plan_room,
        seat=seat,
        status=ExamSeatAssignment.STATUS_ASSIGNED,
        user=user,
        assigned_by=actor_user,
    )


@transaction.atomic
def cancel_seat_assignment(assignment):
    """ยกเลิกการจัดที่นั่ง (หรือล้างสถานะสำรอง/งดใช้)"""
    assignment.delete()


@transaction.atomic
def move_or_swap_seat(source_assignment, target_plan_room, target_seat, actor=None):
    """ย้ายผู้เข้าสอบไปยังที่นั่งว่าง หรือสลับที่นั่งกับผู้เข้าสอบอีกคน (รองรับทั้งในห้องและข้ามห้อง)"""
    if source_assignment.status != ExamSeatAssignment.STATUS_ASSIGNED or not source_assignment.user:
        raise SeatingError('ที่นั่งต้นทางไม่มีผู้เข้าสอบให้ย้าย')
    if source_assignment.plan_id != target_plan_room.plan_id:
        raise SeatingError('ไม่สามารถย้ายข้ามวันสอบหรือรอบสอบได้')
    if target_seat.room_id != target_plan_room.room_id:
        raise SeatingError('ที่นั่งปลายทางไม่อยู่ในห้องสอบที่เลือก')
    if source_assignment.plan_room_id == target_plan_room.pk and source_assignment.seat_id == target_seat.pk:
        return 'same'

    actor_user = actor if getattr(actor, 'is_authenticated', False) else None
    target_asgn = ExamSeatAssignment.objects.filter(
        plan_room=target_plan_room, seat=target_seat,
    ).select_related('user').first()

    if target_asgn is None:
        source_assignment.plan_room = target_plan_room
        source_assignment.seat = target_seat
        source_assignment.assigned_by = actor_user
        source_assignment.save(update_fields=['plan_room', 'seat', 'assigned_by', 'assigned_at'])
        return 'moved'

    if target_asgn.status != ExamSeatAssignment.STATUS_ASSIGNED or not target_asgn.user:
        raise SeatingError(f'ที่นั่งปลายทาง ({target_seat.seat_label}) เป็น "{target_asgn.get_status_display()}" ไม่สามารถย้ายลงได้')

    # สลับผู้เข้าสอบโดยปลด unique constraint (plan, user) ชั่วคราวก่อนสลับ
    src_user = source_assignment.user
    tgt_user = target_asgn.user
    source_assignment.user = None
    source_assignment.save(update_fields=['user'])

    target_asgn.user = src_user
    target_asgn.assigned_by = actor_user
    target_asgn.save(update_fields=['user', 'assigned_by', 'assigned_at'])

    source_assignment.user = tgt_user
    source_assignment.assigned_by = actor_user
    source_assignment.save(update_fields=['user', 'assigned_by', 'assigned_at'])
    return 'swapped'


@transaction.atomic
def set_seat_status(plan_room, seat, status, actor=None):
    """ตั้งค่าสถานะที่นั่งในแผนสอบนี้ ('reserve', 'blocked', 'empty')"""
    if seat.room_id != plan_room.room_id:
        raise SeatingError('ที่นั่งที่เลือกไม่อยู่ในห้องสอบนี้')
    existing = ExamSeatAssignment.objects.filter(plan_room=plan_room, seat=seat).first()
    if existing and existing.status == ExamSeatAssignment.STATUS_ASSIGNED:
        raise SeatingError('ที่นั่งนี้มีผู้เข้าสอบอยู่ กรุณายกเลิกหรือย้ายผู้เข้าสอบก่อน')

    if status == 'empty':
        if existing:
            existing.delete()
        return None

    if status not in (ExamSeatAssignment.STATUS_RESERVE, ExamSeatAssignment.STATUS_BLOCKED):
        raise SeatingError('สถานะที่นั่งไม่ถูกต้อง')

    actor_user = actor if getattr(actor, 'is_authenticated', False) else None
    if existing:
        existing.status = status
        existing.user = None
        existing.assigned_by = actor_user
        existing.save()
        return existing

    return ExamSeatAssignment.objects.create(
        plan=plan_room.plan,
        plan_room=plan_room,
        seat=seat,
        status=status,
        user=None,
        assigned_by=actor_user,
    )


# =============================================================================
# ตัวช่วยสร้าง / แก้ไขผังที่นั่งภายในห้อง (Room & Seat Layout Builders)
# =============================================================================
def _parse_seat_number(label, fallback=0):
    m = re.search(r'\d+', str(label or ''))
    return int(m.group(0)) if m else fallback


@transaction.atomic
def generate_desk_grid_seats(room, rows, cols, start_no=1, direction='lr', clear_existing=True):
    """สร้างที่นั่งสอบแบบตารางแถว × คอลัมน์ (เช่น ห้องประชุม 14 แถว × 5 โต๊ะ)"""
    rows = max(1, min(40, int(rows)))
    cols = max(1, min(30, int(cols)))
    num = int(start_no)
    if clear_existing:
        room.seats.all().delete()

    existing_labels = set(room.seats.values_list('seat_label', flat=True))
    max_row = max(room.seats.values_list('row_index', flat=True), default=0)
    max_blk = max(room.seats.values_list('block_index', flat=True), default=0)

    to_create = []
    for r in range(1, rows + 1):
        col_seq = list(range(1, cols + 1))
        if direction == 'rl':
            col_seq.reverse()
        elif direction == 'snake_lr' and (r % 2 == 0):
            col_seq.reverse()
        elif direction == 'snake_rl' and (r % 2 == 1):
            col_seq.reverse()

        for c in col_seq:
            label = str(num)
            while label in existing_labels:
                num += 1
                label = str(num)
            existing_labels.add(label)
            to_create.append(
                ExamSeat(
                    room=room,
                    seat_label=label,
                    seat_number=num,
                    block_index=max_blk + r,
                    row_index=max_row + r,
                    col_index=c,
                    label_position=ExamSeat.LABEL_BELOW,
                    default_status=ExamSeat.STATUS_NORMAL,
                    is_active=True,
                )
            )
            num += 1
    ExamSeat.objects.bulk_create(to_create)
    return len(to_create)


@transaction.atomic
def add_bench_block(
    room,
    bench_type='single',
    count=10,
    col_offset=0,
    start_no=1,
    dir_top='lr',
    dir_bottom='rl',
    label_pos='below',
    bottom_count=None,
    bottom_start_no=None,
):
    """เพิ่มแถวเครื่องคอมพิวเตอร์ 1 แถว (single) หรือแถวคู่หันหลังชนกัน (double)"""
    count = max(1, min(30, int(count)))
    col_offset = max(0, min(20, int(col_offset)))
    existing_labels = set(room.seats.values_list('seat_label', flat=True))
    next_block = max(room.seats.values_list('block_index', flat=True), default=0) + 1
    next_row = max(room.seats.values_list('row_index', flat=True), default=0) + 1

    to_create = []

    def _build_row(row_idx, cnt, s_no, direction, lpos):
        nums = list(range(int(s_no), int(s_no) + cnt))
        if direction == 'rl':
            nums.reverse()
        for offset, n in enumerate(nums):
            lbl = str(n)
            if lbl in existing_labels:
                raise SeatingError(f'เลขที่นั่ง/เครื่อง "{lbl}" มีอยู่ในห้องนี้แล้ว')
            existing_labels.add(lbl)
            to_create.append(
                ExamSeat(
                    room=room,
                    seat_label=lbl,
                    seat_number=n,
                    block_index=next_block,
                    row_index=row_idx,
                    col_index=1 + col_offset + offset,
                    label_position=lpos,
                    default_status=ExamSeat.STATUS_NORMAL,
                    is_active=True,
                )
            )

    if bench_type == 'double':
        b_cnt = max(1, min(30, int(bottom_count or count)))
        b_start = int(bottom_start_no) if bottom_start_no is not None else (int(start_no) + count)
        _build_row(next_row, count, start_no, dir_top, ExamSeat.LABEL_ABOVE)
        _build_row(next_row + 1, b_cnt, b_start, dir_bottom, ExamSeat.LABEL_BELOW)
    else:
        lpos = ExamSeat.LABEL_ABOVE if label_pos == ExamSeat.LABEL_ABOVE else ExamSeat.LABEL_BELOW
        _build_row(next_row, count, start_no, dir_top, lpos)

    ExamSeat.objects.bulk_create(to_create)
    return len(to_create)


@transaction.atomic
def add_single_seat(
    room,
    block_index,
    row_index,
    col_index,
    seat_label,
    seat_number=None,
    label_position='below',
    default_status='normal',
):
    """เพิ่มที่นั่ง/เครื่อง 1 ตำแหน่งในห้องสอบ"""
    label = str(seat_label or '').strip()
    if not label:
        raise SeatingError('กรุณาระบุเลขที่นั่ง / เลขเครื่อง')
    blk = max(1, int(block_index))
    row = max(1, int(row_index))
    col = max(1, int(col_index))

    if room.seats.filter(seat_label=label).exists():
        raise SeatingError(f'เลขที่นั่ง/เครื่อง "{label}" มีอยู่ในห้องนี้แล้ว')
    if room.seats.filter(row_index=row, col_index=col).exists():
        raise SeatingError(f'ตำแหน่ง แถว {row} คอลัมน์ {col} มีที่นั่งอยู่แล้ว')

    num = int(seat_number) if (seat_number is not None and str(seat_number).strip() != '') else _parse_seat_number(label, 0)
    lpos = ExamSeat.LABEL_ABOVE if label_position == ExamSeat.LABEL_ABOVE else ExamSeat.LABEL_BELOW
    dst = default_status if default_status in dict(ExamSeat.DEFAULT_STATUS_CHOICES) else ExamSeat.STATUS_NORMAL

    return ExamSeat.objects.create(
        room=room,
        seat_label=label,
        seat_number=num,
        block_index=blk,
        row_index=row,
        col_index=col,
        label_position=lpos,
        default_status=dst,
        is_active=True,
    )


@transaction.atomic
def update_seat_meta(seat, seat_label, seat_number=None, default_status='normal', label_position=None):
    """แก้ไขข้อมูลที่นั่ง/เครื่อง"""
    label = str(seat_label or '').strip()
    if not label:
        raise SeatingError('เลขที่นั่ง/เครื่องห้ามว่าง')
    if seat.room.seats.exclude(pk=seat.pk).filter(seat_label=label).exists():
        raise SeatingError(f'เลขที่นั่ง/เครื่อง "{label}" ซ้ำกับตำแหน่งอื่นในห้อง')
    seat.seat_label = label
    if seat_number is not None and str(seat_number).strip() != '':
        seat.seat_number = int(seat_number)
    else:
        seat.seat_number = _parse_seat_number(label, seat.seat_number)
    if default_status in dict(ExamSeat.DEFAULT_STATUS_CHOICES):
        seat.default_status = default_status
    if label_position in (ExamSeat.LABEL_BELOW, ExamSeat.LABEL_ABOVE):
        seat.label_position = label_position
    seat.save()
    return seat


@transaction.atomic
def delete_seat(seat):
    seat.delete()


@transaction.atomic
def delete_room_block(room, block_index):
    return room.seats.filter(block_index=int(block_index)).delete()


@transaction.atomic
def save_room_proctors(plan_room, recipient_ids=None, name_style='short', custom_names=None):
    """บันทึกวิทยากรคุมสอบประจำห้องในแผนสอบ"""
    plan_room.proctors.all().delete()
    to_create = []
    order = 1

    rec_ids = [int(r) for r in (recipient_ids or []) if str(r).strip().isdigit()]
    if rec_ids:
        rec_map = {r.pk: r for r in CertificateRecipient.objects.filter(pk__in=rec_ids)}
        for rid in rec_ids:
            rec = rec_map.get(rid)
            if not rec:
                continue
            disp = format_short_proctor_name(rec.full_name) if name_style == 'short' else rec.full_name.strip()
            to_create.append(
                ExamRoomProctor(
                    plan_room=plan_room,
                    recipient=rec,
                    display_name=disp,
                    sort_order=order,
                )
            )
            order += 1

    if custom_names:
        parts = [p.strip() for p in re.split(r'[\n,;]+', str(custom_names)) if p.strip()]
        for p in parts:
            to_create.append(
                ExamRoomProctor(
                    plan_room=plan_room,
                    recipient=None,
                    display_name=p[:200],
                    sort_order=order,
                )
            )
            order += 1

    if to_create:
        ExamRoomProctor.objects.bulk_create(to_create)
    return to_create


def check_plan_ready_to_open(plan, examinees=None):
    """ตรวจสอบว่าวันสอบนี้จัดที่นั่งครบและพร้อมเปิดให้ผู้เข้าสอบดูผังแล้วหรือไม่"""
    if plan is None:
        reason = 'ไม่พบแผนการจัดที่นั่งสอบ'
        return {
            'is_ready': False,
            'total_examinees': 0,
            'seated_count': 0,
            'unseated_count': 0,
            'closed_room_seated_count': 0,
            'open_rooms_count': 0,
            'reasons': [reason],
            'reason_text': reason,
        }

    if examinees is None:
        examinees = get_exam_examinees(plan.academic_year)

    examinee_ids = {ex['user_id'] for ex in examinees}
    total_examinees = len(examinee_ids)

    open_rooms_count = ExamSeatingPlanRoom.objects.filter(plan=plan, is_open=True).count()
    assigned_qs = ExamSeatAssignment.objects.filter(
        plan=plan,
        status=ExamSeatAssignment.STATUS_ASSIGNED,
        user__isnull=False,
    ).select_related('plan_room')

    open_assigned_ids = set()
    closed_room_seated_count = 0
    for assign in assigned_qs:
        if assign.plan_room.is_open:
            if assign.user_id in examinee_ids:
                open_assigned_ids.add(assign.user_id)
        else:
            if assign.user_id in examinee_ids:
                closed_room_seated_count += 1

    seated_count = len(open_assigned_ids)
    unseated_count = max(0, total_examinees - seated_count)

    reasons = []
    if total_examinees == 0:
        reasons.append('ยังไม่มีรายชื่อผู้มีสิทธิ์สอบในปีการศึกษานี้')
    else:
        if open_rooms_count == 0:
            reasons.append('ยังไม่ได้เปิดใช้ห้องสอบในวันสอบนี้')
        if closed_room_seated_count > 0:
            reasons.append(f'มีผู้เข้าสอบ {closed_room_seated_count} คนอยู่ในห้องที่ปิดอยู่')
        if unseated_count > 0:
            reasons.append(f'ยังจัดที่นั่งไม่ครบ (ขาดอีก {unseated_count}/{total_examinees} คน)')

    is_ready = bool(
        total_examinees > 0
        and open_rooms_count > 0
        and closed_room_seated_count == 0
        and unseated_count == 0
    )
    return {
        'is_ready': is_ready,
        'total_examinees': total_examinees,
        'seated_count': seated_count,
        'unseated_count': unseated_count,
        'closed_room_seated_count': closed_room_seated_count,
        'open_rooms_count': open_rooms_count,
        'reasons': reasons,
        'reason_text': ' · '.join(reasons),
    }


def auto_close_plan_if_incomplete(plan, examinees=None):
    """หากแผนวันสอบเปิดให้ดูอยู่แต่การจัดห้องยังไม่เรียบร้อย ให้ปิดการแสดงผลอัตโนมัติ"""
    readiness = check_plan_ready_to_open(plan, examinees=examinees)
    if plan is not None and plan.is_open and not readiness['is_ready']:
        plan.is_open = False
        plan.save(update_fields=['is_open', 'updated_at'])
        return True, readiness
    return False, readiness
