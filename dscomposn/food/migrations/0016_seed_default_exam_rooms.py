from django.db import migrations


def _build_default_rooms_spec():
    return [
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 134',
            'has_computer': True,
            'front_position': 'right',
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': 'seq_asc',
            'default_alternate': True,
            'sort_order': 1,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 134 (50 เครื่อง)',
            'reserves': {1, 16, 41, 49},
            'rows': [
                (1, 1, 'below', 1, [15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]),
                (2, 2, 'above', 2, [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27]),
                (2, 3, 'below', 2, [39, 38, 37, 36, 35, 34, 33, 32, 31, 30, 29, 28]),
                (3, 4, 'above', 3, [40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50]),
            ],
        },
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 135',
            'has_computer': True,
            'front_position': 'left',
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': 'seq_asc',
            'default_alternate': True,
            'sort_order': 2,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 135 (55 เครื่อง)',
            'reserves': {16, 22, 30, 54},
            'rows': [
                (1, 1, 'below', 1, list(range(1, 16))),
                (2, 2, 'above', 1, list(range(29, 15, -1))),
                (2, 3, 'below', 1, list(range(30, 44))),
                (3, 4, 'above', 1, list(range(55, 43, -1))),
            ],
        },
        {
            'name': 'ห้องปฏิบัติการคอมพิวเตอร์ 136',
            'has_computer': True,
            'front_position': 'left',
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': 'seq_asc',
            'default_alternate': True,
            'sort_order': 3,
            'description': 'ห้องปฏิบัติการคอมพิวเตอร์ 136 (55 เครื่อง)',
            'reserves': {11, 30, 40, 42, 50, 55, 57},
            'rows': [
                (1, 1, 'below', 1, list(range(1, 9))),
                (2, 2, 'above', 1, list(range(20, 10, -1))),
                (2, 3, 'below', 1, list(range(21, 31))),
                (3, 4, 'above', 1, list(range(40, 30, -1))),
                (3, 5, 'below', 1, list(range(41, 51))),
                (4, 6, 'above', 4, list(range(57, 50, -1))),
            ],
        },
        {
            'name': 'ห้องประชุมเทิดพระเกียรติ ชั้น 2',
            'has_computer': False,
            'front_position': 'top',
            'front_label': 'หน้าเวที',
            'default_fill_pattern': 'snake_auto',
            'default_alternate': False,
            'sort_order': 4,
            'description': 'ห้องประชุมเทิดพระเกียรติ ชั้น 2 (โต๊ะสอบ 14 แถว 70 ที่นั่ง ไม่มีคอมพิวเตอร์)',
            'reserves': set(),
            'rows': [
                (r, r, 'below', 1, [(r - 1) * 5 + c for c in range(1, 6)])
                for r in range(1, 15)
            ],
        },
    ]


def seed_default_exam_rooms(apps, schema_editor):
    ExamRoom = apps.get_model('food', 'ExamRoom')
    ExamSeat = apps.get_model('food', 'ExamSeat')

    for spec in _build_default_rooms_spec():
        if ExamRoom.objects.filter(name=spec['name']).exists():
            continue
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
        seats_to_create = []
        reserves = spec['reserves']
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
                        default_status='reserve' if num in reserves else 'normal',
                        is_active=True,
                    )
                )
        ExamSeat.objects.bulk_create(seats_to_create)


class Migration(migrations.Migration):

    dependencies = [
        ('food', '0015_examroom_examseat_examseatingplan_and_more'),
    ]

    operations = [
        migrations.RunPython(seed_default_exam_rooms, migrations.RunPython.noop),
    ]
