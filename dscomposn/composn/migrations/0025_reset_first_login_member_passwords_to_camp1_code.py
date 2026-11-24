from django.contrib.auth.hashers import check_password, make_password
from django.db import migrations


def fix_first_login_member_passwords(apps, schema_editor):
    """แก้ไขรหัสผ่านเริ่มต้นของนักเรียนที่ยังไม่เคยเข้าสู่ระบบ ให้ตรงกับรหัสค่าย 1 (camp1_code)"""
    User = apps.get_model('composn', 'User')
    candidates = list(User.objects.filter(role='member', is_first_login=True))
    to_update = []
    for u in candidates:
        code = (u.camp1_code or '').strip().upper()
        if not code:
            continue
        if not check_password(code, u.password):
            u.password = make_password(code)
            to_update.append(u)
    if to_update:
        try:
            with schema_editor.connection.cursor() as cur:
                cur.execute('PRAGMA busy_timeout = 30000')
        except Exception:
            pass
        User.objects.bulk_update(to_update, ['password'], batch_size=50)


class Migration(migrations.Migration):
    atomic = False

    dependencies = [
        ('composn', '0024_systemsetting_academic_years'),
    ]

    operations = [
        migrations.RunPython(fix_first_login_member_passwords, migrations.RunPython.noop),
    ]
