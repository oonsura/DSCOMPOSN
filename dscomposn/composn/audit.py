"""
Audit Log helper — บันทึกการกระทำสำคัญลงฐานข้อมูล และไฟล์ใน audit_logs/
"""
import json
import logging

from django.conf import settings
from django.utils import timezone

logger = logging.getLogger(__name__)

ACTION_LABELS = {
    'login': 'เข้าสู่ระบบ',
    'logout': 'ออกจากระบบ',
    'login_failed': 'เข้าสู่ระบบไม่สำเร็จ',
    'login_blocked': 'ระงับการเข้าสู่ระบบชั่วคราว (ลองผิดเกินกำหนด)',
    'login_unlock': 'Admin ปลดล็อกการเข้าสู่ระบบ',
    'session_kicked': 'ถูกออกจากระบบ (Login จากเครื่องอื่น)',
    'change_password': 'เปลี่ยนรหัสผ่าน',
    'reset_password': 'Reset รหัสผ่าน',
    'set_password': 'Admin ตั้งรหัสผ่าน',
    'create_user': 'เพิ่มผู้ใช้',
    'update_user': 'แก้ไขข้อมูลผู้ใช้',
    'delete_user': 'ลบผู้ใช้',
    'toggle_user': 'เปิด/ปิดการใช้งานผู้ใช้',
    'update_profile': 'แก้ไขโปรไฟล์ตนเอง',
    'import_students': 'Import ข้อมูลนักเรียน',
    'import_camp2': 'Import รหัสค่าย 2',
    'delete_camp2': 'ลบรหัสค่าย 2',
    'calculate_level': 'คำนวณระดับชั้น',
    'confirm_rights': 'ยืนยันใช้สิทธิ์',
    'cancel_rights': 'ยกเลิกการใช้สิทธิ์',
    'import_rights': 'นำเข้าผู้ใช้สิทธิ์ (ค่าย 1)',
    'waive_rights': 'สละสิทธิ์ (ไม่ยืนยันในรอบ)',
    'open_round': 'เปิดรอบใช้สิทธิ์',
    'close_round': 'ปิดรอบใช้สิทธิ์',
    'delete_use_record': 'ลบบันทึกการใช้สิทธิ์',
    'restore_record': 'กู้คืนข้อมูล',
    'school_request': 'ส่งคำขอแก้ไขชื่อโรงเรียน',
    'school_approve': 'อนุมัติชื่อโรงเรียน',
    'school_reject': 'ไม่อนุมัติชื่อโรงเรียน',
    'school_clear': 'เคลียร์สถานะคำขอโรงเรียน',
    'update_settings': 'แก้ไขตั้งค่าระบบ',
    'save_attendance': 'บันทึกเวลาเรียน',
    'delete_attendance': 'ลบบันทึกเวลาเรียน',
    'save_score': 'บันทึกคะแนน',
    'delete_score': 'ลบคะแนน',
    'score_category': 'จัดการหมวดคะแนน',
    'backup': 'สำรองข้อมูล',
    'delete_backup': 'ลบไฟล์สำรอง',
    'export': 'Export ข้อมูล',
    'food_order': 'จัดการรายการอาหาร',
    'food_select': 'เลือกอาหาร',
    'toggle_cert_publish': 'เปิด/ปิดการแสดงเกียรติบัตร',
}


def get_client_ip(request):
    if request is None:
        return ''
    xff = request.META.get('HTTP_X_FORWARDED_FOR')
    if xff:
        return xff.split(',')[0].strip()
    return request.META.get('REMOTE_ADDR', '') or ''


def _to_text(value):
    if value is None or value == '':
        return ''
    if isinstance(value, str):
        return value
    try:
        return json.dumps(value, ensure_ascii=False, default=str)
    except (TypeError, ValueError):
        return str(value)


def log_action(user, action, target=None, target_model='', target_id=None,
               old_value=None, new_value=None, request=None):
    """บันทึก Audit Log (ห้ามทำให้งานหลักล้ม)"""
    from .models import AuditLog

    if target is not None:
        target_model = target_model or target.__class__.__name__
        target_id = target_id if target_id is not None else getattr(target, 'pk', None)
    if user is not None and not getattr(user, 'is_authenticated', False):
        user = None
    try:
        entry = AuditLog.objects.create(
            user=user, action=action, target_model=target_model or '',
            target_id=target_id, old_value=_to_text(old_value), new_value=_to_text(new_value),
            ip_address=get_client_ip(request),
        )
    except Exception:  # pragma: no cover
        logger.exception('ไม่สามารถบันทึก AuditLog')
        return None
    _write_file(entry)
    return entry


def _write_file(entry):
    try:
        log_dir = settings.AUDIT_LOG_DIR
        log_dir.mkdir(parents=True, exist_ok=True)
        now = timezone.localtime(entry.created_at)
        path = log_dir / f'audit_{now:%Y-%m-%d}.log'
        line = '\t'.join([
            now.strftime('%Y-%m-%d %H:%M:%S'),
            entry.user.camp1_code if entry.user else '-',
            entry.action,
            f'{entry.target_model}#{entry.target_id or ""}',
            entry.ip_address or '-',
            (entry.old_value or '').replace('\n', ' ')[:500],
            '=>',
            (entry.new_value or '').replace('\n', ' ')[:500],
        ])
        with open(path, 'a', encoding='utf-8') as fh:
            fh.write(line + '\n')
    except Exception:  # pragma: no cover
        logger.exception('ไม่สามารถเขียนไฟล์ audit log')


def model_snapshot(obj, fields):
    data = {}
    for f in fields:
        v = getattr(obj, f, None)
        data[f] = v.isoformat() if hasattr(v, 'isoformat') else v
    return data


def diff_snapshot(old, new):
    old_d = {k: v for k, v in old.items() if new.get(k) != v}
    new_d = {k: new[k] for k in old_d}
    return old_d, new_d
