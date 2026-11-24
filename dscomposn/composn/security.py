"""
ระบบป้องกันการโจมตีแบบเดารหัสผ่านซ้ำ (Brute-Force / Rate-Limit Lockout)
อาศัยข้อมูลจาก AuditLog ร่วมกับการตั้งค่าใน settings (ไม่ต้องเพิ่มตาราง Migration)
"""
import datetime as dt
import json
import math

from django.conf import settings
from django.utils import timezone

from .audit import log_action


def _cfg(name, default):
    try:
        return max(1, int(getattr(settings, name, default)))
    except (TypeError, ValueError):
        return default


def normalize_username(username):
    return (username or '').strip().upper()[:30]


def _parse_json(raw):
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except (TypeError, ValueError):
        return {}


def _scan_recent_auth_events(now=None):
    from .models import AuditLog

    now = now or timezone.now()
    window_mins = _cfg('LOGIN_LOCKOUT_WINDOW_MINUTES', 15)
    lock_mins = _cfg('LOGIN_LOCKOUT_MINUTES', 15)
    cutoff = now - dt.timedelta(minutes=max(window_mins, lock_mins))

    events = (
        AuditLog.objects
        .filter(action__in=('login_failed', 'login', 'login_unlock'), created_at__gte=cutoff)
        .select_related('user')
        .order_by('created_at')
    )

    user_fails = {}
    ip_fails = {}

    for ev in events:
        if ev.action == 'login':
            if ev.user and ev.user.camp1_code:
                user_fails.pop(normalize_username(ev.user.camp1_code), None)
            if ev.user and ev.user.camp2_code:
                user_fails.pop(normalize_username(ev.user.camp2_code), None)
        elif ev.action == 'login_unlock':
            data = _parse_json(ev.new_value)
            ltype = data.get('lock_type')
            lval = data.get('lock_value') or ''
            if ltype == 'username':
                user_fails.pop(normalize_username(lval), None)
            elif ltype == 'ip':
                ip_fails.pop(lval.strip(), None)
        elif ev.action == 'login_failed':
            data = _parse_json(ev.new_value)
            uname = normalize_username(data.get('username'))
            if uname:
                user_fails.setdefault(uname, []).append(ev.created_at)
            ip = (ev.ip_address or '').strip()
            if ip:
                ip_fails.setdefault(ip, []).append(ev.created_at)

    return now, dt.timedelta(minutes=window_mins), dt.timedelta(minutes=lock_mins), user_fails, ip_fails


def _eval_bucket(times, now, window_td, lock_td, threshold):
    if not times or len(times) < threshold:
        return None
    # นับความพยายามที่ผิดในช่วงหน้าต่างเวลาก่อนหน้าความพยายามล่าสุด
    last_ts = times[-1]
    unlock_at = last_ts + lock_td
    if unlock_at <= now:
        return None
    in_window = [t for t in times if t >= (last_ts - window_td)]
    if len(in_window) < threshold:
        return None
    rem_sec = max(1, int((unlock_at - now).total_seconds()))
    return {
        'fail_count': len(in_window),
        'last_failed_at': last_ts,
        'unlock_at': unlock_at,
        'remaining_seconds': rem_sec,
        'remaining_minutes': max(1, math.ceil(rem_sec / 60)),
    }


def check_login_lockout(username, ip_address='', now=None):
    """ตรวจสอบว่า username หรือ ip_address กำลังถูกล็อกชั่วคราวจากการล็อกอินผิดเกินกำหนดหรือไม่"""
    uname = normalize_username(username)
    ip = (ip_address or '').strip()
    max_user = _cfg('LOGIN_MAX_FAILURES_PER_USER', 5)
    max_ip = _cfg('LOGIN_MAX_FAILURES_PER_IP', 15)

    now, window_td, lock_td, user_fails, ip_fails = _scan_recent_auth_events(now=now)

    if uname:
        u_res = _eval_bucket(user_fails.get(uname, []), now, window_td, lock_td, max_user)
        if u_res:
            return {'locked': True, 'reason': 'username', 'target': uname, **u_res}

    if ip:
        i_res = _eval_bucket(ip_fails.get(ip, []), now, window_td, lock_td, max_ip)
        if i_res:
            return {'locked': True, 'reason': 'ip', 'target': ip, **i_res}

    return {'locked': False, 'reason': '', 'target': '', 'fail_count': 0, 'remaining_minutes': 0, 'remaining_seconds': 0}


def get_active_lockouts(now=None):
    """คืนรายการบัญชีและ IP ที่กำลังถูกล็อกชั่วคราวในขณะนี้ (สำหรับ Admin ตรวจสอบและปลดล็อก)"""
    max_user = _cfg('LOGIN_MAX_FAILURES_PER_USER', 5)
    max_ip = _cfg('LOGIN_MAX_FAILURES_PER_IP', 15)
    now, window_td, lock_td, user_fails, ip_fails = _scan_recent_auth_events(now=now)

    items = []
    for uname, times in user_fails.items():
        res = _eval_bucket(times, now, window_td, lock_td, max_user)
        if res:
            items.append({'lock_type': 'username', 'lock_value': uname, **res})
    for ip, times in ip_fails.items():
        res = _eval_bucket(times, now, window_td, lock_td, max_ip)
        if res:
            items.append({'lock_type': 'ip', 'lock_value': ip, **res})

    items.sort(key=lambda x: x['last_failed_at'], reverse=True)
    return items


def should_log_login_blocked(username, ip_address='', now=None):
    """ป้องกันการบันทึก AuditLog ซ้ำถี่เกินไปเมื่อผู้โจมตียิงคำขอขณะถูกล็อก"""
    from .models import AuditLog

    now = now or timezone.now()
    uname = normalize_username(username)
    ip = (ip_address or '').strip()
    recent = AuditLog.objects.filter(
        action='login_blocked',
        created_at__gte=now - dt.timedelta(seconds=60),
    )
    if ip:
        recent = recent.filter(ip_address=ip)
    for ev in recent[:10]:
        if normalize_username(_parse_json(ev.new_value).get('username')) == uname:
            return False
    return True


def unlock_login(lock_type, lock_value, by=None, request=None):
    """Admin ปลดล็อกบัญชีหรือ IP ที่ถูกระงับชั่วคราว"""
    val = normalize_username(lock_value) if lock_type == 'username' else (lock_value or '').strip()
    if lock_type not in ('username', 'ip') or not val:
        return None
    return log_action(
        by,
        'login_unlock',
        target_model='LoginLock',
        new_value={'lock_type': lock_type, 'lock_value': val},
        request=request,
    )
