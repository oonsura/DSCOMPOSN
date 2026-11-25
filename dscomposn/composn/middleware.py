"""
Middleware ความปลอดภัย
- SingleSessionMiddleware: 1 Account ใช้ได้เครื่องเดียว หาก Login ซ้ำให้เครื่องเก่าหลุด
- ForcePasswordChangeMiddleware: บังคับเปลี่ยนรหัสผ่านครั้งแรก
"""
from django.contrib import messages
from django.contrib.auth import logout
from django.shortcuts import redirect
from django.urls import reverse


class SingleSessionMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user = getattr(request, 'user', None)
        if user is not None and user.is_authenticated:
            key = request.session.session_key
            if user.current_session_key and key and user.current_session_key != key:
                from .audit import log_action
                log_action(user, 'session_kicked', target=user, request=request)
                logout(request)
                messages.warning(request, 'บัญชีนี้ถูกเข้าสู่ระบบจากอุปกรณ์อื่น ระบบจึงออกจากระบบเครื่องนี้')
                if request.headers.get('x-requested-with') == 'XMLHttpRequest':
                    from django.http import JsonResponse
                    return JsonResponse({'ok': False, 'error': 'session_kicked'}, status=401)
                return redirect('composn:login')

            # บังคับใช้เวลาตัดการเชื่อมต่ออัตโนมัติ (auto_logout_minutes) จากการตั้งค่าระบบ
            try:
                import time
                from .models import SystemSetting
                setting = getattr(request, '_site_setting', None) or SystemSetting.get()
                mins = max(0, int(getattr(setting, 'auto_logout_minutes', 5)))
                is_ping = (request.path == reverse('composn:session_ping'))
                ping_active = (is_ping and request.GET.get('active') == '1')
                now_ts = time.time()
                if mins > 0:
                    timeout_sec = mins * 60
                    last_ts = request.session.get('_last_activity')
                    if last_ts is not None and (now_ts - float(last_ts)) > timeout_sec:
                        if user.current_session_key == key:
                            from .models import User
                            User.objects.filter(pk=user.pk, current_session_key=key).update(current_session_key='')
                        logout(request)
                        if request.headers.get('x-requested-with') == 'XMLHttpRequest':
                            from django.http import JsonResponse
                            return JsonResponse({'ok': False, 'error': 'session_timeout'}, status=401)
                        return redirect('/login/?timeout=1')
                    if (not is_ping or ping_active) and (last_ts is None or (now_ts - float(last_ts)) >= 25):
                        request.session['_last_activity'] = now_ts
                        request.session.set_expiry(max(timeout_sec + 60, 300))
                elif not is_ping or ping_active:
                    last_ts = request.session.get('_last_activity')
                    if last_ts is None or (now_ts - float(last_ts)) >= 60:
                        request.session['_last_activity'] = now_ts
                        request.session.set_expiry(86400)
            except Exception:
                pass
        return self.get_response(request)


class ForcePasswordChangeMiddleware:
    ALLOWED_PREFIXES = ('/static/', '/media/', '/certificate/verify/')

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user = getattr(request, 'user', None)
        if user is not None and user.is_authenticated and user.is_first_login:
            allowed = {reverse('composn:change_password'), reverse('composn:logout'),
                       reverse('composn:login'), reverse('composn:session_ping')}
            if request.path not in allowed and not request.path.startswith(self.ALLOWED_PREFIXES):
                return redirect('composn:change_password')
        return self.get_response(request)


class RememberPageStateMiddleware:
    """
    จดจำตัวเลือก/ตัวกรองล่าสุดของแต่ละหน้า (query string เช่น ปีการศึกษา หัวข้อคะแนน ครั้งที่สอบ คำค้น หน้า)
    เมื่อออกไปหน้าอื่นแล้วกลับมาที่หน้าเดิม (เปิดจากเมนูโดยไม่มี query) ระบบจะพากลับไปที่ตัวเลือกเดิมให้อัตโนมัติ
    - จดจำเฉพาะหน้า HTML ที่แสดงผลสำเร็จ (200) ไม่จดจำการ Export/ดาวน์โหลดไฟล์
    - กดปุ่ม "ล้าง" (ลิงก์ที่มี ?reset=1) จะลืมค่าที่จดจำและเริ่มต้นใหม่
    - ค่าที่จดจำมีอายุภายในวันเดียวกัน (ป้องกันการเปิดค้างวันที่เก่า เช่น หน้าเช็คชื่อ)
    """
    SESSION_KEY = '_page_state'
    SKIP_PREFIXES = ('/static/', '/media/', '/login/', '/logout/', '/session-ping/', '/change-password/',
                     '/display/', '/api/', '/django-admin/', '/certificate/verify/')
    VOLATILE_PARAMS = {'export', 'timeout', 'reset', 'next', 'kind', 'download', 'print', 'format'}
    STRIP_PARAMS = {'show_preview'}
    MAX_PAGES = 80

    def __init__(self, get_response):
        self.get_response = get_response

    def _eligible(self, request):
        user = getattr(request, 'user', None)
        is_sync = request.headers.get('x-sync-page-state') == '1'
        return (request.method == 'GET' and user is not None and user.is_authenticated
                and (is_sync or request.headers.get('x-requested-with') != 'XMLHttpRequest')
                and not request.path.startswith(self.SKIP_PREFIXES))

    def _save_state(self, request, states, path, query):
        if states.get(path) != query:
            states.pop(path, None)
            states[path] = query
            keys = [k for k in states if k != '_date']
            for old in keys[:-self.MAX_PAGES]:
                states.pop(old, None)
            request.session[self.SESSION_KEY] = states

    def __call__(self, request):
        if not self._eligible(request):
            return self.get_response(request)

        from django.http import HttpResponse
        from django.utils import timezone

        today = timezone.localdate().isoformat()
        states = request.session.get(self.SESSION_KEY) or {}
        if states.get('_date') != today:
            states = {'_date': today}
        path = request.path

        # ซิงก์สถานะจากฝั่งเบราว์เซอร์ (เช่น การสลับ Bootstrap Tab ที่ไม่รีโหลดหน้า)
        if request.headers.get('x-sync-page-state') == '1':
            if request.GET and not any(k in request.GET for k in self.VOLATILE_PARAMS):
                clean = request.GET.copy()
                for k in self.STRIP_PARAMS:
                    clean.pop(k, None)
                q_str = clean.urlencode()
                if q_str:
                    self._save_state(request, states, path, q_str)
            return HttpResponse(status=204)

        if 'reset' in request.GET:
            if states.pop(path, None) is not None:
                request.session[self.SESSION_KEY] = states
            return self.get_response(request)

        if any(k in request.GET for k in self.VOLATILE_PARAMS):
            return self.get_response(request)

        if not request.GET:
            saved = states.get(path)
            if saved:
                return redirect(f'{path}?{saved}')
            return self.get_response(request)

        response = self.get_response(request)
        clean = request.GET.copy()
        for k in self.STRIP_PARAMS:
            clean.pop(k, None)
        query = clean.urlencode()
        is_html = 'text/html' in response.get('Content-Type', '')
        if query and response.status_code == 200 and is_html and not response.has_header('Content-Disposition'):
            self._save_state(request, states, path, query)
        elif states.get(path) == query:
            # ค่าที่จดจำไว้ใช้ไม่ได้แล้ว (เช่น ข้อมูลถูกลบ) -> ลืมค่า เพื่อป้องกันการวนซ้ำ
            states.pop(path, None)
            request.session[self.SESSION_KEY] = states
        return response
