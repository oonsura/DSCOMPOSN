from functools import wraps

from django.contrib import messages
from django.contrib.auth.views import redirect_to_login
from django.core.exceptions import PermissionDenied


def role_required(*roles):
    """จำกัดสิทธิ์ตามบทบาท: @role_required('admin'), @role_required('admin', 'staff')"""

    def decorator(view_func):
        @wraps(view_func)
        def _wrapped(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect_to_login(request.get_full_path())
            if request.user.role not in roles:
                raise PermissionDenied('คุณไม่มีสิทธิ์เข้าถึงหน้านี้')
            return view_func(request, *args, **kwargs)

        return _wrapped

    return decorator


def staff_perm_required(perm_field):
    """อนุญาต Admin เสมอ หรือ Staff ที่ได้รับสิทธิ์ perm_field เป็น True"""

    def decorator(view_func):
        @wraps(view_func)
        def _wrapped(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect_to_login(request.get_full_path())
            if request.user.role == 'admin':
                return view_func(request, *args, **kwargs)
            if request.user.role == 'staff' and getattr(request.user, perm_field, False):
                return view_func(request, *args, **kwargs)
            raise PermissionDenied('คุณไม่มีสิทธิ์เข้าถึงหน้านี้')

        return _wrapped

    return decorator


admin_required = role_required('admin')
staff_required = role_required('admin', 'staff')
member_required = role_required('member')
