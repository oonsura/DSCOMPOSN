from django.conf import settings
from django.contrib import admin
from django.urls import include, path, re_path
from django.views.static import serve as media_serve

admin.site.site_header = 'DS-COMPOSN (Django Admin)'
admin.site.site_title = 'DS-COMPOSN'

urlpatterns = [
    # Django admin ย้ายไปที่ /django-admin/ เพื่อให้ /admin/... เป็นหน้าจัดการของระบบ
    path('django-admin/', admin.site.urls),
    path('', include('food.urls')),
    path('', include('composn.urls')),
]

if settings.DEBUG or getattr(settings, 'SERVE_MEDIA', False):
    urlpatterns += [
        re_path(r'^media/(?P<path>.*)$', media_serve, {'document_root': settings.MEDIA_ROOT}),
    ]

handler403 = 'composn.views_errors.permission_denied'
