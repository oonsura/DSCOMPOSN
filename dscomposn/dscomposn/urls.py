from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import include, path

admin.site.site_header = 'DS-COMPOSN (Django Admin)'
admin.site.site_title = 'DS-COMPOSN'

urlpatterns = [
    # Django admin ย้ายไปที่ /django-admin/ เพื่อให้ /admin/... เป็นหน้าจัดการของระบบ
    path('django-admin/', admin.site.urls),
    path('', include('food.urls')),
    path('', include('composn.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

handler403 = 'composn.views_errors.permission_denied'
