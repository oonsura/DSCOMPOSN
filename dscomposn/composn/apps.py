from django.apps import AppConfig
from django.db.models.signals import post_migrate


def ensure_defaults(sender, **kwargs):
    """ค่าเริ่มต้น (Fixtures): Superuser admin / 1234 และตั้งค่าระบบ"""
    from django.conf import settings
    from .models import SystemSetting, User

    SystemSetting.get()
    if not User.objects.filter(camp1_code=settings.DEFAULT_ADMIN_USERNAME).exists():
        User.objects.create_superuser(
            settings.DEFAULT_ADMIN_USERNAME, settings.DEFAULT_ADMIN_PASSWORD,
            first_name='ผู้ดูแลระบบ', role=User.ROLE_ADMIN,
        )


class ComposnConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'composn'
    verbose_name = 'COMPOSN (สิทธิ์ค่าย)'

    def ready(self):
        post_migrate.connect(ensure_defaults, sender=self)
