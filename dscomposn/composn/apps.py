from django.apps import AppConfig
from django.db.backends.signals import connection_created
from django.db.models.signals import post_migrate


def _configure_sqlite_wal(sender, connection, **kwargs):
    """เปิดโหมด WAL และปรับแต่ง SQLite ให้รองรับการอ่าน/เขียนพร้อมกันได้ดีขึ้น"""
    if connection.vendor != 'sqlite':
        return
    db_name = str(connection.settings_dict.get('NAME') or '')
    if not db_name or db_name == ':memory:' or db_name.startswith('file:memorydb'):
        return
    with connection.cursor() as cursor:
        cursor.execute('PRAGMA journal_mode=WAL;')
        cursor.execute('PRAGMA synchronous=NORMAL;')
        cursor.execute('PRAGMA busy_timeout=30000;')
        cursor.execute('PRAGMA cache_size=-16000;')
        cursor.execute('PRAGMA foreign_keys=ON;')


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
        connection_created.connect(_configure_sqlite_wal)
