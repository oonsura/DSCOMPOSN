from django.apps import AppConfig
from django.db.models.signals import post_migrate

DEFAULT_REASONS = ['ปกติ', 'ป่วย', 'ลากิจ', 'อื่นๆ']
DEFAULT_COLORS = [
    ('สีเหลือง', '#ffc107', 1),
    ('สีส้ม', '#fd7e14', 3),
    ('สีแดง', '#dc3545', 5),
]


def ensure_defaults(sender, **kwargs):
    from .models import AbsenceReason, WarningColor

    if not AbsenceReason.objects.exists():
        for i, name in enumerate(DEFAULT_REASONS, start=1):
            AbsenceReason.objects.create(name=name, order=i)
    if not WarningColor.objects.exists():
        for i, (name, color, min_abs) in enumerate(DEFAULT_COLORS, start=1):
            WarningColor.objects.create(name=name, color=color, min_absences=min_abs, order=i)


class FoodConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'food'
    verbose_name = 'Food (เวลาเรียน / คะแนน / อาหาร)'

    def ready(self):
        post_migrate.connect(ensure_defaults, sender=self)
