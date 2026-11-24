import json

from django.db import migrations, models


def prepend_quick_toggles(apps, schema_editor):
    SystemSetting = apps.get_model('composn', 'SystemSetting')
    default_widgets = [
        'quick_toggles', 'kpi_overview', 'student_summary', 'attendance_summary',
        'scores_certificates', 'rights_summary', 'food_orders', 'survey_summary',
        'school_requests', 'recent_logs',
    ]
    for s in SystemSetting.objects.all():
        raw = s.dashboard_widgets
        if raw is None:
            s.dashboard_widgets = json.dumps(default_widgets)
            s.save(update_fields=['dashboard_widgets'])
            continue
        try:
            data = json.loads(raw)
        except (TypeError, ValueError):
            data = None
        if not isinstance(data, list):
            s.dashboard_widgets = json.dumps(default_widgets)
            s.save(update_fields=['dashboard_widgets'])
            continue
        if 'quick_toggles' not in data:
            s.dashboard_widgets = json.dumps(['quick_toggles'] + data)
            s.save(update_fields=['dashboard_widgets'])


class Migration(migrations.Migration):

    dependencies = [
        ('composn', '0026_user_staff_can_exam_seating'),
    ]

    operations = [
        migrations.AlterField(
            model_name='systemsetting',
            name='dashboard_widgets',
            field=models.TextField(
                default='["quick_toggles", "kpi_overview", "student_summary", "attendance_summary", "scores_certificates", "rights_summary", "food_orders", "survey_summary", "school_requests", "recent_logs"]',
                verbose_name='ส่วนที่แสดงและลำดับบน Dashboard (JSON)',
            ),
        ),
        migrations.RunPython(prepend_quick_toggles, migrations.RunPython.noop),
    ]
