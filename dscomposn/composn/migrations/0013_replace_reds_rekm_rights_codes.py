from django.db import migrations
from django.db.models import Q


def replace_reds_rekm(apps, schema_editor):
    UseRecord = apps.get_model('composn', 'UseRecord')
    ArchiveRecord = apps.get_model('composn', 'ArchiveRecord')

    re_filter = Q(rights_code__istartswith='REDS') | Q(rights_code__istartswith='REKM')

    for rec in UseRecord._base_manager.select_related('user').filter(re_filter):
        target = (
            (rec.user.camp2_code if rec.camp_level == 2 else rec.user.camp1_code)
            or rec.camp_code_snapshot
            or rec.user.camp1_code
        )
        if not target:
            continue
        if not rec.is_deleted:
            conflict = UseRecord._base_manager.filter(
                is_deleted=False,
                academic_year=rec.academic_year,
                rights_code=target,
            ).exclude(pk=rec.pk).exists()
            if conflict:
                raise RuntimeError(
                    f'Cannot migrate UseRecord pk={rec.pk} to rights_code={target}: '
                    f'active record already exists for year {rec.academic_year}'
                )
        rec.rights_code = target
        rec.camp_code_snapshot = target
        rec.save(update_fields=['rights_code', 'camp_code_snapshot'])

    for arc in ArchiveRecord._base_manager.select_related('user').filter(re_filter):
        target = (
            (arc.user.camp2_code if arc.camp_level == 2 else arc.user.camp1_code)
            or arc.user.camp1_code
        )
        if not target:
            continue
        arc.rights_code = target
        arc.save(update_fields=['rights_code'])


class Migration(migrations.Migration):

    dependencies = [
        ('composn', '0012_alter_systemsetting_announcement_title_and_more'),
    ]

    operations = [
        migrations.RunPython(replace_reds_rekm, migrations.RunPython.noop),
    ]
