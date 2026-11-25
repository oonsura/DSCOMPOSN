import re

from django.db import migrations


_SCHOOL_INSTITUTION_PREFIXES = (
    'โรงเรียน',
    'สถาบัน',
    'ศูนย์การเรียน',
    'วิทยาลัย',
    'มหาวิทยาลัย',
)
_SCHOOL_ABBREV_RE = re.compile(r'^(?:ร\.\s*ร\.?|รร\.|รร\s+)\s*(.*)$')
_SCHOOL_DASH_ONLY_RE = re.compile(r'^[-–—\s]+$')


def _normalize_school_name(value, max_length=255):
    if value is None:
        return ''
    text = str(value).strip()
    if not text or _SCHOOL_DASH_ONLY_RE.match(text):
        return ''
    text = re.sub(r'\s+', ' ', text)
    m = _SCHOOL_ABBREV_RE.match(text)
    if m:
        rest = re.sub(r'^โรงเรียน\s*', '', m.group(1).strip())
        if not rest:
            return ''
        return f'โรงเรียน{rest}'[:max_length]
    if text.startswith('โรงเรียน'):
        rest = text[len('โรงเรียน'):].lstrip()
        return (f'โรงเรียน{rest}' if rest else 'โรงเรียน')[:max_length]
    if any(text.startswith(p) for p in _SCHOOL_INSTITUTION_PREFIXES):
        return text[:max_length]
    return f'โรงเรียน{text}'[:max_length]


def forwards(apps, schema_editor):
    User = apps.get_model('composn', 'User')
    for u in User.objects.exclude(school_name=''):
        norm = _normalize_school_name(u.school_name)
        if norm != u.school_name:
            u.school_name = norm
            u.save(update_fields=['school_name'])

    UseRecord = apps.get_model('composn', 'UseRecord')
    for rec in UseRecord._base_manager.exclude(school_snapshot=''):
        norm = _normalize_school_name(rec.school_snapshot)
        if norm != rec.school_snapshot:
            rec.school_snapshot = norm
            rec.save(update_fields=['school_snapshot'])

    SchoolChangeRequest = apps.get_model('composn', 'SchoolChangeRequest')
    for req in SchoolChangeRequest.objects.all():
        changed = []
        if req.old_school_name:
            norm_old = _normalize_school_name(req.old_school_name)
            if norm_old != req.old_school_name:
                req.old_school_name = norm_old
                changed.append('old_school_name')
        if req.new_school_name:
            norm_new = _normalize_school_name(req.new_school_name)
            if norm_new != req.new_school_name:
                req.new_school_name = norm_new
                changed.append('new_school_name')
        if changed:
            req.save(update_fields=changed)


class Migration(migrations.Migration):

    dependencies = [
        ('composn', '0028_alter_user_staff_can_scores'),
    ]

    operations = [
        migrations.RunPython(forwards, migrations.RunPython.noop),
    ]
