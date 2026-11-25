"""
Models ของแอปหลัก COMPOSN
- User (Custom User Model: camp1_code เป็น USERNAME_FIELD)
- SystemSetting (ตั้งค่าระบบ แบบ Singleton)
- RightsRound (รอบใช้สิทธิ์)
- UseRecord / ArchiveRecord (ประวัติการใช้สิทธิ์)
- SchoolChangeRequest (คำขอแก้ไขชื่อโรงเรียน)
- AuditLog (บันทึกการกระทำสำคัญ)
"""
import json
import re

from django.conf import settings
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.core.cache import cache
from django.core.exceptions import ValidationError
from django.core.validators import RegexValidator
from django.db import models
from django.db.models import Q
from django.utils import timezone

CAMP_CODE_REGEX = r'^\d{2}(DS|KM)\d{2}$'
camp_code_validator = RegexValidator(
    CAMP_CODE_REGEX,
    'รูปแบบรหัสค่ายไม่ถูกต้อง ต้องเป็น เลข 2 หลัก + DS/KM + เลข 2 หลัก เช่น 69DS01',
)

EXPIRED_TEXT = 'หมดสิทธิ์'

SCHOOL_INSTITUTION_PREFIXES = (
    'โรงเรียน',
    'สถาบัน',
    'ศูนย์การเรียน',
    'วิทยาลัย',
    'มหาวิทยาลัย',
)
_SCHOOL_ABBREV_RE = re.compile(r'^(?:ร\.\s*ร\.?|รร\.|รร\s+)\s*(.*)$')
_SCHOOL_DASH_ONLY_RE = re.compile(r'^[-–—\s]+$')


def normalize_school_name(value, max_length=255):
    """ปรับชื่อโรงเรียนให้มีคำว่า 'โรงเรียน' นำหน้าโดยอัตโนมัติ
    - ค่าว่างหรือมีเฉพาะขีด (-) -> คืนค่าว่าง ''
    - คำย่อ รร. / ร.ร. -> แปลงเป็น 'โรงเรียน'
    - 'โรงเรียน   ...' -> ลบช่องว่างเกินหลังคำว่า 'โรงเรียน'
    - สถานศึกษาที่ขึ้นต้นด้วย สถาบัน / ศูนย์การเรียน / วิทยาลัย / มหาวิทยาลัย -> คงไว้ตามเดิม
    - กรณีอื่น ๆ ที่ยังไม่มีคำว่า 'โรงเรียน' นำหน้า -> เติม 'โรงเรียน' นำหน้าอัตโนมัติ
    """
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
    if any(text.startswith(p) for p in SCHOOL_INSTITUTION_PREFIXES):
        return text[:max_length]
    return f'โรงเรียน{text}'[:max_length]


# =============================================================================
# Soft Delete
# =============================================================================
class SoftDeleteQuerySet(models.QuerySet):
    def delete(self, user=None):
        """Soft delete ทั้ง QuerySet"""
        return super().update(is_deleted=True, deleted_at=timezone.now(), deleted_by=user)

    def hard_delete(self):
        return super().delete()

    def alive(self):
        return self.filter(is_deleted=False)

    def dead(self):
        return self.filter(is_deleted=True)


class SoftDeleteManager(models.Manager):
    """Manager ปกติ: ซ่อนแถวที่ถูกลบแบบ Soft Delete"""

    def get_queryset(self):
        return SoftDeleteQuerySet(self.model, using=self._db).filter(is_deleted=False)


class AllObjectsManager(models.Manager):
    def get_queryset(self):
        return SoftDeleteQuerySet(self.model, using=self._db)


class SoftDeleteModel(models.Model):
    is_deleted = models.BooleanField('ถูกลบ', default=False, db_index=True)
    deleted_at = models.DateTimeField('ลบเมื่อ', null=True, blank=True)
    deleted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='ลบโดย',
    )

    objects = SoftDeleteManager()
    all_objects = AllObjectsManager()

    class Meta:
        abstract = True

    def delete(self, using=None, keep_parents=False, user=None):  # noqa: D401
        """Soft delete — ไม่ลบจริงจากฐานข้อมูล"""
        self.is_deleted = True
        self.deleted_at = timezone.now()
        self.deleted_by = user
        self.save(update_fields=['is_deleted', 'deleted_at', 'deleted_by'])

    def hard_delete(self):
        return super().delete()

    def restore(self):
        self.is_deleted = False
        self.deleted_at = None
        self.deleted_by = None
        self.save(update_fields=['is_deleted', 'deleted_at', 'deleted_by'])


# =============================================================================
# User
# =============================================================================
class UserManager(BaseUserManager):
    use_in_migrations = True

    def create_user(self, camp1_code, password=None, **extra_fields):
        if not camp1_code:
            raise ValueError('ต้องระบุ camp1_code')
        camp1_code = str(camp1_code).strip().upper() if extra_fields.get('role', 'member') == 'member' \
            else str(camp1_code).strip()
        user = self.model(camp1_code=camp1_code, **extra_fields)
        # รหัสผ่านเริ่มต้นของ Member = รหัสค่าย 1 เต็ม
        user.set_password(password if password else camp1_code)
        user.save(using=self._db)
        return user

    def create_superuser(self, camp1_code, password=None, **extra_fields):
        extra_fields.setdefault('role', User.ROLE_ADMIN)
        extra_fields.setdefault('is_superuser', True)
        return self.create_user(camp1_code, password or settings.DEFAULT_ADMIN_PASSWORD, **extra_fields)

    def get_by_natural_key(self, username):
        # รองรับการเข้าสู่ระบบด้วยรหัสค่าย 1 (camp1_code) หรือรหัสค่าย 2 (camp2_code) แบบไม่สนตัวพิมพ์เล็ก-ใหญ่
        username = (username or '').strip()
        if not username:
            raise self.model.DoesNotExist
        try:
            return self.get(camp1_code=username)
        except self.model.DoesNotExist:
            pass
        try:
            return self.get(camp1_code__iexact=username)
        except self.model.DoesNotExist:
            pass
        try:
            return self.get(camp2_code=username.upper())
        except self.model.DoesNotExist:
            return self.get(camp2_code__iexact=username)


class User(AbstractBaseUser, PermissionsMixin):
    ROLE_ADMIN = 'admin'
    ROLE_STAFF = 'staff'
    ROLE_MEMBER = 'member'
    ROLE_CHOICES = [
        (ROLE_ADMIN, 'ผู้ดูแลระบบ'),
        (ROLE_STAFF, 'เจ้าหน้าที่'),
        (ROLE_MEMBER, 'นักเรียน/ผู้ใช้สิทธิ์'),
    ]
    CENTER_CHOICES = [('DS', 'ศูนย์ DS'), ('KM', 'ศูนย์ KM')]

    camp1_code = models.CharField('รหัสค่าย 1', max_length=10, unique=True)
    camp2_code = models.CharField('รหัสค่าย 2', max_length=10, null=True, blank=True, unique=True,
                                  validators=[camp_code_validator])
    first_name = models.CharField('ชื่อ', max_length=150, blank=True)
    last_name = models.CharField('นามสกุล', max_length=150, blank=True)
    nickname = models.CharField('ชื่อเล่น', max_length=60, blank=True)
    school_name = models.CharField('โรงเรียน', max_length=255, blank=True)
    birth_date = models.DateField('วันเกิด', null=True, blank=True)
    camp_center = models.CharField('ศูนย์', max_length=2, choices=CENTER_CHOICES, blank=True)
    academic_year = models.IntegerField('ปีการศึกษา (พ.ศ.)', null=True, blank=True)
    current_level = models.CharField('ระดับชั้นปัจจุบัน', max_length=30, blank=True)
    year_plus_1 = models.CharField('ระดับชั้น +1 ปี', max_length=30, blank=True)
    year_plus_2 = models.CharField('ระดับชั้น +2 ปี', max_length=30, blank=True)
    year_plus_3 = models.CharField('ระดับชั้น +3 ปี', max_length=30, blank=True)
    year_plus_4 = models.CharField('ระดับชั้น +4 ปี', max_length=30, blank=True)
    year_plus_5 = models.CharField('ระดับชั้น +5 ปี', max_length=30, blank=True)
    role = models.CharField('บทบาท', max_length=10, choices=ROLE_CHOICES, default=ROLE_MEMBER, db_index=True)
    is_first_login = models.BooleanField('ต้องเปลี่ยนรหัสผ่านครั้งแรก', default=True)
    is_active = models.BooleanField('เปิดใช้งาน', default=True)
    student_phone = models.CharField('เบอร์โทรนักเรียน', max_length=30, blank=True)
    parent_phone = models.CharField('เบอร์โทรผู้ปกครอง', max_length=30, blank=True)
    email = models.EmailField('อีเมล', max_length=150, blank=True)
    chronic_disease = models.CharField('โรคประจำตัว', max_length=255, blank=True)
    food_allergy = models.CharField('แพ้อาหาร', max_length=255, blank=True)
    is_muslim = models.BooleanField('ศาสนาอิสลาม', default=False)
    profile_photo = models.ImageField('รูปภาพประจำตัว', upload_to='profiles/', null=True, blank=True)
    # สิทธิ์รายบุคคลสำหรับเจ้าหน้าที่ (Staff) — Admin มีสิทธิ์ทั้งหมดโดยอัตโนมัติ
    staff_can_view_users = models.BooleanField('ผู้ใช้ (นักเรียน/สมาชิก) — ดูอย่างเดียว', default=False)
    staff_can_attendance = models.BooleanField('บันทึกเวลาเรียน', default=True)
    staff_can_signin_sheet = models.BooleanField('แบบลงชื่อ', default=True)
    staff_can_manage_food = models.BooleanField('รายการอาหาร / แบบสอบถาม', default=True)
    staff_can_scores = models.BooleanField('บันทึกคะแนน / การประกาศผล', default=True)
    staff_can_exam_seating = models.BooleanField('จัดที่นั่งสอบ', default=False)
    note = models.TextField('หมายเหตุ', blank=True)
    current_session_key = models.CharField(max_length=64, blank=True, editable=False)
    created_at = models.DateTimeField('สร้างเมื่อ', auto_now_add=True)
    updated_at = models.DateTimeField('แก้ไขเมื่อ', auto_now=True)

    objects = UserManager()

    USERNAME_FIELD = 'camp1_code'
    REQUIRED_FIELDS = []

    class Meta:
        verbose_name = 'ผู้ใช้'
        verbose_name_plural = 'ผู้ใช้'
        ordering = ['camp1_code']

    def __str__(self):
        return f'{self.camp1_code} {self.get_full_name()}'.strip()

    # ---------- Validation ----------
    def clean(self):
        super().clean()
        if self.camp1_code:
            self.camp1_code = self.camp1_code.strip()
        if self.camp2_code:
            self.camp2_code = self.camp2_code.strip().upper()
        else:
            self.camp2_code = None
        if self.role == self.ROLE_MEMBER:
            self.camp1_code = (self.camp1_code or '').upper()
            if not re.match(CAMP_CODE_REGEX, self.camp1_code):
                raise ValidationError({'camp1_code': camp_code_validator.message})
        errors = {}
        # รหัสค่ายต้อง Unique ทั้งระบบ (ข้ามคอลัมน์ camp1/camp2 ด้วย)
        if self.camp1_code and User.objects.exclude(pk=self.pk).filter(camp2_code__iexact=self.camp1_code).exists():
            errors['camp1_code'] = 'รหัสนี้ถูกใช้เป็นรหัสค่าย 2 ของผู้ใช้อื่นแล้ว'
        if self.camp2_code:
            if self.camp2_code == self.camp1_code:
                errors['camp2_code'] = 'รหัสค่าย 2 ต้องไม่ซ้ำกับรหัสค่าย 1'
            elif User.objects.exclude(pk=self.pk).filter(camp1_code__iexact=self.camp2_code).exists():
                errors['camp2_code'] = 'รหัสนี้ถูกใช้เป็นรหัสค่าย 1 ของผู้ใช้อื่นแล้ว'
        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        # camp2_code ว่าง ต้องเก็บเป็น NULL เพื่อไม่ชน unique
        self.camp2_code = (self.camp2_code or '').strip().upper() or None
        if self.role == self.ROLE_MEMBER and self.camp1_code:
            self.camp1_code = self.camp1_code.strip().upper()
        if self.school_name:
            self.school_name = normalize_school_name(self.school_name)
        if self.profile_photo and not getattr(self.profile_photo, '_committed', True):
            from .image_utils import compress_profile_fallback_image, ensure_compressed_field
            ensure_compressed_field(self, 'profile_photo', compress_profile_fallback_image)
        super().save(*args, **kwargs)

    def check_password(self, raw_password):
        """นักเรียนที่เข้าสู่ระบบครั้งแรกสามารถพิมพ์รหัสผ่านเริ่มต้น (รหัสค่าย 1) ได้ทั้งตัวพิมพ์เล็กและตัวพิมพ์ใหญ่"""
        ok = super().check_password(raw_password)
        if not ok and self.role == self.ROLE_MEMBER and self.is_first_login and raw_password:
            if str(raw_password).strip().upper() == (self.camp1_code or '').upper():
                return super().check_password(self.camp1_code)
        return ok

    # ---------- Helpers ----------
    def get_full_name(self):
        return f'{self.first_name} {self.last_name}'.strip()

    def get_short_name(self):
        return self.first_name

    @property
    def username(self):
        return self.camp1_code

    @property
    def is_staff(self):
        """ให้ Admin เข้า Django admin (/django-admin/) ได้"""
        return self.role == self.ROLE_ADMIN and self.is_active

    @property
    def is_admin_role(self):
        return self.role == self.ROLE_ADMIN

    @property
    def is_staff_role(self):
        return self.role in (self.ROLE_ADMIN, self.ROLE_STAFF)

    @property
    def is_member_role(self):
        return self.role == self.ROLE_MEMBER

    def has_staff_perm(self, perm_field):
        if self.role == self.ROLE_ADMIN:
            return True
        if self.role == self.ROLE_STAFF:
            return bool(getattr(self, perm_field, False))
        return False

    @property
    def can_view_users(self):
        return self.has_staff_perm('staff_can_view_users')

    @property
    def can_attendance(self):
        return self.has_staff_perm('staff_can_attendance')

    @property
    def can_signin_sheet(self):
        return self.has_staff_perm('staff_can_signin_sheet')

    @property
    def can_manage_food(self):
        return self.has_staff_perm('staff_can_manage_food')

    @property
    def can_scores(self):
        return self.has_staff_perm('staff_can_scores')

    @property
    def can_exam_seating(self):
        return self.has_staff_perm('staff_can_exam_seating')

    def staff_permission_labels(self):
        """คืนรายการชื่อสิทธิ์ของ Staff ที่เปิดใช้งาน"""
        if self.role == self.ROLE_ADMIN:
            return ['สิทธิ์ทั้งหมด (Admin)']
        if self.role != self.ROLE_STAFF:
            return []
        items = []
        if self.staff_can_view_users:
            items.append('ผู้ใช้ (นักเรียน/สมาชิก)')
        if self.staff_can_attendance:
            items.append('บันทึกเวลาเรียน')
        if self.staff_can_signin_sheet:
            items.append('แบบลงชื่อ')
        if self.staff_can_manage_food:
            items.append('รายการอาหาร / แบบสอบถาม')
        if self.staff_can_scores:
            items.append('บันทึกคะแนน')
        if self.staff_can_exam_seating:
            items.append('จัดที่นั่งสอบ')
        return items

    @property
    def camp_level(self):
        """มีรหัสค่าย 2 = ใช้สิทธิ์ค่าย 2, ไม่มี = ค่าย 1"""
        return 2 if self.camp2_code else 1

    @property
    def base_academic_year(self):
        """ปีการศึกษาฐานของ current_level (ใช้ academic_year หรือคำนวณจาก 2 หลักแรกของรหัสค่าย 1)"""
        if self.academic_year:
            return self.academic_year
        m = re.match(r'^(\d{2})', self.camp1_code or '')
        return 2500 + int(m.group(1)) if m else None

    def future_levels(self):
        """[(ปี พ.ศ., ระดับชั้น), ...] ตั้งแต่ปีปัจจุบันถึง +5"""
        base = self.base_academic_year
        levels = [self.current_level, self.year_plus_1, self.year_plus_2,
                  self.year_plus_3, self.year_plus_4, self.year_plus_5]
        return [((base + i) if base else None, lv or '-') for i, lv in enumerate(levels)]

    def level_for_year(self, year):
        """ระดับชั้นในปีการศึกษาที่ระบุ ('' = ไม่ทราบ, 'หมดสิทธิ์' = เกินระดับ)"""
        base = self.base_academic_year
        if not base or not year:
            return ''
        offset = int(year) - int(base)
        if offset < 0:
            return ''
        if offset > 5:
            return EXPIRED_TEXT
        return [self.current_level, self.year_plus_1, self.year_plus_2,
                self.year_plus_3, self.year_plus_4, self.year_plus_5][offset] or ''

    def rights_used_count(self):
        """นับจำนวนครั้งที่ใช้สิทธิ์จากจำนวนแถวใน ARCHIVE_RECORD (ไม่ใช้ฟิลด์ total_rights_used)"""
        return ArchiveRecord.objects.filter(user=self).count()


# =============================================================================
# System Setting (Singleton)
# =============================================================================
PROFILE_EDITABLE_CHOICES = [
    ('first_name', 'ชื่อ-นามสกุล'),
    ('nickname', 'ชื่อเล่น'),
    ('birth_date', 'วันเกิด'),
    ('student_phone', 'เบอร์โทรนักเรียน'),
    ('parent_phone', 'เบอร์โทรผู้ปกครอง'),
    ('email', 'อีเมล'),
    ('chronic_disease', 'โรคประจำตัว'),
    ('food_allergy', 'แพ้อาหาร'),
    ('is_muslim', 'ศาสนาอิสลาม (Checkbox)'),
    ('note', 'หมายเหตุ'),
]

DASHBOARD_SUMMARY_CHOICES = [
    ('is_muslim', 'จำนวนศาสนาอิสลาม (อาหารฮาลาล)'),
    ('food_allergy', 'คนที่แพ้อาหาร'),
    ('chronic_disease', 'คนที่มีโรคประจำตัว'),
    ('student_phone', 'กรอกเบอร์โทรนักเรียนแล้ว'),
    ('parent_phone', 'กรอกเบอร์โทรผู้ปกครองแล้ว'),
    ('email', 'กรอกอีเมลแล้ว'),
    ('profile_photo', 'อัปโหลดรูปประจำตัวแล้ว'),
    ('camp2_code', 'ได้รับรหัสค่าย 2 แล้ว'),
]

DASHBOARD_WIDGET_CHOICES = [
    ('quick_toggles', 'ศูนย์ควบคุมเปิด / ปิดระบบ'),
    ('kpi_overview', 'การ์ดตัวเลขภาพรวม (KPI)'),
    ('student_summary', 'สรุปข้อมูลนักเรียน'),
    ('attendance_summary', 'สรุปเวลาเรียนและการเข้าอบรม'),
    ('scores_certificates', 'สรุปคะแนน การคัดเลือก และเกียรติบัตร'),
    ('rights_summary', 'รอบและสรุปการใช้สิทธิ์'),
    ('food_orders', 'รายการอาหารและยอดผู้เลือก'),
    ('survey_summary', 'สรุปผลแบบสอบถาม'),
    ('school_requests', 'คำขอเปลี่ยนชื่อโรงเรียนรออนุมัติ'),
    ('recent_logs', 'กิจกรรมล่าสุด (Audit Log)'),
]
DEFAULT_DASHBOARD_WIDGETS = [k for k, _ in DASHBOARD_WIDGET_CHOICES]

MIN_ACADEMIC_YEAR = 2566
MAX_ACADEMIC_YEAR = 2700
DEFAULT_ACADEMIC_YEARS = [2566, 2567, 2568, 2569]


class SystemSetting(models.Model):
    current_academic_year = models.IntegerField('ปีการศึกษาปัจจุบัน (พ.ศ.)', default=2569)
    academic_years = models.TextField(
        'รายการปีการศึกษาในระบบ (JSON)',
        default='[2566, 2567, 2568, 2569]',
        help_text='รายการปีการศึกษา (พ.ศ.) ที่เปิดใช้งานในระบบ เริ่มตั้งแต่ปี 2566',
    )
    max_rights = models.PositiveIntegerField('จำนวนครั้งใช้สิทธิ์สูงสุดตลอดชีวิต', default=5)
    max_absences = models.PositiveIntegerField('จำนวนที่ขาดอบรมสูงสุด (ภาคเวลา)', default=5)
    auto_logout_minutes = models.PositiveIntegerField(
        'ตั้งเวลาตัดการเชื่อมต่ออัตโนมัติ (นาที)', default=5,
        help_text='ออกจากระบบอัตโนมัติเมื่อไม่มีการใช้งานตามจำนวนนาทีที่กำหนด (ใส่ 0 = ไม่ตัดอัตโนมัติ)',
    )
    editable_profile_fields = models.TextField(
        'ฟิลด์ที่ Member แก้ไขได้ (JSON)', default='["student_phone", "parent_phone", "email", "chronic_disease", "food_allergy", "is_muslim"]',
        help_text='รายการชื่อฟิลด์ที่อนุญาตให้ Member แก้ไขเองได้',
    )
    dashboard_summary_fields = models.TextField(
        'ฟิลด์ข้อมูลนักเรียนที่แสดงสรุปบน Dashboard (JSON)',
        default='["is_muslim", "food_allergy", "chronic_disease", "student_phone", "parent_phone", "email"]',
    )
    dashboard_widgets = models.TextField(
        'ส่วนที่แสดงและลำดับบน Dashboard (JSON)',
        default='["quick_toggles", "kpi_overview", "student_summary", "attendance_summary", "scores_certificates", "rights_summary", "food_orders", "survey_summary", "school_requests", "recent_logs"]',
    )
    site_title = models.CharField('ชื่อระบบ (บรรทัดที่ 1)', max_length=150, default='ระบบ DS-COMPOSN')
    site_title_line2 = models.CharField('ชื่อระบบ (บรรทัดที่ 2)', max_length=150, blank=True, default='')
    display_title = models.CharField('หัวข้อหน้าแสดงผลสาธารณะ', max_length=200,
                                     default='สรุปการใช้สิทธิ์ สอวน.คอมพิวเตอร์ ศูนย์อบรมโรงเรียนเทพศิรินทร์')
    # การตั้งค่าระบบคะแนนและการคัดเลือก
    passing_score = models.DecimalField('คะแนนเกณฑ์ผ่านการคัดเลือก', max_digits=7, decimal_places=2, default=50)
    max_qualified_count = models.PositiveIntegerField('จำนวนผู้ที่ผ่านการคัดเลือก (คน)', default=30)
    announcement_title = models.CharField(
        'หัวข้อเอกสารประกาศผล', max_length=255,
        default='ประกาศผลการคัดเลือกเข้าค่ายโอลิมปิกวิชาการ (สอวน.)',
    )
    announcement_school_name = models.CharField('ชื่อโรงเรียนในประกาศ', max_length=150, default='โรงเรียนเทพศิรินทร์')
    announcement_center_name = models.CharField(
        'ชื่อศูนย์อบรมในประกาศ', max_length=255,
        default='ศูนย์โรงเรียนเทพศิรินทร์-สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหาร ลาดกระบัง',
    )
    announcement_director_name = models.CharField('ชื่อผู้อำนวยการ', max_length=150, default='นายวิธาน พรหมสินธุศักดิ์')
    announcement_director_position = models.TextField(
        'ตำแหน่งผู้อำนวยการ',
        default='ผู้อำนวยการโรงเรียนเทพศิรินทร์\nประธานศูนย์โรงเรียนเทพศิรินทร์ สอวน.วิชาคอมพิวเตอร์ ค่าย 1',
    )
    announcement_body_p1 = models.TextField(
        'ข้อความย่อหน้าที่ 1', blank=True, default='',
        help_text='เว้นว่างเพื่อใช้ข้อความเริ่มต้นตามศูนย์อบรม',
    )
    announcement_body_p2 = models.TextField(
        'ข้อความย่อหน้าที่ 2', blank=True, default='',
        help_text='เว้นว่างเพื่อใช้ข้อความเริ่มต้น (รองรับ {count} สำหรับจำนวนคน เช่น 30 คน)',
    )
    announcement_date_text = models.CharField(
        'วันที่ในประกาศ', max_length=150, blank=True, default='',
        help_text='เช่น ประกาศ ณ วันที่ 2 มกราคม พ.ศ. 2570 (เว้นว่างเพื่อใช้วันที่ปัจจุบัน)',
    )
    announcement_list_title = models.CharField(
        'หัวข้อรายชื่อแนบท้ายประกาศ', max_length=255, blank=True, default='',
        help_text='เช่น ประกาศรายชื่อผู้ที่ผ่านการคัดเลือกเข้า สอวน.คอมพิวเตอร์ (เว้นว่างเพื่อใช้ข้อความเริ่มต้นพร้อมปีการศึกษา)',
    )
    announcement_note = models.TextField('หมายเหตุท้ายประกาศผล', blank=True)
    # การปรับแต่งระยะขอบ ขนาดตัวอักษร และตารางประกาศผล (A4)
    announcement_margin_top_mm = models.FloatField('ระยะขอบบน (มม.)', default=25.0)
    announcement_margin_bottom_mm = models.FloatField('ระยะขอบล่าง (มม.)', default=25.0)
    announcement_margin_left_mm = models.FloatField('กั้นหน้า (ซ้าย) (มม.)', default=25.0)
    announcement_margin_right_mm = models.FloatField('กั้นหลัง (ขวา) (มม.)', default=25.0)
    announcement_font_size_title1 = models.FloatField('ขนาดตัวอักษรหัวข้อประกาศบรรทัดที่ 1 (pt)', default=17.0)
    announcement_font_size_title2 = models.FloatField('ขนาดตัวอักษรหัวข้อประกาศบรรทัดที่ 2 (pt)', default=15.0)
    announcement_font_size_title3 = models.FloatField('ขนาดตัวอักษรหัวข้อประกาศบรรทัดที่ 3 (pt)', default=14.0)
    announcement_font_size_body = models.FloatField('ขนาดตัวอักษรเนื้อหาประกาศ (pt)', default=14.5)
    announcement_font_size_director_name = models.FloatField('ขนาดตัวอักษรชื่อผู้ลงนาม (pt)', default=14.5)
    announcement_font_size_director_pos = models.FloatField('ขนาดตัวอักษรตำแหน่งผู้ลงนาม (pt)', default=13.0)
    announcement_font_size_list_title1 = models.FloatField('ขนาดตัวอักษรหัวข้อรายชื่อแนบท้าย (pt)', default=14.0)
    announcement_font_size_list_title2 = models.FloatField('ขนาดตัวอักษรศูนย์อบรมแผ่นที่ 2 (pt)', default=12.0)
    announcement_font_size_table = models.FloatField('ขนาดตัวอักษรตารางรายชื่อ (pt)', default=10.0)
    announcement_line_spacing = models.FloatField('ระยะห่างบรรทัดเนื้อหา', default=1.85)
    announcement_paragraph_spacing_mm = models.FloatField('ระยะห่างระหว่างย่อหน้า (มม.)', default=4.0)
    announcement_indent_mm = models.FloatField('ระยะย่อหน้าเนื้อหา/วันที่ (มม.)', default=25.0)
    announcement_garuda_height_mm = models.FloatField('ความสูงตราครุฑ (มม.)', default=30.0)
    announcement_col_w_no = models.FloatField('ความกว้างคอลัมน์ ลำดับ (มม.)', default=14.0)
    announcement_col_w_camp2 = models.FloatField('ความกว้างคอลัมน์ เลขค่าย 2 (มม.)', default=24.0)
    announcement_col_w_rights = models.FloatField('ความกว้างคอลัมน์ เลขค่าย 1/สิทธิ์ (มม.)', default=24.0)
    announcement_col_w_name = models.FloatField('ความกว้างคอลัมน์ ชื่อ-นามสกุล (มม.)', default=50.0)
    announcement_col_w_school = models.FloatField('ความกว้างคอลัมน์ โรงเรียน (มม.)', default=54.0)
    announcement_col_w_grade = models.FloatField('ความกว้างคอลัมน์ ระดับชั้น (มม.)', default=14.0)
    # การตั้งค่าหน้าอาหาร
    food_system_open = models.BooleanField('เปิดระบบเลือกอาหารสำหรับสมาชิก', default=True)
    food_closed_message = models.CharField('ข้อความเมื่อปิดระบบอาหาร', max_length=255,
                                           default='ขณะนี้ปิดระบบการเลือกอาหารชั่วคราว')
    # การตั้งค่าหน้าการใช้สิทธิ์
    rights_warning_text = models.CharField(
        'ข้อความแจ้งเตือนหน้าการใช้สิทธิ์', max_length=255,
        default='กรุณาตรวจสอบวันสอบท้ายค่าย ตามตารางอบรมฯ ก่อนกดยืนยันใช้สิทธิ์สอบท้ายค่าย',
    )
    exam_date_camp1 = models.CharField('ข้อความวันสอบค่าย 1', max_length=150, blank=True)
    exam_date_camp2 = models.CharField('ข้อความวันสอบค่าย 2', max_length=150, blank=True)
    # โทนสีประจำวัน (สำหรับการ์ดอาหาร)
    color_monday = models.CharField('สีวันจันทร์', max_length=20, default='#f9a825')
    color_tuesday = models.CharField('สีวันอังคาร', max_length=20, default='#e91e8c')
    color_wednesday = models.CharField('สีวันพุธ', max_length=20, default='#2e7d32')
    color_thursday = models.CharField('สีวันพฤหัสบดี', max_length=20, default='#e65100')
    color_friday = models.CharField('สีวันศุกร์', max_length=20, default='#1565c0')
    color_saturday = models.CharField('สีวันเสาร์', max_length=20, default='#6a1b9a')
    color_sunday = models.CharField('สีวันอาทิตย์', max_length=20, default='#c62828')
    # โทนสี Sidebar สำหรับสมาชิก และเจ้าหน้าที่ (Staff)
    member_theme_primary = models.CharField('สีหลักสมาชิก', max_length=20, default='#1b5e20')
    member_theme_gradient = models.CharField('สีเฉดรองสมาชิก', max_length=20, default='#2e7d32')
    staff_theme_primary = models.CharField('สีหลักแถบข้างเจ้าหน้าที่ (ด้านบน)', max_length=20, default='#0f4c5c')
    staff_theme_gradient = models.CharField('สีเฉดรองแถบข้างเจ้าหน้าที่ (ด้านล่าง)', max_length=20, default='#1a759f')
    # การปรับแต่งหน้า Login และตราสัญลักษณ์ 3 รูป (ตราโรงเรียน, ตราศูนย์อบรม, สพฐ.)
    login_subtitle = models.CharField('ข้อความรองหน้าล็อกอิน', max_length=200, default='กรุณาเข้าสู่ระบบ')
    login_logo = models.ImageField('โลโก้หน้าล็อกอิน', upload_to='system/', null=True, blank=True)
    logo_school = models.ImageField('รูปตราโรงเรียน', upload_to='system/', null=True, blank=True)
    logo_posn = models.ImageField('รูปตรา สอวน.', upload_to='system/', null=True, blank=True)
    logo_center = models.ImageField('รูปตราศูนย์อบรม', upload_to='system/', null=True, blank=True)
    logo_obec = models.ImageField('รูปสพฐ.', upload_to='system/', null=True, blank=True)
    logo_garuda = models.ImageField('รูปตราครุฑ', upload_to='system/', null=True, blank=True)
    site_favicon = models.ImageField('ไอคอนหน้า Titlebar (Favicon)', upload_to='system/', null=True, blank=True)
    login_bg_color1 = models.CharField('สีพื้นหลังล็อกอิน 1', max_length=20, default='#1a237e')
    login_bg_color2 = models.CharField('สีพื้นหลังล็อกอิน 2', max_length=20, default='#283593')
    login_bg_color3 = models.CharField('สีพื้นหลังล็อกอิน 3', max_length=20, default='#3949ab')
    login_btn_color1 = models.CharField('สีปุ่มล็อกอิน 1', max_length=20, default='#1a73e8')
    login_btn_color2 = models.CharField('สีปุ่มล็อกอิน 2', max_length=20, default='#0d47a1')
    updated_at = models.DateTimeField(auto_now=True)

    @property
    def site_title_inline(self):
        parts = [(self.site_title or '').strip(), (self.site_title_line2 or '').strip()]
        return ' '.join(p for p in parts if p) or 'ระบบ DS-COMPOSN'

    def get_day_color(self, weekday_int):
        return {
            0: self.color_monday, 1: self.color_tuesday, 2: self.color_wednesday,
            3: self.color_thursday, 4: self.color_friday, 5: self.color_saturday, 6: self.color_sunday,
        }.get(weekday_int, '#1a237e')

    class Meta:
        verbose_name = 'ตั้งค่าระบบ'
        verbose_name_plural = 'ตั้งค่าระบบ'

    def __str__(self):
        return 'ตั้งค่าระบบ'

    CACHE_KEY = 'dscomposn:system_setting'

    def save(self, *args, **kwargs):
        self.pk = 1
        from .image_utils import ensure_compressed_field
        from .logo_utils import LOGO_FIELD_NAMES, process_logo_upload
        for fname in LOGO_FIELD_NAMES:
            fld = getattr(self, fname, None)
            if fld and not getattr(fld, '_committed', True):
                ensure_compressed_field(
                    self, fname,
                    lambda f, is_fav=(fname == 'site_favicon'): process_logo_upload(f, is_favicon=is_fav),
                )
        super().save(*args, **kwargs)
        cache.delete_many([self.CACHE_KEY, 'dscomposn:academic_year_choices'])

    def delete(self, *args, **kwargs):
        cache.delete_many([self.CACHE_KEY, 'dscomposn:academic_year_choices'])
        return super().delete(*args, **kwargs)

    @classmethod
    def get(cls):
        obj = cache.get(cls.CACHE_KEY)
        if obj is None:
            obj, _ = cls.objects.get_or_create(pk=1)
            cache.set(cls.CACHE_KEY, obj, 60)
        return obj

    def get_editable_fields(self):
        try:
            data = json.loads(self.editable_profile_fields or '[]')
        except (TypeError, ValueError):
            data = []
        valid = {k for k, _ in PROFILE_EDITABLE_CHOICES}
        return [f for f in data if f in valid]

    def set_editable_fields(self, fields):
        self.editable_profile_fields = json.dumps(list(fields))

    def get_dashboard_summary_fields(self):
        try:
            data = json.loads(self.dashboard_summary_fields or '[]')
        except (TypeError, ValueError):
            data = ['is_muslim', 'food_allergy', 'chronic_disease']
        valid = {k for k, _ in DASHBOARD_SUMMARY_CHOICES}
        return [f for f in data if f in valid]

    def set_dashboard_summary_fields(self, fields):
        self.dashboard_summary_fields = json.dumps(list(fields))

    def get_dashboard_widgets(self):
        raw = self.dashboard_widgets
        if raw is None:
            return list(DEFAULT_DASHBOARD_WIDGETS)
        try:
            data = json.loads(raw)
            if not isinstance(data, list):
                return list(DEFAULT_DASHBOARD_WIDGETS)
        except (TypeError, ValueError):
            return list(DEFAULT_DASHBOARD_WIDGETS)
        valid = {k for k, _ in DASHBOARD_WIDGET_CHOICES}
        seen = set()
        result = []
        for k in data:
            if k in valid and k not in seen:
                seen.add(k)
                result.append(k)
        return result

    def set_dashboard_widgets(self, widgets):
        valid = {k for k, _ in DASHBOARD_WIDGET_CHOICES}
        seen = set()
        cleaned = []
        for k in widgets:
            if k in valid and k not in seen:
                seen.add(k)
                cleaned.append(k)
        self.dashboard_widgets = json.dumps(cleaned)

    def get_academic_years(self):
        """คืนค่ารายการปีการศึกษาในระบบ (เรียงจากปีล่าสุดลงไป) เริ่มตั้งแต่ 2566"""
        years = set(DEFAULT_ACADEMIC_YEARS)
        try:
            data = json.loads(self.academic_years or '[]')
            if isinstance(data, list):
                for item in data:
                    try:
                        y = int(item)
                    except (TypeError, ValueError):
                        continue
                    if MIN_ACADEMIC_YEAR <= y <= MAX_ACADEMIC_YEAR:
                        years.add(y)
        except (TypeError, ValueError):
            pass
        try:
            cur = int(self.current_academic_year)
            if MIN_ACADEMIC_YEAR <= cur <= MAX_ACADEMIC_YEAR:
                years.add(cur)
        except (TypeError, ValueError):
            pass
        return sorted(years, reverse=True)

    def set_academic_years(self, years):
        """บันทึกรายการปีการศึกษา (รวมค่าเริ่มต้น 2566–2569 เสมอ)"""
        cleaned = set(DEFAULT_ACADEMIC_YEARS)
        for item in years or []:
            try:
                y = int(item)
            except (TypeError, ValueError):
                continue
            if MIN_ACADEMIC_YEAR <= y <= MAX_ACADEMIC_YEAR:
                cleaned.add(y)
        self.academic_years = json.dumps(sorted(cleaned))

    def add_academic_year(self, year):
        """เพิ่มปีการศึกษาใหม่เข้าระบบ คืนค่า (added: bool, years_desc: list[int])"""
        try:
            y = int(year)
        except (TypeError, ValueError) as exc:
            raise ValueError('ปีการศึกษาต้องเป็นตัวเลข พ.ศ.') from exc
        if not (MIN_ACADEMIC_YEAR <= y <= MAX_ACADEMIC_YEAR):
            raise ValueError(f'ปีการศึกษาต้องอยู่ระหว่าง {MIN_ACADEMIC_YEAR} ถึง {MAX_ACADEMIC_YEAR}')
        existing = self.get_academic_years()
        if y in existing:
            return False, existing
        self.set_academic_years(existing + [y])
        return True, self.get_academic_years()

    def get_auto_exam_date_camp1(self, year=None):
        """นำวันสอบทั้งทฤษฎี และ ปฏิบัติ จากตารางวันสอบ (TrainingDate) มาลงเป็นข้อความเริ่มต้นแบบอัตโนมัติ"""
        try:
            from food.models import TrainingDate
            from composn.templatetags.composn_tags import date_thai_short
            y = int(year or self.current_academic_year or 2569)
            exam_dates = list(TrainingDate.objects.filter(academic_year=y, event_type=TrainingDate.TYPE_EXAM).order_by('date'))
            if not exam_dates:
                return ''
            theory_td, practice_td = None, None
            for td in exam_dates:
                t = (td.title or '').strip()
                et = getattr(td, 'exam_type', '') or ''
                if et == 'theory' or 'ทฤษฎี' in t:
                    theory_td = td
                elif et == 'practice' or 'ปฏิบัติ' in t or 'ปฎิบัติ' in t:
                    practice_td = td
            if not theory_td and not practice_td and len(exam_dates) >= 2:
                theory_td, practice_td = exam_dates[0], exam_dates[1]
            elif not theory_td and not practice_td and len(exam_dates) == 1:
                return f'สอบทฤษฎีและปฏิบัติ: {date_thai_short(exam_dates[0].date)}'

            parts = []
            if theory_td:
                parts.append(f'สอบทฤษฎี: {date_thai_short(theory_td.date)}')
            if practice_td:
                parts.append(f'สอบปฏิบัติ: {date_thai_short(practice_td.date)}')
            return ', '.join(parts)
        except Exception:
            return ''

    @property
    def effective_exam_date_camp1(self):
        """คืนค่าข้อความวันสอบค่าย 1 (หากเว้นว่าง จะดึงวันสอบทฤษฎีและปฏิบัติจากตารางอบรมอัตโนมัติ)"""
        return self.exam_date_camp1 or self.get_auto_exam_date_camp1()


# =============================================================================
# รอบใช้สิทธิ์
# =============================================================================
CAMP_LEVEL_CHOICES = [(1, 'ค่าย 1'), (2, 'ค่าย 2')]


class RightsRound(models.Model):
    academic_year = models.IntegerField('ปีการศึกษา (พ.ศ.)')
    camp_level = models.PositiveSmallIntegerField('ค่าย', choices=CAMP_LEVEL_CHOICES)
    title = models.CharField('ชื่อรอบ', max_length=150, blank=True)
    is_open = models.BooleanField('เปิดอยู่', default=True, db_index=True)
    scheduled_open_at = models.DateTimeField('เวลาเปิดอัตโนมัติ (ไม่บังคับ)', null=True, blank=True)
    scheduled_close_at = models.DateTimeField('เวลาปิดอัตโนมัติ (ไม่บังคับ)', null=True, blank=True)
    opened_at = models.DateTimeField('เปิดเมื่อ', default=timezone.now)
    opened_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                  related_name='+', verbose_name='เปิดโดย')
    closed_at = models.DateTimeField('ปิดเมื่อ', null=True, blank=True)
    closed_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                  related_name='+', verbose_name='ปิดโดย')
    backup_file = models.CharField('ไฟล์สำรอง', max_length=255, blank=True)

    class Meta:
        verbose_name = 'รอบใช้สิทธิ์'
        verbose_name_plural = 'รอบใช้สิทธิ์'
        ordering = ['-opened_at']

    def __str__(self):
        return f'{self.title or "รอบใช้สิทธิ์"} ค่าย {self.camp_level} ปี {self.academic_year}'

    def is_within_schedule(self, now=None):
        """ตรวจสอบว่าอยู่ในช่วงเวลาเปิด-ปิดที่กำหนดไว้หรือไม่"""
        if not self.is_open:
            return False
        now = now or timezone.now()
        if self.scheduled_open_at and now < self.scheduled_open_at:
            return False
        if self.scheduled_close_at and now > self.scheduled_close_at:
            return False
        return True


class UseRecord(SoftDeleteModel):
    """USE_RECORD — บันทึกการใช้สิทธิ์/สละสิทธิ์ ในแต่ละรอบ"""
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name='use_records',
                             verbose_name='ผู้ใช้')
    rights_round = models.ForeignKey(RightsRound, null=True, blank=True, on_delete=models.SET_NULL,
                                     related_name='use_records', verbose_name='รอบ')
    academic_year = models.IntegerField('ปีการศึกษา', db_index=True)
    camp_level = models.PositiveSmallIntegerField('ค่าย', choices=CAMP_LEVEL_CHOICES)
    rights_code = models.CharField('รหัสใช้สิทธิ์', max_length=20, blank=True)
    grade_level = models.CharField('ระดับชั้นขณะใช้สิทธิ์', max_length=30, blank=True)
    confirmed_at = models.DateTimeField('ยืนยันเมื่อ', null=True, blank=True)
    is_used = models.BooleanField('ใช้สิทธิ์แล้ว', default=False)
    # Snapshot ข้อมูล ณ เวลาที่ใช้สิทธิ์
    camp_code_snapshot = models.CharField(max_length=10, blank=True)
    full_name_snapshot = models.CharField(max_length=255, blank=True)
    school_snapshot = models.CharField(max_length=255, blank=True)
    center_snapshot = models.CharField(max_length=2, blank=True)

    class Meta:
        verbose_name = 'บันทึกการใช้สิทธิ์ (USE_RECORD)'
        verbose_name_plural = 'บันทึกการใช้สิทธิ์ (USE_RECORD)'
        ordering = ['-academic_year', 'camp_level', 'rights_code']
        constraints = [
            # รหัสใช้สิทธิ์ห้ามซ้ำในรายการที่ยังไม่ถูกลบในปีการศึกษาเดียวกัน
            models.UniqueConstraint(
                fields=['academic_year', 'rights_code'],
                condition=Q(is_deleted=False) & ~Q(rights_code=''),
                name='uniq_use_record_rights_code',
            ),
        ]

    def __str__(self):
        return f'{self.rights_code or "สละสิทธิ์"} - {self.user}'

    @property
    def status_text(self):
        return 'ใช้สิทธิ์' if self.is_used else 'สละสิทธิ์'

    def save(self, *args, **kwargs):
        if self.school_snapshot:
            self.school_snapshot = normalize_school_name(self.school_snapshot)
        super().save(*args, **kwargs)


class ArchiveRecord(SoftDeleteModel):
    """ARCHIVE_RECORD — ประวัติการใช้สิทธิ์ถาวร ใช้นับจำนวนครั้ง (สูงสุด 5 ครั้งตลอดชีวิต)"""
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name='archive_records',
                             verbose_name='ผู้ใช้')
    use_record = models.OneToOneField(UseRecord, null=True, blank=True, on_delete=models.SET_NULL,
                                      related_name='archive', verbose_name='อ้างอิง USE_RECORD')
    academic_year = models.IntegerField('ปีการศึกษา')
    camp_level = models.PositiveSmallIntegerField('ค่าย', choices=CAMP_LEVEL_CHOICES)
    rights_code = models.CharField('รหัสใช้สิทธิ์', max_length=20)
    used_at = models.DateTimeField('ใช้สิทธิ์เมื่อ', default=timezone.now)

    class Meta:
        verbose_name = 'ประวัติการใช้สิทธิ์ (ARCHIVE_RECORD)'
        verbose_name_plural = 'ประวัติการใช้สิทธิ์ (ARCHIVE_RECORD)'
        ordering = ['-used_at']

    def __str__(self):
        return f'{self.rights_code} - {self.user}'


# =============================================================================
# คำขอแก้ไขชื่อโรงเรียน
# =============================================================================
class SchoolChangeRequest(models.Model):
    STATUS_PENDING = 'pending'
    STATUS_APPROVED = 'approved'
    STATUS_REJECTED = 'rejected'
    STATUS_CHOICES = [
        (STATUS_PENDING, 'รออนุมัติ'),
        (STATUS_APPROVED, 'อนุมัติแล้ว'),
        (STATUS_REJECTED, 'ไม่อนุมัติ'),
    ]
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='school_requests',
                             verbose_name='ผู้ขอ')
    old_school_name = models.CharField('ชื่อโรงเรียนเดิม', max_length=255, blank=True)
    new_school_name = models.CharField('ชื่อโรงเรียนใหม่', max_length=255)
    reason = models.TextField('เหตุผล', blank=True)
    status = models.CharField('สถานะ', max_length=10, choices=STATUS_CHOICES, default=STATUS_PENDING)
    admin_note = models.TextField('หมายเหตุจาก Admin', blank=True)
    is_cleared = models.BooleanField('เคลียร์สถานะแล้ว', default=False,
                                     help_text='เมื่อเคลียร์แล้ว ผู้ใช้จึงส่งคำขอใหม่ได้')
    reviewed_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                    related_name='+', verbose_name='ตรวจโดย')
    reviewed_at = models.DateTimeField('ตรวจเมื่อ', null=True, blank=True)
    created_at = models.DateTimeField('ส่งคำขอเมื่อ', auto_now_add=True)

    class Meta:
        verbose_name = 'คำขอแก้ไขชื่อโรงเรียน'
        verbose_name_plural = 'คำขอแก้ไขชื่อโรงเรียน'
        ordering = ['is_cleared', '-created_at']

    def __str__(self):
        return f'{self.user} : {self.new_school_name} ({self.get_status_display()})'

    def save(self, *args, **kwargs):
        if self.old_school_name:
            self.old_school_name = normalize_school_name(self.old_school_name)
        if self.new_school_name:
            self.new_school_name = normalize_school_name(self.new_school_name)
        super().save(*args, **kwargs)


# =============================================================================
# Audit Log
# =============================================================================
class AuditLog(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                             related_name='audit_logs', verbose_name='ผู้กระทำ')
    action = models.CharField('การกระทำ', max_length=50, db_index=True)
    target_model = models.CharField('ตาราง', max_length=100, blank=True)
    target_id = models.IntegerField('ID เป้าหมาย', null=True, blank=True)
    old_value = models.TextField('ค่าเดิม', blank=True)
    new_value = models.TextField('ค่าใหม่', blank=True)
    ip_address = models.CharField('IP', max_length=50, blank=True)
    created_at = models.DateTimeField('เวลา', auto_now_add=True, db_index=True)

    class Meta:
        verbose_name = 'Audit Log'
        verbose_name_plural = 'Audit Log'
        ordering = ['-created_at']

    def __str__(self):
        return f'[{self.created_at:%Y-%m-%d %H:%M}] {self.user} {self.action}'
