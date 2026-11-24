"""
Models ของแอป Food
- ระบบเวลาเรียน (Attendance) + เหตุผลการขาด + สีเตือนภัย
- ระบบคะแนน (ScoreCategory / Score)
- ระบบสั่งอาหาร (FoodOrder / FoodMenu / FoodSelection) ย้ายมาจากระบบ food เดิม
"""
import colorsys
import random
import re
import secrets
from decimal import Decimal

from django.conf import settings
from django.db import models
from django.db.models import Sum
from django.utils import timezone

from composn.models import SoftDeleteModel, SystemSetting


def current_year():
    try:
        return SystemSetting.get().current_academic_year
    except Exception:
        return timezone.localdate().year + 543


# =============================================================================
# Attendance
# =============================================================================
class AbsenceReason(models.Model):
    name = models.CharField('เหตุผล', max_length=100, unique=True)
    order = models.PositiveIntegerField('ลำดับ', default=0)
    is_active = models.BooleanField('ใช้งาน', default=True)

    class Meta:
        verbose_name = 'เหตุผลการขาดเรียน'
        verbose_name_plural = 'เหตุผลการขาดเรียน'
        ordering = ['order', 'id']

    def __str__(self):
        return self.name


class WarningColor(models.Model):
    """สีเตือนภัย: ใช้สีของระดับที่ min_absences สูงสุดที่ไม่เกินจำนวนครั้งที่ขาด"""
    name = models.CharField('ชื่อระดับ', max_length=50)
    color = models.CharField('สี (HEX)', max_length=7, default='#ffc107')
    min_absences = models.PositiveIntegerField('เริ่มเมื่อขาดตั้งแต่ (ครั้ง)', default=1)
    order = models.PositiveIntegerField('ลำดับ', default=0)

    class Meta:
        verbose_name = 'สีเตือนภัย'
        verbose_name_plural = 'สีเตือนภัย'
        ordering = ['min_absences', 'order']

    def __str__(self):
        return f'{self.name} ({self.color}) >= {self.min_absences}'


class Attendance(SoftDeleteModel):
    PERIOD_MORNING = 'morning'
    PERIOD_AFTERNOON = 'afternoon'
    PERIOD_CHOICES = [(PERIOD_MORNING, 'เช้า'), (PERIOD_AFTERNOON, 'บ่าย')]
    STATUS_PRESENT = 'present'
    STATUS_ABSENT = 'absent'
    STATUS_SICK = 'sick'
    STATUS_PERSONAL = 'personal'
    STATUS_LATE = 'late'
    STATUS_CHOICES = [
        (STATUS_PRESENT, 'เข้าอบรม'),
        (STATUS_ABSENT, 'ขาดอบรม'),
        (STATUS_SICK, 'ลาป่วย'),
        (STATUS_PERSONAL, 'ลากิจ'),
    ]
    FIELD_STATUS_CHOICES = STATUS_CHOICES + [(STATUS_LATE, 'มาสาย')]
    ABSENT_STATUSES = (STATUS_ABSENT, STATUS_SICK, STATUS_PERSONAL)
    ALL_VALID_STATUSES = (STATUS_PRESENT, STATUS_ABSENT, STATUS_SICK, STATUS_PERSONAL, STATUS_LATE)

    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name='attendances',
                             verbose_name='นักเรียน')
    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    date = models.DateField('วันที่')
    session_period = models.CharField('ช่วง', max_length=10, choices=PERIOD_CHOICES)
    status = models.CharField('สถานะ', max_length=10, choices=FIELD_STATUS_CHOICES, default=STATUS_PRESENT)
    reason = models.CharField('เหตุผล', max_length=100, blank=True)
    note = models.TextField('หมายเหตุ', blank=True)
    recorded_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                    related_name='+', verbose_name='บันทึกโดย')
    recorded_at = models.DateTimeField('บันทึกเมื่อ', auto_now=True)

    class Meta:
        verbose_name = 'เวลาเรียน'
        verbose_name_plural = 'เวลาเรียน'
        ordering = ['-date', 'session_period', 'user__camp1_code']
        constraints = [
            models.UniqueConstraint(fields=['user', 'date', 'session_period'], condition=models.Q(is_deleted=False),
                                    name='uniq_attendance_alive'),
        ]

    def __str__(self):
        return f'{self.user} {self.date} {self.get_session_period_display()} {self.get_status_display()}'


def to_ce_date(d):
    """แปลงวันที่ที่เก็บเป็นปี พ.ศ. (> 2400) ให้เป็น ค.ศ. เพื่อเปรียบเทียบกับวันที่ปัจจุบัน"""
    if d and d.year > 2400:
        try:
            return d.replace(year=d.year - 543)
        except ValueError:
            return d
    return d


def is_future_attendance_date(d, today=None):
    """คืนค่า True หากวันที่ดังกล่าวยังไม่ถึง (มากกว่าวันปัจจุบัน)"""
    today = today or timezone.localdate()
    return to_ce_date(d) > today


def non_training_dates_for_year(academic_year=None):
    """คืน set ของวันที่ที่เป็นวันสอบหรือวันหยุดในปีการศึกษานั้น (ไม่นับเป็นวันอบรม)"""
    year = academic_year or current_year()
    tds = TrainingDate.objects.filter(
        academic_year=year,
        event_type__in=(TrainingDate.TYPE_EXAM, TrainingDate.TYPE_HOLIDAY),
    ).values_list('date', flat=True)
    res = set()
    for d in tds:
        res.add(d)
        res.add(to_ce_date(d))
    return res


def absence_count(user, academic_year=None):
    """นับจำนวนภาคการอบรมที่ขาดเรียนในปีการศึกษา (นับเฉพาะวันอบรม ไม่นับวันสอบและวันหยุด)"""
    year = academic_year or current_year()
    qs = Attendance.objects.filter(
        user=user, academic_year=year, status__in=Attendance.ABSENT_STATUSES
    )
    excl = non_training_dates_for_year(year)
    if excl:
        qs = qs.exclude(date__in=excl)
    return qs.count()


def warning_for(count):
    """คืนค่า WarningColor ที่ตรงกับจำนวนครั้งที่ขาด (หรือ None)"""
    if count <= 0:
        return None
    return WarningColor.objects.filter(min_absences__lte=count).order_by('-min_absences').first()


def absence_status(user, academic_year=None, count=None):
    setting = SystemSetting.get()
    if count is None:
        count = absence_count(user, academic_year)
    max_abs = setting.max_absences
    warn = warning_for(count)
    text = f'ขาดอบรม {count} จาก {max_abs} ภาคเวลา'
    near = max_abs > 1 and count == max_abs - 1
    if count >= max_abs:
        text += ' (ครบจำนวนที่ขาดอบรมสูงสุดแล้ว)'
    elif near:
        text += ' (ใกล้ถึงจำนวนที่ขาดอบรมสูงสุด)'
    return {
        'count': count, 'max': max_abs, 'text': text, 'near_max': near, 'reached': count >= max_abs,
        'color': warn.color if warn else '', 'level': warn.name if warn else '',
    }


# =============================================================================
# Score & Training Schedule
# =============================================================================
class TrainingDate(models.Model):
    """ตารางวันที่อบรม วันสอบ และวันหยุด"""
    TYPE_TRAINING = 'training'
    TYPE_EXAM = 'exam'
    TYPE_HOLIDAY = 'holiday'
    TYPE_CHOICES = [
        (TYPE_TRAINING, 'วันอบรม'),
        (TYPE_EXAM, 'วันสอบ'),
        (TYPE_HOLIDAY, 'วันหยุด'),
    ]
    EXAM_BOTH = 'both'
    EXAM_MORNING = 'morning'
    EXAM_AFTERNOON = 'afternoon'
    EXAM_PERIOD_CHOICES = [
        (EXAM_BOTH, 'ทั้งวัน (เช้า + บ่าย)'),
        (EXAM_MORNING, 'เฉพาะภาคเช้า'),
        (EXAM_AFTERNOON, 'เฉพาะภาคบ่าย'),
    ]
    EXAM_TYPE_AUTO = ''
    EXAM_TYPE_THEORY = 'theory'
    EXAM_TYPE_PRACTICE = 'practice'
    EXAM_TYPE_GENERAL = 'general'
    EXAM_TYPE_CHOICES = [
        (EXAM_TYPE_AUTO, 'ตามรายละเอียด'),
        (EXAM_TYPE_THEORY, 'สอบทฤษฎี'),
        (EXAM_TYPE_PRACTICE, 'สอบปฏิบัติ'),
        (EXAM_TYPE_GENERAL, 'สอบทั่วไป'),
    ]
    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    date = models.DateField('วันที่')
    event_type = models.CharField('ประเภท', max_length=15, choices=TYPE_CHOICES, default=TYPE_TRAINING)
    exam_period = models.CharField('ช่วงเวลาสอบ', max_length=15, choices=EXAM_PERIOD_CHOICES, default=EXAM_BOTH)
    exam_type = models.CharField('ประเภทการสอบ', max_length=20, choices=EXAM_TYPE_CHOICES, blank=True, default='')
    camp_level = models.PositiveSmallIntegerField('ค่าย', default=1, choices=[(1, 'ค่าย 1'), (2, 'ค่าย 2')])
    title = models.CharField('หัวข้อ / รายละเอียด', max_length=200, blank=True)

    @property
    def effective_exam_type_display(self):
        if self.exam_type == self.EXAM_TYPE_THEORY:
            return 'สอบทฤษฎี'
        if self.exam_type == self.EXAM_TYPE_PRACTICE:
            return 'สอบปฏิบัติ'
        if self.exam_type == self.EXAM_TYPE_GENERAL:
            return 'สอบทั่วไป'
        t = (self.title or '').strip()
        if 'ปฏิบัติ' in t or 'ปฎิบัติ' in t:
            return 'สอบปฏิบัติ'
        if 'ทฤษฎี' in t:
            return 'สอบทฤษฎี'
        return t or 'วันสอบ'

    @property
    def seating_chart_subtitle(self):
        if self.exam_type == self.EXAM_TYPE_THEORY:
            return 'แผนผังที่นั่งสอบทฤษฎี'
        if self.exam_type == self.EXAM_TYPE_PRACTICE:
            return 'แผนผังที่นั่งสอบปฏิบัติ'
        if self.exam_type == self.EXAM_TYPE_GENERAL:
            return 'แผนผังที่นั่งสอบ'
        t = (self.title or '').strip()
        if 'ปฏิบัติ' in t or 'ปฎิบัติ' in t:
            return 'แผนผังที่นั่งสอบปฏิบัติ'
        if 'ทฤษฎี' in t:
            return 'แผนผังที่นั่งสอบทฤษฎี'
        if t.startswith('สอบ'):
            return f'แผนผังที่นั่ง{t}'
        if t:
            return f'แผนผังที่นั่งสอบ ({t})'
        return 'แผนผังที่นั่งสอบ'

    def active_periods(self):
        """ช่วงเวลาที่เปิดให้บันทึกเวลาในวันนี้"""
        if self.event_type == self.TYPE_HOLIDAY:
            return []
        if self.event_type == self.TYPE_EXAM:
            if self.exam_period == self.EXAM_MORNING:
                return [Attendance.PERIOD_MORNING]
            if self.exam_period == self.EXAM_AFTERNOON:
                return [Attendance.PERIOD_AFTERNOON]
            return [Attendance.PERIOD_MORNING, Attendance.PERIOD_AFTERNOON]
        return [Attendance.PERIOD_MORNING, Attendance.PERIOD_AFTERNOON]

    class Meta:
        verbose_name = 'วันที่อบรม/วันสอบ'
        verbose_name_plural = 'วันที่อบรม/วันสอบ'
        ordering = ['date', 'id']

    def __str__(self):
        return f'{self.date} ({self.get_event_type_display()}) {self.title}'


class ScoreCategory(models.Model):
    GROUP_COLLECT = 'คะแนนเก็บ'
    GROUP_THEORY = 'คะแนนสอบทฤษฎี'
    GROUP_PRACTICE = 'คะแนนสอบปฏิบัติ'
    GROUP_CHOICES = [
        (GROUP_COLLECT, 'คะแนนเก็บ'),
        (GROUP_THEORY, 'คะแนนสอบทฤษฎี'),
        (GROUP_PRACTICE, 'คะแนนสอบปฏิบัติ'),
    ]
    BOX_COLOR_PALETTE = [
        '#0d6efd', '#198754', '#6f42c1', '#d63384', '#fd7e14', '#0891b2', '#dc3545',
        '#0d9488', '#4f46e5', '#b45309', '#be123c', '#0369a1', '#4d7c0f', '#7e22ce',
    ]

    parent = models.ForeignKey('self', null=True, blank=True, on_delete=models.CASCADE,
                               related_name='subcategories', verbose_name='หัวข้อใหญ่')
    group_type = models.CharField('ประเภทหัวข้อใหญ่', max_length=50, choices=GROUP_CHOICES,
                                  default=GROUP_COLLECT, blank=True)
    name = models.CharField('ชื่อหมวด', max_length=150)
    weight = models.DecimalField('น้ำหนัก', max_digits=7, decimal_places=2, default=Decimal('100'))
    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    order = models.PositiveIntegerField('ลำดับ', default=0)
    default_max_score = models.DecimalField('คะแนนเต็มเริ่มต้น', max_digits=7, decimal_places=2,
                                            default=Decimal('100'))
    box_color = models.CharField('สีกล่องหัวข้อ', max_length=20, blank=True, default='')
    is_visible = models.BooleanField('แสดงผลคะแนนส่วนนี้', default=True)
    is_open = models.BooleanField('เปิดให้กรอกคะแนน', default=True)
    open_at = models.DateTimeField('เวลาเริ่มเปิดกรอก (ไม่บังคับ)', null=True, blank=True)
    close_at = models.DateTimeField('เวลาปิดกรอก (เว้นว่าง = ไม่ปิด)', null=True, blank=True)

    @classmethod
    def pick_random_unused_color(cls, academic_year):
        """สุ่มเลือกสีสำหรับกล่องหัวข้อใหญ่โดยไม่ให้ซ้ำกับหัวข้อใหญ่ที่มีอยู่แล้วในปีการศึกษาเดียวกัน"""
        used = {
            c.strip().lower()
            for c in cls.objects.filter(academic_year=academic_year, parent__isnull=True)
            .exclude(box_color='')
            .values_list('box_color', flat=True)
        }
        avail = [c for c in cls.BOX_COLOR_PALETTE if c.lower() not in used]
        if avail:
            return random.choice(avail)
        for _ in range(40):
            h = random.random()
            l_val = random.uniform(0.36, 0.46)
            s_val = random.uniform(0.65, 0.85)
            r, g, b = colorsys.hls_to_rgb(h, l_val, s_val)
            candidate = '#{:02x}{:02x}{:02x}'.format(int(r * 255), int(g * 255), int(b * 255))
            if candidate.lower() not in used:
                return candidate
        return '#0d6efd'

    def save(self, *args, **kwargs):
        if not self.parent_id and not (self.box_color or '').strip():
            self.box_color = self.pick_random_unused_color(self.academic_year)
        super().save(*args, **kwargs)

    class Meta:
        verbose_name = 'หมวดคะแนน'
        verbose_name_plural = 'หมวดคะแนน'
        ordering = ['-academic_year', 'order', 'id']

    def __str__(self):
        if self.parent_id:
            return f'{self.parent.name} › {self.name} ({self.academic_year})'
        return f'{self.name} ({self.academic_year}) น้ำหนัก {self.weight}'

    def is_currently_open(self, now=None):
        """ตรวจสอบว่าหมวดคะแนนนี้เปิดให้กรอกคะแนนอยู่ในขณะนี้หรือไม่ (ถ้าเป็นหัวข้อย่อย ใช้เวลาเปิด-ปิดของหัวข้อใหญ่)"""
        if self.parent_id:
            if not self.is_open:
                return False
            return self.parent.is_currently_open(now=now)
        if not self.is_open:
            return False
        now = now or timezone.now()
        if self.open_at and now < self.open_at:
            return False
        if self.close_at and now > self.close_at:
            return False
        return True


class Score(SoftDeleteModel):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name='scores',
                             verbose_name='นักเรียน')
    category = models.ForeignKey(ScoreCategory, on_delete=models.PROTECT, related_name='scores',
                                 verbose_name='หมวด')
    score = models.DecimalField('คะแนน', max_digits=7, decimal_places=2)
    max_score = models.DecimalField('คะแนนเต็ม', max_digits=7, decimal_places=2, default=Decimal('100'))
    exam_no = models.PositiveIntegerField('ครั้งที่สอบ', default=1)
    recorded_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                    related_name='+', verbose_name='บันทึกโดย')
    recorded_at = models.DateTimeField('บันทึกเมื่อ', auto_now=True)

    class Meta:
        verbose_name = 'คะแนน'
        verbose_name_plural = 'คะแนน'
        ordering = ['category__order', 'exam_no']
        constraints = [
            models.UniqueConstraint(fields=['user', 'category', 'exam_no'], condition=models.Q(is_deleted=False),
                                    name='uniq_score_alive'),
        ]

    def __str__(self):
        return f'{self.user} {self.category.name} #{self.exam_no}: {self.score}/{self.max_score}'


def total_score(user, academic_year=None, only_visible=False):
    """คะแนนรวม = ผลรวมของ (score/max_score * weight) ของแต่ละหัวข้อใหญ่ (หรือหัวข้อที่ไม่มีหัวข้อใหญ่)
    - กรณีหัวข้อใหญ่มีหัวข้อย่อย (subcategories):
      - ถ้าหัวข้อย่อยมีน้ำหนัก (weight > 0) จะรวมตามน้ำหนักของแต่ละหัวข้อย่อย
      - ถ้าหัวข้อย่อยทุกอันมี weight == 0 จะคิดจาก ผลรวมคะแนนที่ได้ของทุกหัวข้อย่อย / ผลรวมคะแนนเต็ม * น้ำหนักหัวข้อใหญ่
    - กรณีไม่มีหัวข้อย่อย: คิดจากค่าเฉลี่ยของ (score/max_score) ในหมวดนั้น * weight
    - หาก only_visible=True จะคำนวณและแสดงเฉพาะหมวดคะแนนที่เปิดแสดงผล (is_visible=True)"""
    year = academic_year or current_year()
    total = Decimal('0')
    breakdown = []
    main_qs = ScoreCategory.objects.filter(academic_year=year, parent__isnull=True).order_by('order', 'id')
    if only_visible:
        main_qs = main_qs.filter(is_visible=True)

    for cat in main_qs:
        sub_qs = cat.subcategories.all().order_by('order', 'id')
        if only_visible:
            sub_qs = sub_qs.filter(is_visible=True)
        subs = list(sub_qs)

        if subs:
            sub_items = []
            earned_sum = Decimal('0')
            max_sum = Decimal('0')
            has_any_score = False

            for idx_sub, sc in enumerate(subs, 1):
                sc_scores = list(Score.objects.filter(user=user, category=sc))
                sub_max = sc.default_max_score or sc.weight or Decimal('100')
                if sc_scores:
                    has_any_score = True
                    s_earned = sum((s.score for s in sc_scores), Decimal('0')) / len(sc_scores)
                    s_max = sum(((s.max_score or sub_max) for s in sc_scores), Decimal('0')) / len(sc_scores)
                    s_ratio = (s_earned / s_max) if s_max else Decimal('0')
                    earned_sum += s_earned
                    max_sum += s_max
                    s_weighted = s_earned
                else:
                    s_ratio = None
                    s_weighted = None
                sub_items.append({
                    'index': idx_sub,
                    'category': sc,
                    'scores': sc_scores,
                    'ratio': s_ratio,
                    'weighted': s_weighted.quantize(Decimal('0.01')) if s_weighted is not None else None,
                })

            if has_any_score:
                ratio = (earned_sum / max_sum) if max_sum else Decimal('0')
                weighted = ratio * cat.weight
                total += weighted
            else:
                ratio = None
                weighted = None

            breakdown.append({
                'category': cat,
                'scores': [],
                'sub_breakdown': sub_items,
                'earned_sum': earned_sum.quantize(Decimal('0.01')) if has_any_score else None,
                'max_sum': max_sum.quantize(Decimal('0.01')) if has_any_score else None,
                'ratio': ratio,
                'weighted': weighted.quantize(Decimal('0.01')) if weighted is not None else None,
            })
        else:
            scores = list(Score.objects.filter(user=user, category=cat))
            if scores:
                ratios = [(s.score / s.max_score) if s.max_score else Decimal('0') for s in scores]
                ratio = sum(ratios) / len(ratios)
                weighted = ratio * cat.weight
                total += weighted
            else:
                ratio = None
                weighted = None
            breakdown.append({
                'category': cat,
                'scores': scores,
                'sub_breakdown': [],
                'ratio': ratio,
                'weighted': weighted.quantize(Decimal('0.01')) if weighted is not None else None,
            })
    return total.quantize(Decimal('0.01')), breakdown


# =============================================================================
# อนุมัติประกาศผลคะแนน / ผลการคัดเลือกเข้าค่าย 2 (รายปีการศึกษา)
# =============================================================================
class ScoreApproval(models.Model):
    """Admin อนุมัติประกาศผลคะแนนของปีการศึกษา (เก็บเกณฑ์ ณ เวลาอนุมัติ)"""
    academic_year = models.IntegerField('ปีการศึกษา', unique=True)
    is_approved = models.BooleanField('อนุมัติประกาศผลแล้ว', default=False)
    passing_score = models.DecimalField('คะแนนเกณฑ์ผ่าน', max_digits=7, decimal_places=2, default=Decimal('0'))
    max_qualified_count = models.PositiveIntegerField('จำนวนผู้ผ่านสูงสุด', default=0)
    approved_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                    related_name='+', verbose_name='อนุมัติโดย')
    approved_at = models.DateTimeField('อนุมัติเมื่อ', null=True, blank=True)

    class Meta:
        verbose_name = 'การอนุมัติประกาศผลคะแนน'
        verbose_name_plural = 'การอนุมัติประกาศผลคะแนน'
        ordering = ['-academic_year']

    def __str__(self):
        return f'ปี {self.academic_year}: {"อนุมัติแล้ว" if self.is_approved else "ยังไม่อนุมัติ"}'

    @classmethod
    def is_year_approved(cls, year):
        if year is None or year == '':
            return False
        try:
            year_int = int(year)
        except (TypeError, ValueError):
            return False
        return cls.objects.filter(academic_year=year_int, is_approved=True).exists()


def year_locked_message(year):
    """ข้อความแจ้งเตือนเมื่อปีการศึกษาถูกอนุมัติประกาศผลแล้วและถูกล็อกไม่ให้แก้ไข"""
    return (
        f'ปีการศึกษา {year} อนุมัติประกาศผลแล้ว ไม่สามารถบันทึกหรือแก้ไขข้อมูลที่มีผลต่อผลการคัดเลือกได้ '
        f'หากต้องการแก้ไขกรุณายกเลิกประกาศผลก่อน'
    )


def approved_years_set():
    """คืนค่า set ของปีการศึกษา (int) ที่อนุมัติประกาศผลแล้ว"""
    return set(ScoreApproval.objects.filter(is_approved=True).values_list('academic_year', flat=True))


def camp2_code_year(code):
    """ดึงปีการศึกษา (พ.ศ.) จากรหัสค่าย 2 เช่น '69KM01' -> 2569 (หรือ None หากไม่ตรงรูปแบบ)"""
    m = re.match(r'^(\d{2})KM', str(code or '').strip(), re.IGNORECASE)
    return (2500 + int(m.group(1))) if m else None


def user_approved_years(user):
    """คืนรายการปีการศึกษาที่อนุมัติประกาศผลแล้วและเกี่ยวข้องกับนักเรียนคนนี้ (ปีรุ่น / ปีค่าย 2 / ปีที่ใช้สิทธิ์ / มีผลคัดเลือก)"""
    from composn.models import UseRecord, User
    if not user or getattr(user, 'role', '') != User.ROLE_MEMBER:
        return []
    appr_years = approved_years_set()
    if not appr_years:
        return []
    candidate_years = set()
    for y in (getattr(user, 'academic_year', None), getattr(user, 'base_academic_year', None),
              camp2_code_year(getattr(user, 'camp2_code', None))):
        if y:
            try:
                candidate_years.add(int(y))
            except (TypeError, ValueError):
                pass
    candidate_years.update(UseRecord.objects.filter(user=user, is_used=True).values_list('academic_year', flat=True))
    candidate_years.update(QualificationResult.objects.filter(user=user).values_list('academic_year', flat=True))
    return sorted(candidate_years & appr_years)


class QualificationResult(models.Model):
    """ผลคะแนนรวม/อันดับ/ผลการคัดเลือกเข้าค่าย 2 ของนักเรียนแต่ละคน (บันทึก ณ เวลาที่ Admin อนุมัติ)"""
    academic_year = models.IntegerField('ปีการศึกษา', db_index=True)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='qualification_results',
                             verbose_name='นักเรียน')
    total = models.DecimalField('คะแนนรวม', max_digits=8, decimal_places=2, default=Decimal('0'))
    rank = models.PositiveIntegerField('อันดับ', default=0)
    passed = models.BooleanField('ผ่านการคัดเลือกเข้าค่าย 2', default=False)
    created_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'ผลการคัดเลือกเข้าค่าย 2'
        verbose_name_plural = 'ผลการคัดเลือกเข้าค่าย 2'
        ordering = ['-academic_year', 'rank']
        constraints = [
            models.UniqueConstraint(fields=['academic_year', 'user'], name='uniq_qualification_year_user'),
        ]

    def __str__(self):
        return f'{self.user} ปี {self.academic_year}: {self.total} ({"ผ่าน" if self.passed else "ไม่ผ่าน"})'


def is_current_camp_student(user, year):
    """ตรวจสอบว่าเป็นนักเรียนปัจจุบันของปีการศึกษานั้นหรือไม่ (ไม่ใช่นักเรียนใช้สิทธิ์)"""
    from composn.models import UseRecord, User
    if not user or getattr(user, 'role', '') != User.ROLE_MEMBER:
        return False
    year_int = int(year)
    # ถ้ามีประวัติการใช้สิทธิ์ในปีนี้ = เป็นนักเรียนใช้สิทธิ์
    if UseRecord.objects.filter(user=user, academic_year=year_int, is_used=True).exists():
        return False
    # ตรวจสอบว่าเป็นรุ่นปีนั้น
    yy = f'{year_int % 100:02d}'
    code = (getattr(user, 'camp1_code', '') or '').upper()
    return bool(code.startswith(yy) or (getattr(user, 'academic_year', None) == year_int) or (getattr(user, 'base_academic_year', None) == year_int))


def score_members_for_year(year):
    """นักเรียนที่อยู่ในการคิดคะแนนของปีการศึกษา: นักเรียนรุ่นปีนั้น + ผู้ใช้สิทธิ์ปีนั้น (เรียงนักเรียนปัจจุบันก่อน ตามด้วยนักเรียนใช้สิทธิ์)"""
    from composn.models import UseRecord, User
    year_int = int(year)
    yy = f'{year_int % 100:02d}'
    used_ids = set(UseRecord.objects.filter(academic_year=year_int, is_used=True).values_list('user_id', flat=True))
    is_current_q = ~models.Q(pk__in=used_ids) & (models.Q(camp1_code__istartswith=yy) | models.Q(academic_year=year_int))
    qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        (models.Q(camp1_code__istartswith=yy) | models.Q(academic_year=year_int)) | models.Q(pk__in=used_ids)
    ).annotate(
        cohort_order=models.Case(
            models.When(is_current_q, then=models.Value(0)),
            default=models.Value(1),
            output_field=models.IntegerField(),
        )
    ).order_by('cohort_order', 'camp1_code')
    return qs, used_ids


def compute_year_ranking(year, members=None, passing_score=None, max_qualified_count=None, only_visible=True):
    """คำนวณคะแนนรวม อันดับ และผลผ่านการคัดเลือกของปีการศึกษา"""
    setting = SystemSetting.get()
    if members is None:
        members, _ = score_members_for_year(year)
    ps = setting.passing_score if passing_score is None else passing_score
    mq = setting.max_qualified_count if max_qualified_count is None else max_qualified_count
    rows = []
    for m in members:
        total, breakdown = total_score(m, year, only_visible=only_visible)
        rows.append({'user': m, 'total': total, 'breakdown': breakdown})
    rows.sort(key=lambda r: (-r['total'], r['user'].camp1_code))
    for rank, r in enumerate(rows, start=1):
        r['rank'] = rank
        if mq and mq > 0:
            passed = (rank <= mq) and (r['total'] >= ps if ps and ps > 0 else True)
        else:
            passed = r['total'] >= ps
        r['passed'] = passed
    return rows


def _natural_code_key(code):
    return [int(p) if p.isdigit() else p for p in re.split(r'(\d+)', (code or '').upper())]


def sort_announcement_rows(rows, year, used_records=None):
    """เรียงลำดับรายชื่อสำหรับเอกสารประกาศผล:
    1) นักเรียนปัจจุบัน เรียงตามรหัสจากน้อยไปมาก (เช่น 69DS01, 69DS02, ...)
    2) นักเรียนใช้สิทธิ์ เรียงตามรหัสจากน้อยไปมาก (เช่น 67DS15, 68DS02, ...)
    """
    from composn.models import UseRecord
    year_int = int(year)
    if used_records is None:
        used_records = {
            ur.user_id: ur.rights_code
            for ur in UseRecord.objects.filter(academic_year=year_int, is_used=True)
        }
    yy = f'{year_int % 100:02d}'

    def _key(r):
        u = r['user']
        is_rights = u.pk in used_records
        code = (used_records.get(u.pk) if is_rights else None) or u.camp1_code or ''
        is_current = (not is_rights) and (
            code.upper().startswith(yy)
            or getattr(u, 'academic_year', None) == year_int
            or getattr(u, 'base_academic_year', None) == year_int
        )
        group = 0 if is_current else 1
        return (group, _natural_code_key(code), _natural_code_key(u.camp1_code or ''))

    return sorted(rows, key=_key)


def approve_year_scores(year, admin):
    """Admin อนุมัติประกาศผลคะแนนปีการศึกษา: บันทึกผลคะแนนรวม/อันดับ/ผลการคัดเลือกเข้าค่าย 2 ของทุกคน
    และกำหนดรหัสค่าย 2 ให้กับนักเรียนที่ผ่านการคัดเลือกเข้าค่าย 2 ตามที่ระบบทำประกาศ (เช่น 69KM01, 69KM02, ...)
    """
    from django.db import transaction
    from composn.models import User
    setting = SystemSetting.get()
    year = int(year)
    rows = compute_year_ranking(year)
    prefix = f'{str(year)[-2:]}KM'
    with transaction.atomic():
        appr, _ = ScoreApproval.objects.get_or_create(academic_year=year)
        appr.is_approved = True
        appr.passing_score = setting.passing_score
        appr.max_qualified_count = setting.max_qualified_count
        appr.approved_by = admin
        appr.approved_at = timezone.now()
        appr.save()
        QualificationResult.objects.filter(academic_year=year).delete()
        QualificationResult.objects.bulk_create([
            QualificationResult(academic_year=year, user=r['user'], total=r['total'], rank=r['rank'],
                                passed=r['passed'])
            for r in rows
        ])

        # เคลียร์รหัสค่าย 2 ของปีนี้ก่อน เพื่อป้องกัน unique constraint ชนกันกรณีสลับอันดับ
        User.objects.filter(camp2_code__startswith=prefix).update(camp2_code=None)

        # มอบรหัสค่าย 2 ให้กับนักเรียนที่ผ่านการคัดเลือกเข้าค่าย 2 ตามลำดับประกาศผล
        passed_rows = sort_announcement_rows([r for r in rows if r['passed']], year)
        for idx, r in enumerate(passed_rows, start=1):
            km_code = f'{prefix}{idx:02d}'
            u = r['user']
            u.camp2_code = km_code
            u.save(update_fields=['camp2_code'])
            r['camp2_code'] = km_code

    return appr, rows


def unapprove_year_scores(year):
    from django.db import transaction
    from composn.models import User
    year_int = int(year)
    prefix = f'{str(year_int)[-2:]}KM'
    with transaction.atomic():
        User.objects.filter(camp2_code__startswith=prefix).update(camp2_code=None)
        ScoreApproval.objects.filter(academic_year=year_int).update(is_approved=False)


def member_year_results(user):
    """ข้อมูลคะแนน/เวลาเรียน/ผลการคัดเลือกของนักเรียน แยกตามปีการศึกษา (ปีล่าสุดอยู่บนสุด)
    - ปีที่เข้าอบรม (ปีรุ่นของนักเรียน): แสดงเวลาเรียน + คะแนน (เฉพาะหมวดที่เปิดแสดงผล) + ผลคัดเลือกเมื่อ Admin อนุมัติ
    - ปีที่ใช้สิทธิ์: ไม่มีเวลาเรียน แสดงคะแนนและผลคัดเลือกต่อเมื่อ ใช้สิทธิ์แล้ว + ปิดระบบใช้สิทธิ์แล้ว + Admin อนุมัติแล้ว
    """
    from composn.models import RightsRound, UseRecord
    cohort_year = user.base_academic_year
    camp2_year = None
    if user.camp2_code:
        m_c2 = re.match(r'^(\d{2})', user.camp2_code)
        camp2_year = (2500 + int(m_c2.group(1))) if m_c2 else None
    used_records = {r.academic_year: r for r in UseRecord.objects.filter(user=user, is_used=True)
                    .select_related('rights_round')}
    years = set(used_records.keys())
    if cohort_year:
        years.add(int(cohort_year))

    results = []
    for year in sorted(years, reverse=True):
        is_training_year = bool(cohort_year) and int(year) == int(cohort_year)
        rec = used_records.get(year)
        approved = ScoreApproval.is_year_approved(year)
        item = {
            'year': year,
            'is_training_year': is_training_year,
            'is_rights_year': not is_training_year,
            'use_record': rec,
            'approved': approved,
            'show_scores': False,
            'show_result': False,
            'show_exam_status': False,
            'exam_status_camp': None,
            'exam_status_text': '',
            'pending_reasons': [],
            'total': None,
            'breakdown': [],
            'result': None,
            'absence': None,
            'attendances': [],
            'att_days': [],
            'att_summary': None,
        }
        if is_training_year:
            item['show_scores'] = True
            item['absence'] = absence_status(user, academic_year=year)
            tds = list(TrainingDate.objects.filter(academic_year=year).order_by('date', 'id'))
            td_by_ce = {to_ce_date(td.date): td for td in tds}
            non_training_ce = {
                to_ce_date(td.date) for td in tds
                if td.event_type in (TrainingDate.TYPE_EXAM, TrainingDate.TYPE_HOLIDAY)
            }
            att_list = list(Attendance.objects.filter(user=user, academic_year=year).order_by('date', 'session_period'))
            item['attendances'] = att_list
            # นับสถิติเวลาเรียนเฉพาะวันอบรม (วันสอบ และ วันหยุด ไม่นับเป็นวันอบรม)
            train_atts = [a for a in att_list if to_ce_date(a.date) not in non_training_ce]
            cnt_p = sum(1 for a in train_atts if a.status == Attendance.STATUS_PRESENT)
            cnt_a = sum(1 for a in train_atts if a.status == Attendance.STATUS_ABSENT)
            cnt_s = sum(1 for a in train_atts if a.status == Attendance.STATUS_SICK)
            cnt_per = sum(1 for a in train_atts if a.status == Attendance.STATUS_PERSONAL)
            tot_sess = len(train_atts)
            training_days = sum(1 for td in tds if td.event_type == TrainingDate.TYPE_TRAINING)
            item['att_summary'] = {
                'present': cnt_p,
                'absent': cnt_a,
                'sick': cnt_s,
                'personal': cnt_per,
                'total_absent': cnt_a + cnt_s + cnt_per,
                'total_sessions': tot_sess,
                'display_total_sessions': tot_sess if tot_sess > 0 else (training_days * 2),
                'missed_exams': 0,
                'training_days': training_days,
                'exam_days': sum(1 for td in tds if td.event_type == TrainingDate.TYPE_EXAM),
                'holiday_days': sum(1 for td in tds if td.event_type == TrainingDate.TYPE_HOLIDAY),
                'percent': round(cnt_p * 100.0 / tot_sess, 1) if tot_sess else 100.0,
                'present_pct': round(cnt_p * 100.0 / tot_sess, 1) if tot_sess else 100.0,
                'absent_pct': round(cnt_a * 100.0 / tot_sess, 1) if tot_sess else 0.0,
                'sick_pct': round(cnt_s * 100.0 / tot_sess, 1) if tot_sess else 0.0,
                'personal_pct': round(cnt_per * 100.0 / tot_sess, 1) if tot_sess else 0.0,
            }
            att_by_ce = {}
            for a in att_list:
                ce_d = to_ce_date(a.date)
                entry = att_by_ce.setdefault(ce_d, {'morning': None, 'afternoon': None})
                if a.session_period == Attendance.PERIOD_AFTERNOON:
                    entry['afternoon'] = a
                else:
                    entry['morning'] = a

            all_ce_dates = sorted(set(td_by_ce.keys()) | set(att_by_ce.keys()))
            att_days = []
            for ce_d in all_ce_dates:
                td = td_by_ce.get(ce_d)
                ev_type = td.event_type if td else TrainingDate.TYPE_TRAINING
                act_p = td.active_periods() if td else [Attendance.PERIOD_MORNING, Attendance.PERIOD_AFTERNOON]
                a_pair = att_by_ce.get(ce_d, {'morning': None, 'afternoon': None})
                att_days.append({
                    'date': td.date if td else ce_d,
                    'event_type': ev_type,
                    'title': td.title if td else '',
                    'exam_period': td.exam_period if td else TrainingDate.EXAM_BOTH,
                    'has_morning': Attendance.PERIOD_MORNING in act_p,
                    'has_afternoon': Attendance.PERIOD_AFTERNOON in act_p,
                    'morning': a_pair['morning'],
                    'afternoon': a_pair['afternoon'],
                    'is_future': is_future_attendance_date(ce_d),
                })
            item['att_days'] = att_days
            missed_exams = 0
            for day in att_days:
                if day['event_type'] == TrainingDate.TYPE_EXAM:
                    if day['has_morning'] and day['morning'] is not None and day['morning'].status != Attendance.STATUS_PRESENT:
                        missed_exams += 1
                    if day['has_afternoon'] and day['afternoon'] is not None and day['afternoon'].status != Attendance.STATUS_PRESENT:
                        missed_exams += 1
            item['att_summary']['missed_exams'] = missed_exams
            if not approved:
                item['pending_reasons'].append('รอผู้ดูแลระบบอนุมัติประกาศผลการคัดเลือก')
        else:
            round_closed = True
            if rec is not None:
                if rec.rights_round_id and rec.rights_round.is_open:
                    round_closed = False
                if RightsRound.objects.filter(academic_year=year, camp_level=rec.camp_level, is_open=True).exists():
                    round_closed = False
            if rec is None:
                item['pending_reasons'].append('ยังไม่ได้ใช้สิทธิ์ในปีการศึกษานี้')
            if not round_closed:
                item['pending_reasons'].append('รอปิดระบบการใช้สิทธิ์')
            if not approved:
                item['pending_reasons'].append('รอผู้ดูแลระบบอนุมัติประกาศผลคะแนน')
            item['show_scores'] = rec is not None and round_closed and approved

        if item['show_scores']:
            item['total'], item['breakdown'] = total_score(user, year, only_visible=True)
            for b in item['breakdown']:
                w_val = float(b['category'].weight or 0)
                b['pct'] = min(100.0, round(float(b['weighted']) * 100.0 / w_val, 1)) if (b['weighted'] is not None and w_val > 0) else 0.0
        if approved and item['show_scores']:
            item['result'] = QualificationResult.objects.filter(academic_year=year, user=user).first()
            item['show_result'] = True
        if item['show_result']:
            item['show_exam_status'] = True
            has_camp2_in_year = bool(
                user.camp2_code and (
                    (item['result'] is not None and item['result'].passed and (camp2_year is None or camp2_year <= year))
                    or (camp2_year is not None and camp2_year <= year)
                    or (rec is not None and rec.camp_level == 2)
                )
            )
            if has_camp2_in_year:
                item['exam_status_camp'] = 2
                item['exam_status_text'] = 'มีสิทธิ์สอบท้ายค่าย 2'
            else:
                item['exam_status_camp'] = 1
                item['exam_status_text'] = 'มีสิทธิ์สอบท้ายค่าย 1'
        try:
            from .certificate_pdf import check_certificate_eligibility
            item['certificate_eligibility'] = check_certificate_eligibility(user, year)
        except Exception:
            item['certificate_eligibility'] = None
        results.append(item)
    return results


# =============================================================================
# Food ordering (ย้ายจากระบบ food เดิม)
# =============================================================================
class FoodOrder(models.Model):
    SCOPE_ALL = 'all'
    SCOPE_CHOICES = [('all', 'ทุกศูนย์'), ('DS', 'เฉพาะศูนย์ DS'), ('KM', 'เฉพาะศูนย์ KM')]
    MODE_UNIQUE = 'unique'
    MODE_REPEAT = 'repeat'
    MODE_CHOICES = [(MODE_UNIQUE, 'เลือกเมนูไม่ซ้ำ'), (MODE_REPEAT, 'เลือกเมนูซ้ำได้')]

    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    food_date = models.DateField('วันที่จัดอาหาร')
    restaurant_name = models.CharField('ชื่อร้านอาหาร', max_length=255)
    scope = models.CharField('ขอบเขต', max_length=5, choices=SCOPE_CHOICES, default=SCOPE_ALL)
    allow_current_students = models.BooleanField('นักเรียนปัจจุบัน', default=True)
    allow_rights_students = models.BooleanField('นักเรียนใช้สิทธิ์', default=False)
    max_selection_count = models.PositiveIntegerField('จำนวนสูงสุดที่เลือกได้', default=1)
    selection_mode = models.CharField('รูปแบบการเลือก', max_length=10, choices=MODE_CHOICES, default=MODE_UNIQUE)
    edit_toggle_open = models.BooleanField('เปิดให้เลือก/แก้ไข', default=True)
    overall_image = models.ImageField('รูปภาพรวม', upload_to='food_orders/', null=True, blank=True)
    additional_detail = models.TextField('รายละเอียดเพิ่มเติม', blank=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                   related_name='+')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'รายการสั่งอาหาร'
        verbose_name_plural = 'รายการสั่งอาหาร'
        ordering = ['-food_date', '-created_at']

    def __str__(self):
        return f'{self.food_date} - {self.restaurant_name}'

    def get_status(self):
        today = timezone.localdate()
        if self.food_date < today:
            return 'completed'
        if not self.edit_toggle_open:
            return 'closed'
        return 'active'

    def get_status_display_th(self):
        return {'active': 'กำลังเปิดเลือก', 'closed': 'ปิดการแก้ไข', 'completed': 'เสร็จสิ้น'}[self.get_status()]

    def is_editable_by_member(self):
        return self.edit_toggle_open and self.food_date >= timezone.localdate()

    def target_groups_label(self):
        parts = []
        if self.allow_current_students:
            parts.append('นักเรียนปัจจุบัน')
        if self.allow_rights_students:
            parts.append('นักเรียนใช้สิทธิ์')
        return ', '.join(parts) if parts else 'ไม่ได้เลือกกลุ่ม'

    def is_visible_to(self, user, year=None, is_rights_user=None):
        if self.scope != self.SCOPE_ALL and self.scope != (user.camp_center or ''):
            return False
        if user.role == 'member':
            cur_y = int(year if year is not None else current_year())
            if self.academic_year != cur_y:
                return False
            yy = f'{cur_y % 100:02d}'
            is_current = user.camp1_code.upper().startswith(yy) or (user.academic_year == cur_y)
            if is_rights_user is None:
                from composn.models import UseRecord
                is_rights_user = UseRecord.objects.filter(
                    user=user, academic_year=cur_y, is_used=True
                ).exists()
            allowed = (self.allow_current_students and is_current) or (self.allow_rights_students and is_rights_user)
            if not allowed:
                return False
        return True


class FoodMenu(models.Model):
    food_order = models.ForeignKey(FoodOrder, on_delete=models.CASCADE, related_name='menus',
                                   verbose_name='รายการสั่งอาหาร')
    menu_name = models.CharField('ชื่อเมนู', max_length=255)
    menu_description = models.TextField('รายละเอียดเมนู', blank=True)
    menu_image = models.ImageField('รูปภาพเมนู', upload_to='food_menus/', null=True, blank=True)
    sort_order = models.IntegerField('ลำดับ', default=0)

    class Meta:
        verbose_name = 'เมนูอาหาร'
        verbose_name_plural = 'เมนูอาหาร'
        ordering = ['sort_order', 'id']

    def __str__(self):
        return self.menu_name

    def total_selected(self):
        return self.selections.aggregate(t=Sum('quantity'))['t'] or 0


class FoodSelection(models.Model):
    food_order = models.ForeignKey(FoodOrder, on_delete=models.CASCADE, related_name='selections')
    food_menu = models.ForeignKey(FoodMenu, on_delete=models.CASCADE, related_name='selections')
    user = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, on_delete=models.SET_NULL,
                             related_name='food_selections')
    quantity = models.PositiveIntegerField('จำนวน', default=1)
    camp_code_snapshot = models.CharField(max_length=10, blank=True)
    full_name_snapshot = models.CharField(max_length=255, blank=True)
    center_snapshot = models.CharField(max_length=2, blank=True)
    selected_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'การเลือกอาหาร'
        verbose_name_plural = 'การเลือกอาหาร'
        unique_together = [['food_order', 'food_menu', 'user']]

    def save(self, *args, **kwargs):
        if self.user and not self.camp_code_snapshot:
            self.camp_code_snapshot = self.user.camp1_code
            self.full_name_snapshot = self.user.get_full_name()
            self.center_snapshot = self.user.camp_center
        super().save(*args, **kwargs)


# =============================================================================
# แบบสอบถามสำหรับนักเรียนในค่าย (Camp Survey — สำหรับสำรวจการสั่งอาหารและสรุปภาพรวม)
# =============================================================================
class CampSurvey(models.Model):
    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    title = models.CharField('หัวข้อแบบสอบถาม', max_length=200)
    description = models.TextField('คำอธิบาย', blank=True)
    allow_current_students = models.BooleanField('นักเรียนปัจจุบัน', default=True)
    allow_rights_students = models.BooleanField('นักเรียนใช้สิทธิ์', default=False)
    is_open = models.BooleanField('เปิดให้นักเรียนตอบแบบสอบถาม', default=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
                                   related_name='+', verbose_name='สร้างโดย')
    created_at = models.DateTimeField('สร้างเมื่อ', auto_now_add=True)
    updated_at = models.DateTimeField('แก้ไขเมื่อ', auto_now=True)

    class Meta:
        verbose_name = 'แบบสอบถามในค่าย'
        verbose_name_plural = 'แบบสอบถามในค่าย'
        ordering = ['-academic_year', '-created_at']

    def __str__(self):
        return f'{self.title} ({self.academic_year})'

    def target_groups_label(self):
        parts = []
        if self.allow_current_students:
            parts.append('นักเรียนปัจจุบัน')
        if self.allow_rights_students:
            parts.append('นักเรียนใช้สิทธิ์')
        return ', '.join(parts) if parts else 'ไม่ได้เลือกกลุ่ม'

    def is_visible_to(self, user, year=None, is_rights_user=None):
        if user.role == 'member':
            cur_y = int(year if year is not None else current_year())
            if self.academic_year != cur_y:
                return False
            yy = f'{cur_y % 100:02d}'
            is_current = user.camp1_code.upper().startswith(yy) or (user.academic_year == cur_y)
            if is_rights_user is None:
                from composn.models import UseRecord
                is_rights_user = UseRecord.objects.filter(
                    user=user, academic_year=cur_y, is_used=True
                ).exists()
            allowed = (self.allow_current_students and is_current) or (self.allow_rights_students and is_rights_user)
            if not allowed:
                return False
        return True


class CampSurveyQuestion(models.Model):
    TYPE_CHOICE = 'choice'
    TYPE_MULTIPLE = 'multiple'
    TYPE_RATING = 'rating'
    TYPE_TEXT = 'text'
    TYPE_CHOICES = [
        (TYPE_CHOICE, 'ตัวเลือกเดียว (Radio)'),
        (TYPE_MULTIPLE, 'เลือกได้หลายข้อ (Checkbox)'),
        (TYPE_RATING, 'คะแนนความพึงพอใจ (1-5)'),
        (TYPE_TEXT, 'ข้อความ / ข้อเสนอแนะ'),
    ]
    survey = models.ForeignKey(CampSurvey, on_delete=models.CASCADE, related_name='questions',
                               verbose_name='แบบสอบถาม')
    question_text = models.CharField('คำถาม', max_length=255)
    question_type = models.CharField('รูปแบบคำถาม', max_length=20, choices=TYPE_CHOICES, default=TYPE_CHOICE)
    choices_text = models.TextField('ตัวเลือก (คั่นด้วยเครื่องหมายจุลภาค ,)', blank=True,
                                    help_text='เช่น ข้าวผัด, กะเพราไก่, ก๋วยเตี๋ยว, อาหารฮาลาล')
    is_required = models.BooleanField('จำเป็นต้องตอบ', default=True)
    sort_order = models.PositiveIntegerField('ลำดับ', default=1)

    class Meta:
        verbose_name = 'คำถามแบบสอบถาม'
        verbose_name_plural = 'คำถามแบบสอบถาม'
        ordering = ['sort_order', 'id']

    def __str__(self):
        return f'{self.survey.title} - {self.question_text}'

    def get_choices_list(self):
        if self.question_type == self.TYPE_RATING:
            return ['5', '4', '3', '2', '1']
        return [c.strip() for c in (self.choices_text or '').split(',') if c.strip()]


class CampSurveyAnswer(models.Model):
    survey = models.ForeignKey(CampSurvey, on_delete=models.CASCADE, related_name='answers',
                               verbose_name='แบบสอบถาม')
    question = models.ForeignKey(CampSurveyQuestion, on_delete=models.CASCADE, related_name='answers',
                                 verbose_name='คำถาม')
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='survey_answers',
                             verbose_name='นักเรียน')
    answer_text = models.TextField('คำตอบ', blank=True)
    submitted_at = models.DateTimeField('ตอบเมื่อ', auto_now=True)

    class Meta:
        verbose_name = 'คำตอบแบบสอบถาม'
        verbose_name_plural = 'คำตอบแบบสอบถาม'
        constraints = [
            models.UniqueConstraint(fields=['question', 'user'], name='uniq_survey_question_user'),
        ]


def build_survey_summary(survey):
    """สรุปผลคำตอบของแบบสอบถามเพื่อแสดงบน Dashboard และหน้าสรุปภาพรวมการสั่งอาหาร"""
    questions = list(survey.questions.all().order_by('sort_order', 'id'))
    all_answers = list(survey.answers.select_related('user').all())
    respondents = {a.user_id for a in all_answers if (a.answer_text or '').strip()}
    total_respondents = len(respondents)

    q_summaries = []
    for q in questions:
        q_ans = [a for a in all_answers if a.question_id == q.pk and (a.answer_text or '').strip()]
        q_count = len(q_ans)
        item = {
            'question': q,
            'answer_count': q_count,
            'option_stats': [],
            'avg_rating': None,
            'text_answers': [],
        }
        if q.question_type in (CampSurveyQuestion.TYPE_CHOICE, CampSurveyQuestion.TYPE_MULTIPLE,
                               CampSurveyQuestion.TYPE_RATING):
            opts = q.get_choices_list()
            counts = {opt: 0 for opt in opts}
            for a in q_ans:
                if q.question_type == CampSurveyQuestion.TYPE_MULTIPLE:
                    parts = [p.strip() for p in a.answer_text.split(',') if p.strip()]
                else:
                    parts = [a.answer_text.strip()]
                for p in parts:
                    counts[p] = counts.get(p, 0) + 1
            for opt, cnt in counts.items():
                pct = round((cnt * 100.0 / q_count), 1) if q_count else 0.0
                item['option_stats'].append({'option': opt, 'count': cnt, 'percent': pct})
            if q.question_type == CampSurveyQuestion.TYPE_RATING and q_count:
                nums = []
                for a in q_ans:
                    try:
                        nums.append(float(a.answer_text.strip()))
                    except ValueError:
                        pass
                if nums:
                    item['avg_rating'] = round(sum(nums) / len(nums), 2)
        else:
            item['text_answers'] = [{'user': a.user, 'text': a.answer_text.strip(), 'at': a.submitted_at}
                                    for a in q_ans]
        q_summaries.append(item)

    return {
        'survey': survey,
        'respondent_count': total_respondents,
        'questions': q_summaries,
    }


# =============================================================================
# Certificate Settings & Lecturer/Committee Recipients
# =============================================================================
DEFAULT_CERT_HEADER1_TEXT = (
    'มูลนิธิส่งเสริมโอลิมปิกวิชาการและพัฒนามาตรฐานวิทยาศาสตร์ศึกษาฯ สอวน.\n'
    'ร่วมกับ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย 1 โรงเรียนเทพศิรินทร์'
)
DEFAULT_CERT_HEADER2_TEXT = 'ขอมอบเกียรติบัตรฉบับนี้ให้ไว้เพื่อแสดงว่า'
DEFAULT_CERT_STUDENT_LINE3 = (
    'ได้ผ่านโครงการโอลิมปิกวิชาการ สอวน. ค่าย ๑ วิชา คอมพิวเตอร์ ปีการศึกษา {year_thai}\n'
    'ณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์'
)
DEFAULT_CERT_LECTURER_LINE3 = (
    'ได้เข้าร่วมเป็นวิทยากรอบรมนักเรียน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\n'
    'ณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์ จังหวัดกรุงเทพมหานคร'
)
DEFAULT_CERT_COMMITTEE_LINE3 = (
    'ได้เข้าร่วมเป็นคณะกรรมการดำเนินงาน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\n'
    'ระหว่างวันที่ ๒๘ กันยายน - ๑๔ ตุลาคม {year_thai}\n'
    'ณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์ จังหวัดกรุงเทพมหานคร'
)


class CertificateSetting(models.Model):
    """การตั้งค่าเกียรติบัตรออนไลน์ (Singleton): รูปพื้นหลัง + ข้อความส่วนหัว/ตัว + QR Code ตรวจสอบ (TH Sarabun)"""
    background_image = models.ImageField(
        'รูปพื้นหลังเกียรติบัตร', upload_to='certificates/backgrounds/', null=True, blank=True
    )

    # ข้อความส่วนหัวที่ 1: มูลนิธิ / ศูนย์อบรม (รองรับหลายบรรทัด)
    header1_text = models.TextField(
        'ข้อความส่วนหัวที่ 1 (มูลนิธิและศูนย์อบรม)', blank=True, default=DEFAULT_CERT_HEADER1_TEXT
    )
    header1_y = models.FloatField('ตำแหน่งแกน Y ส่วนหัวที่ 1 (mm จากขอบบน)', default=45.0)
    header1_font_size = models.FloatField('ขนาดตัวอักษร ส่วนหัวที่ 1 (pt)', default=22.0)
    header1_line_spacing = models.FloatField('ระยะห่างระหว่างบรรทัด ส่วนหัวที่ 1 (mm)', default=10.0)
    header1_bold = models.BooleanField('ตัวหนา ส่วนหัวที่ 1', default=True)
    header1_color = models.CharField('สีตัวอักษร ส่วนหัวที่ 1', max_length=20, default='#0f4c5c')

    # ข้อความส่วนหัวที่ 2: ขอมอบเกียรติบัตรฉบับนี้ให้ไว้เพื่อแสดงว่า
    header2_text = models.CharField(
        'ข้อความส่วนหัวที่ 2 (คำมอบเกียรติบัตร)', max_length=255, blank=True, default=DEFAULT_CERT_HEADER2_TEXT
    )
    header2_y = models.FloatField('ตำแหน่งแกน Y ส่วนหัวที่ 2 (mm จากขอบบน)', default=71.0)
    header2_font_size = models.FloatField('ขนาดตัวอักษร ส่วนหัวที่ 2 (pt)', default=24.0)
    header2_bold = models.BooleanField('ตัวหนา ส่วนหัวที่ 2', default=True)
    header2_color = models.CharField('สีตัวอักษร ส่วนหัวที่ 2', max_length=20, default='#b8860b')

    # ข้อความที่ 1: ชื่อ นามสกุล
    line1_y = models.FloatField('ตำแหน่งแกน Y ชื่อ-นามสกุล (mm จากขอบบน)', default=94.0)
    line1_font_size = models.FloatField('ขนาดตัวอักษร ชื่อ-นามสกุล (pt)', default=34.0)
    line1_bold = models.BooleanField('ตัวหนา ชื่อ-นามสกุล', default=True)
    line1_color = models.CharField('สีตัวอักษร ชื่อ-นามสกุล', max_length=20, default='#000000')

    # ข้อความที่ 2: โรงเรียน จังหวัด
    line2_y = models.FloatField('ตำแหน่งแกน Y โรงเรียน จังหวัด (mm จากขอบบน)', default=111.0)
    line2_font_size = models.FloatField('ขนาดตัวอักษร โรงเรียน จังหวัด (pt)', default=22.0)
    line2_bold = models.BooleanField('ตัวหนา โรงเรียน จังหวัด', default=True)
    line2_color = models.CharField('สีตัวอักษร โรงเรียน จังหวัด', max_length=20, default='#000000')
    student_default_province = models.CharField(
        'จังหวัดต่อท้ายโรงเรียนนักเรียน (กรณีในฐานข้อมูลยังไม่มี)', max_length=100, blank=True, default='กรุงเทพมหานคร'
    )

    # ข้อความที่ 3: รายละเอียดการเข้าร่วม / ผ่านการอบรม (รองรับหลายบรรทัด)
    line3_y = models.FloatField('ตำแหน่งแกน Y รายละเอียด (mm จากขอบบน)', default=127.0)
    line3_font_size = models.FloatField('ขนาดตัวอักษร รายละเอียด (pt)', default=21.0)
    line3_bold = models.BooleanField('ตัวหนา รายละเอียด', default=True)
    line3_color = models.CharField('สีตัวอักษร รายละเอียด', max_length=20, default='#000000')
    line3_line_spacing = models.FloatField('ระยะห่างระหว่างบรรทัดของข้อความที่ 3 (mm)', default=11.0)
    student_line3_text = models.TextField('ข้อความที่ 3 สำหรับนักเรียน', default=DEFAULT_CERT_STUDENT_LINE3)
    lecturer_line3_text = models.TextField('ข้อความที่ 3 เริ่มต้นสำหรับวิทยากร', default=DEFAULT_CERT_LECTURER_LINE3)
    committee_line3_text = models.TextField('ข้อความที่ 3 เริ่มต้นสำหรับคณะกรรมการ', default=DEFAULT_CERT_COMMITTEE_LINE3)

    # ข้อความที่ 4: ให้ไว้ ณ วันที่ (ใช้ร่วมกันทุกกลุ่ม ตั้งค่าที่เดียว)
    line4_text = models.CharField(
        'ข้อความที่ 4 (ให้ไว้ ณ วันที่ — ใช้ร่วมกันทั้งหมด)', max_length=255, default='ให้ไว้ ณ วันที่ ๑๔ ตุลาคม {year_thai}'
    )
    line4_y = models.FloatField('ตำแหน่งแกน Y ให้ไว้ ณ วันที่ (mm จากขอบบน)', default=162.0)
    line4_font_size = models.FloatField('ขนาดตัวอักษร ให้ไว้ ณ วันที่ (pt)', default=21.0)
    line4_bold = models.BooleanField('ตัวหนา ให้ไว้ ณ วันที่', default=True)
    line4_color = models.CharField('สีตัวอักษร ให้ไว้ ณ วันที่', max_length=20, default='#000000')

    # เลขที่อ้างอิงเกียรติบัตร & QR Code ตรวจสอบความถูกต้อง
    show_qr_verify = models.BooleanField('แสดง QR Code และเลขที่อ้างอิงตรวจสอบ', default=True)
    qr_x = models.FloatField('ตำแหน่งแกน X ของ QR Code (mm จากขอบซ้าย)', default=252.0)
    qr_y = models.FloatField('ตำแหน่งแกน Y ของ QR Code (mm จากขอบบน)', default=164.0)
    qr_size = models.FloatField('ขนาด QR Code (mm)', default=24.0)
    ref_code_font_size = models.FloatField('ขนาดตัวอักษรเลขที่อ้างอิง (pt)', default=9.5)
    ref_code_color = models.CharField('สีตัวอักษรเลขที่อ้างอิง', max_length=20, default='#333333')

    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'ตั้งค่าเกียรติบัตร'
        verbose_name_plural = 'ตั้งค่าเกียรติบัตร'

    def __str__(self):
        return 'การตั้งค่าเกียรติบัตร'

    @classmethod
    def get(cls):
        """คืนค่า CertificateSetting แบบ Singleton"""
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj

    def render_text(self, text, year):
        """แทนที่ตัวแปร {year} และ {year_thai} ในข้อความ"""
        from composn.templatetags.composn_tags import to_thai_num
        raw = str(text or '')
        y_str = str(year) if year is not None else str(current_year())
        return raw.replace('{year}', y_str).replace('{year_thai}', to_thai_num(y_str))

    def format_student_school(self, school_name):
        """จัดรูปแบบชื่อโรงเรียน + จังหวัดต่อท้ายสำหรับนักเรียน"""
        school = (school_name or '').strip()
        prov = (self.student_default_province or '').strip()
        if school and prov and prov not in school:
            return f'{school} {prov}'
        return school

    def default_line3_for(self, group):
        """คืนข้อความเริ่มต้นข้อ 3 ตามกลุ่ม ('student', 'lecturer', 'committee')"""
        if group == 'lecturer':
            return self.lecturer_line3_text
        if group == 'committee':
            return self.committee_line3_text
        return self.student_line3_text

    def reset_layout_defaults(self):
        """รีเซ็ตเฉพาะตำแหน่งแกน Y/X, ขนาดตัวอักษร, ตัวหนา, สี, และระยะบรรทัดกลับค่าเริ่มต้น (คงรูปพื้นหลังและข้อความไว้)"""
        self.header1_y = 45.0
        self.header1_font_size = 22.0
        self.header1_line_spacing = 10.0
        self.header1_bold = True
        self.header1_color = '#0f4c5c'

        self.header2_y = 71.0
        self.header2_font_size = 24.0
        self.header2_bold = True
        self.header2_color = '#b8860b'

        self.line1_y = 94.0
        self.line1_font_size = 34.0
        self.line1_bold = True
        self.line1_color = '#000000'

        self.line2_y = 111.0
        self.line2_font_size = 22.0
        self.line2_bold = True
        self.line2_color = '#000000'

        self.line3_y = 127.0
        self.line3_font_size = 21.0
        self.line3_bold = True
        self.line3_color = '#000000'
        self.line3_line_spacing = 11.0

        self.line4_y = 162.0
        self.line4_font_size = 21.0
        self.line4_bold = True
        self.line4_color = '#000000'

        self.qr_x = 252.0
        self.qr_y = 164.0
        self.qr_size = 24.0
        self.ref_code_font_size = 9.5
        self.ref_code_color = '#333333'


class CertificateYearStatus(models.Model):
    """สถานะการเปิด/ปิดแสดงเกียรติบัตรให้สมาชิกแต่ละปีการศึกษา (ค่าเริ่มต้น: ปิด)"""
    academic_year = models.IntegerField('ปีการศึกษา', unique=True, db_index=True)
    is_published = models.BooleanField('เปิดแสดงเกียรติบัตรให้นักเรียน', default=False)
    updated_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='อัปเดตโดย',
    )
    updated_at = models.DateTimeField('อัปเดตล่าสุด', auto_now=True)

    class Meta:
        verbose_name = 'สถานะแสดงเกียรติบัตรประจำปี'
        verbose_name_plural = 'สถานะแสดงเกียรติบัตรประจำปี'
        ordering = ['-academic_year']

    def __str__(self):
        st = 'เปิดแสดง' if self.is_published else 'ปิด'
        return f'เกียรติบัตร ปีการศึกษา {self.academic_year}: {st}'

    @classmethod
    def is_year_published(cls, year):
        if year is None:
            return False
        try:
            y = int(year)
        except (TypeError, ValueError):
            return False
        return cls.objects.filter(academic_year=y, is_published=True).exists()


class CertificateRecipient(models.Model):
    """รายชื่อผู้รับเกียรติบัตรกลุ่มวิทยากร และคณะกรรมการ"""
    TYPE_LECTURER = 'lecturer'
    TYPE_COMMITTEE = 'committee'
    TYPE_CHOICES = [
        (TYPE_LECTURER, 'วิทยากร'),
        (TYPE_COMMITTEE, 'คณะกรรมการ'),
    ]

    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    recipient_type = models.CharField('ประเภท', max_length=20, choices=TYPE_CHOICES, default=TYPE_LECTURER, db_index=True)
    full_name = models.CharField('ชื่อ - นามสกุล', max_length=200)
    school_province = models.CharField('โรงเรียน จังหวัด', max_length=255, blank=True)
    description = models.TextField('ข้อความรายละเอียด (ข้อ 3)', blank=True)
    sort_order = models.PositiveIntegerField('ลำดับ', default=0)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='เพิ่มโดย',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'รายชื่อผู้รับเกียรติบัตร (วิทยากร/คณะกรรมการ)'
        verbose_name_plural = 'รายชื่อผู้รับเกียรติบัตร (วิทยากร/คณะกรรมการ)'
        ordering = ['academic_year', 'recipient_type', 'sort_order', 'id']

    def __str__(self):
        return f'{self.get_recipient_type_display()}: {self.full_name} ({self.academic_year})'

    def resolved_description(self, cert_setting=None):
        """คืนข้อความข้อ 3 ที่ใช้งานจริง (หากเว้นว่างจะใช้ค่าเริ่มต้นของกลุ่มจาก CertificateSetting)"""
        desc = (self.description or '').strip()
        if desc:
            return desc
        cs = cert_setting or CertificateSetting.get()
        return cs.default_line3_for(self.recipient_type)


class CertificateRecord(models.Model):
    """ทะเบียนเลขที่อ้างอิงเกียรติบัตรสำหรับตรวจสอบความถูกต้องผ่าน QR Code / เว็บไซต์ศูนย์อบรม"""
    TYPE_STUDENT = 'student'
    TYPE_LECTURER = 'lecturer'
    TYPE_COMMITTEE = 'committee'
    TYPE_CHOICES = [
        (TYPE_STUDENT, 'นักเรียน'),
        (TYPE_LECTURER, 'วิทยากร'),
        (TYPE_COMMITTEE, 'คณะกรรมการ'),
    ]

    ref_code = models.CharField('เลขที่อ้างอิงเกียรติบัตร', max_length=64, unique=True, db_index=True)
    academic_year = models.IntegerField('ปีการศึกษา', db_index=True)
    recipient_type = models.CharField('ประเภทผู้รับ', max_length=20, choices=TYPE_CHOICES, db_index=True)
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='certificate_records', verbose_name='นักเรียน',
    )
    recipient = models.ForeignKey(
        CertificateRecipient, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='certificate_records', verbose_name='วิทยากร/คณะกรรมการ',
    )
    camp_code = models.CharField('รหัสประจำตัว', max_length=50, blank=True)
    full_name = models.CharField('ชื่อ - นามสกุล', max_length=200)
    school_province = models.CharField('โรงเรียน จังหวัด', max_length=255, blank=True)
    description = models.TextField('รายละเอียดเกียรติบัตร', blank=True)
    issued_date_text = models.CharField('ให้ไว้ ณ วันที่', max_length=255, blank=True)
    is_revoked = models.BooleanField('ยกเลิกแล้ว', default=False)
    issued_at = models.DateTimeField('ออกเมื่อ', auto_now_add=True)
    updated_at = models.DateTimeField('อัปเดตล่าสุด', auto_now=True)

    class Meta:
        verbose_name = 'ทะเบียนเกียรติบัตรตรวจสอบได้'
        verbose_name_plural = 'ทะเบียนเกียรติบัตรตรวจสอบได้'
        ordering = ['-academic_year', '-issued_at', 'ref_code']

    def __str__(self):
        return f'{self.ref_code} - {self.full_name} ({self.academic_year})'

    @classmethod
    def _generate_ref_code(cls, year, prefix_letter, ident):
        ident_clean = re.sub(r'[^A-Z0-9]', '', str(ident).upper())[:12] or '0000'
        while True:
            tok = secrets.token_hex(3).upper()
            code = f'DSPOSN-{year}-{prefix_letter}-{ident_clean}-{tok}'
            if not cls.objects.filter(ref_code=code).exists():
                return code

    @classmethod
    def issue_for_student(cls, user, year, cert_setting=None):
        """ออกหรืออัปเดตทะเบียนเกียรติบัตรนักเรียนโดยคง ref_code เดิมไว้เสมอ"""
        cs = cert_setting or CertificateSetting.get()
        full_name = (user.get_full_name() or '')[:200]
        school_province = (cs.format_student_school(user.school_name) or '')[:255]
        description = cs.render_text(cs.student_line3_text, year)
        issued_date_text = (cs.render_text(cs.line4_text, year) or '')[:255]
        camp_code = (getattr(user, 'camp1_code', '') or '')[:50]

        existing = cls.objects.filter(
            academic_year=year, recipient_type=cls.TYPE_STUDENT, user=user
        ).first()
        if existing:
            changed = []
            for fld, val in (
                ('camp_code', camp_code),
                ('full_name', full_name),
                ('school_province', school_province),
                ('description', description),
                ('issued_date_text', issued_date_text),
            ):
                if getattr(existing, fld) != val:
                    setattr(existing, fld, val)
                    changed.append(fld)
            if changed:
                existing.save(update_fields=changed + ['updated_at'])
            return existing

        ref_code = cls._generate_ref_code(year, 'S', camp_code or f'{user.pk:04d}')
        return cls.objects.create(
            ref_code=ref_code,
            academic_year=year,
            recipient_type=cls.TYPE_STUDENT,
            user=user,
            camp_code=camp_code,
            full_name=full_name,
            school_province=school_province,
            description=description,
            issued_date_text=issued_date_text,
        )

    @classmethod
    def issue_for_recipient(cls, recipient, cert_setting=None):
        """ออกหรืออัปเดตทะเบียนเกียรติบัตรวิทยากร/คณะกรรมการโดยคง ref_code เดิมไว้เสมอ"""
        cs = cert_setting or CertificateSetting.get()
        year = recipient.academic_year
        prefix = 'L' if recipient.recipient_type == CertificateRecipient.TYPE_LECTURER else 'C'
        full_name = (recipient.full_name or '')[:200]
        school_province = (recipient.school_province or '')[:255]
        description = cs.render_text(recipient.resolved_description(cs), year)
        issued_date_text = (cs.render_text(cs.line4_text, year) or '')[:255]

        existing = cls.objects.filter(academic_year=year, recipient=recipient).first()
        if existing:
            changed = []
            for fld, val in (
                ('recipient_type', recipient.recipient_type),
                ('full_name', full_name),
                ('school_province', school_province),
                ('description', description),
                ('issued_date_text', issued_date_text),
            ):
                if getattr(existing, fld) != val:
                    setattr(existing, fld, val)
                    changed.append(fld)
            if changed:
                existing.save(update_fields=changed + ['updated_at'])
            return existing

        ref_code = cls._generate_ref_code(year, prefix, f'{recipient.pk:04d}')
        return cls.objects.create(
            ref_code=ref_code,
            academic_year=year,
            recipient_type=recipient.recipient_type,
            recipient=recipient,
            full_name=full_name,
            school_province=school_province,
            description=description,
            issued_date_text=issued_date_text,
        )


# =============================================================================
# Exam Seating Management (ระบบจัดที่นั่งสอบ)
# =============================================================================
class ExamRoom(models.Model):
    """ห้องสอบ (ห้องปฏิบัติการคอมพิวเตอร์ หรือ ห้องประชุม/ห้องสอบทั่วไป)"""
    FRONT_TOP = 'top'
    FRONT_LEFT = 'left'
    FRONT_RIGHT = 'right'
    FRONT_BOTTOM = 'bottom'
    FRONT_CHOICES = [
        (FRONT_TOP, 'ด้านบน (หน้าเวที)'),
        (FRONT_LEFT, 'ด้านซ้าย'),
        (FRONT_RIGHT, 'ด้านขวา'),
        (FRONT_BOTTOM, 'ด้านล่าง'),
    ]

    FILL_SEQ_ASC = 'seq_asc'
    FILL_LR = 'lr'
    FILL_RL = 'rl'
    FILL_SNAKE_AUTO = 'snake_auto'
    FILL_SNAKE_LR = 'snake_lr'
    FILL_SNAKE_RL = 'snake_rl'
    FILL_PATTERN_CHOICES = [
        (FILL_SEQ_ASC, 'เรียงตามหมายเลขเครื่อง/โต๊ะ (น้อย → มาก)'),
        (FILL_SNAKE_AUTO, 'สลับฟันปลา (อัตโนมัติตามที่นั่งเริ่มต้น)'),
        (FILL_SNAKE_LR, 'สลับฟันปลา เริ่มซ้าย→ขวา (ซ้าย→ขวา, ขวา→ซ้าย, ...)'),
        (FILL_SNAKE_RL, 'สลับฟันปลา เริ่มขวา→ซ้าย (ขวา→ซ้าย, ซ้าย→ขวา, ...)'),
        (FILL_LR, 'ซ้าย → ขวา ทุกแถว'),
        (FILL_RL, 'ขวา → ซ้าย ทุกแถว'),
    ]

    name = models.CharField('ชื่อห้องสอบ', max_length=150, unique=True)
    has_computer = models.BooleanField('มีคอมพิวเตอร์ (ห้องปฏิบัติการ)', default=True)
    front_position = models.CharField('ตำแหน่งหน้าห้อง / หน้าเวที', max_length=10, choices=FRONT_CHOICES, default=FRONT_TOP)
    front_label = models.CharField('ข้อความป้ายหน้าห้อง', max_length=50, default='หน้าห้อง')
    default_fill_pattern = models.CharField('รูปแบบการจัดเริ่มต้น', max_length=20, choices=FILL_PATTERN_CHOICES, default=FILL_SEQ_ASC)
    default_alternate = models.BooleanField('ค่าเริ่มต้นแบบเครื่องเว้นเครื่อง', default=True)
    is_active = models.BooleanField('เปิดใช้งานห้อง', default=True)
    sort_order = models.PositiveIntegerField('ลำดับ', default=0)
    description = models.CharField('หมายเหตุ/รายละเอียดห้อง', max_length=255, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'ห้องสอบ'
        verbose_name_plural = 'ห้องสอบ'
        ordering = ['sort_order', 'id']

    def __str__(self):
        return self.name

    @property
    def room_type_label(self):
        return 'ห้องปฏิบัติการ (มีคอมพิวเตอร์)' if self.has_computer else 'ห้องสอบโต๊ะ (ไม่มีคอมพิวเตอร์)'


class ExamSeat(models.Model):
    """ที่นั่ง / เครื่องคอมพิวเตอร์แต่ละจุดภายในห้องสอบ"""
    LABEL_BELOW = 'below'
    LABEL_ABOVE = 'above'
    LABEL_POS_CHOICES = [
        (LABEL_BELOW, 'รหัสผู้สอบอยู่ด้านล่างเลขที่นั่ง'),
        (LABEL_ABOVE, 'รหัสผู้สอบอยู่ด้านบนเลขที่นั่ง'),
    ]
    STATUS_NORMAL = 'normal'
    STATUS_RESERVE = 'reserve'
    STATUS_DISABLED = 'disabled'
    DEFAULT_STATUS_CHOICES = [
        (STATUS_NORMAL, 'ที่นั่งสอบปกติ'),
        (STATUS_RESERVE, 'เครื่องสำรอง (ค่าเริ่มต้น)'),
        (STATUS_DISABLED, 'งดใช้งาน/ชำรุด'),
    ]

    room = models.ForeignKey(ExamRoom, on_delete=models.CASCADE, related_name='seats', verbose_name='ห้องสอบ')
    seat_label = models.CharField('เลขที่นั่ง / เลขเครื่อง', max_length=20)
    seat_number = models.IntegerField('เลขลำดับ', default=0, db_index=True)
    block_index = models.PositiveSmallIntegerField('กลุ่มแถว (Block)', default=1)
    row_index = models.PositiveSmallIntegerField('แถวที่ (Row)', default=1)
    col_index = models.PositiveSmallIntegerField('คอลัมน์ที่ (Col)', default=1)
    label_position = models.CharField('ตำแหน่งช่องผู้สอบเทียบกับเลขที่นั่ง', max_length=10, choices=LABEL_POS_CHOICES, default=LABEL_BELOW)
    default_status = models.CharField('สถานะเริ่มต้น', max_length=15, choices=DEFAULT_STATUS_CHOICES, default=STATUS_NORMAL)
    is_active = models.BooleanField('ใช้งาน', default=True)

    class Meta:
        verbose_name = 'ที่นั่งสอบ'
        verbose_name_plural = 'ที่นั่งสอบ'
        ordering = ['room', 'block_index', 'row_index', 'col_index']
        constraints = [
            models.UniqueConstraint(fields=['room', 'seat_label'], name='uniq_exam_seat_room_label'),
            models.UniqueConstraint(fields=['room', 'row_index', 'col_index'], name='uniq_exam_seat_room_pos'),
        ]

    def __str__(self):
        return f'{self.room.name} - {self.seat_label}'


class ExamSeatingPlan(models.Model):
    """ผังการจัดที่นั่งสอบประจำวันสอบและช่วงเวลาสอบ"""
    TYPE_THEORY = 'theory'
    TYPE_PRACTICE = 'practice'
    TYPE_GENERAL = 'general'
    EXAM_TYPE_CHOICES = [
        (TYPE_THEORY, 'สอบทฤษฎี'),
        (TYPE_PRACTICE, 'สอบปฏิบัติ'),
        (TYPE_GENERAL, 'สอบทั่วไป'),
    ]

    PERIOD_BOTH = 'both'
    PERIOD_MORNING = 'morning'
    PERIOD_AFTERNOON = 'afternoon'
    PERIOD_CHOICES = [
        (PERIOD_BOTH, 'ทั้งวัน'),
        (PERIOD_MORNING, 'ภาคเช้า'),
        (PERIOD_AFTERNOON, 'ภาคบ่าย'),
    ]

    academic_year = models.IntegerField('ปีการศึกษา', default=current_year, db_index=True)
    training_date = models.ForeignKey(
        TrainingDate, on_delete=models.CASCADE, related_name='seating_plans', verbose_name='วันสอบ',
    )
    title = models.CharField('ชื่อแผนการจัดสอบ', max_length=200, blank=True)
    exam_type = models.CharField('ประเภทการสอบ', max_length=20, choices=EXAM_TYPE_CHOICES, default=TYPE_THEORY)
    session_period = models.CharField('ช่วงเวลาสอบ', max_length=15, choices=PERIOD_CHOICES, default=PERIOD_BOTH)
    is_open = models.BooleanField('เปิดประกาศที่นั่งสอบให้นักเรียนดู', default=False, db_index=True)
    note = models.TextField('หมายเหตุ', blank=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='สร้างโดย',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'แผนการจัดที่นั่งสอบ'
        verbose_name_plural = 'แผนการจัดที่นั่งสอบ'
        ordering = ['-academic_year', 'training_date__date', 'session_period', 'id']
        constraints = [
            models.UniqueConstraint(fields=['training_date', 'session_period'], name='uniq_exam_seating_plan_date_period'),
        ]

    def __str__(self):
        return f'{self.training_date.date} ({self.effective_period_display}) - {self.effective_exam_type_display}'

    @property
    def effective_exam_type_display(self):
        if self.training_date_id:
            return self.training_date.effective_exam_type_display
        return self.get_exam_type_display()

    @property
    def effective_period_display(self):
        if self.training_date_id:
            return self.training_date.get_exam_period_display()
        return self.get_session_period_display()

    @property
    def default_chart_subtitle(self):
        if self.training_date_id:
            return self.training_date.seating_chart_subtitle
        if self.exam_type == self.TYPE_THEORY:
            return 'แผนผังที่นั่งสอบทฤษฎี'
        if self.exam_type == self.TYPE_PRACTICE:
            return 'แผนผังที่นั่งสอบปฏิบัติ'
        return 'แผนผังที่นั่งสอบ'


class ExamSeatingPlanRoom(models.Model):
    """การเปิดใช้ห้องสอบและรูปแบบการจัดเรียงภายในวันสอบนั้น"""
    plan = models.ForeignKey(ExamSeatingPlan, on_delete=models.CASCADE, related_name='plan_rooms', verbose_name='แผนสอบ')
    room = models.ForeignKey(ExamRoom, on_delete=models.CASCADE, related_name='plan_rooms', verbose_name='ห้องสอบ')
    is_open = models.BooleanField('เปิดใช้ห้องสอบนี้', default=True)
    exam_subtitle = models.CharField('หัวข้อแผนผังห้องนี้ (เว้นว่างใช้ค่าจากแผน)', max_length=150, blank=True)
    fill_pattern = models.CharField('รูปแบบการจัดที่นั่ง', max_length=20, choices=ExamRoom.FILL_PATTERN_CHOICES, blank=True, default='')
    start_seat = models.ForeignKey(
        ExamSeat, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='ที่นั่งเริ่มต้น',
    )
    alternate_seats = models.BooleanField('จัดแบบเครื่องเว้นเครื่อง', default=True)
    sort_order = models.PositiveIntegerField('ลำดับแสดงผล', default=0)
    note = models.CharField('หมายเหตุห้องสอบ', max_length=255, blank=True)

    class Meta:
        verbose_name = 'ห้องสอบในแผนจัดสอบ'
        verbose_name_plural = 'ห้องสอบในแผนจัดสอบ'
        ordering = ['sort_order', 'room__sort_order', 'id']
        constraints = [
            models.UniqueConstraint(fields=['plan', 'room'], name='uniq_exam_plan_room'),
        ]

    def __str__(self):
        return f'{self.plan} - {self.room.name}'

    @property
    def effective_subtitle(self):
        return (self.exam_subtitle or '').strip() or self.plan.default_chart_subtitle

    @property
    def effective_fill_pattern(self):
        return (self.fill_pattern or '').strip() or self.room.default_fill_pattern


class ExamRoomProctor(models.Model):
    """วิทยากร / คณะกรรมการคุมสอบประจำห้องในวันสอบนั้น"""
    plan_room = models.ForeignKey(
        ExamSeatingPlanRoom, on_delete=models.CASCADE, related_name='proctors', verbose_name='ห้องสอบในแผน',
    )
    recipient = models.ForeignKey(
        CertificateRecipient, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='proctoring_assignments', verbose_name='วิทยากรในระบบ',
    )
    display_name = models.CharField('ชื่อผู้คุมสอบที่แสดง', max_length=200)
    sort_order = models.PositiveIntegerField('ลำดับ', default=0)

    class Meta:
        verbose_name = 'วิทยากรคุมสอบ'
        verbose_name_plural = 'วิทยากรคุมสอบ'
        ordering = ['sort_order', 'id']

    def __str__(self):
        return f'{self.plan_room.room.name}: {self.display_name}'


class ExamSeatAssignment(models.Model):
    """การจัดวางผู้เข้าสอบ / สถานะสำรอง / งดใช้งาน บนที่นั่งสอบแต่ละจุด"""
    STATUS_ASSIGNED = 'assigned'
    STATUS_RESERVE = 'reserve'
    STATUS_BLOCKED = 'blocked'
    STATUS_CHOICES = [
        (STATUS_ASSIGNED, 'จัดผู้เข้าสอบแล้ว'),
        (STATUS_RESERVE, 'สำรอง'),
        (STATUS_BLOCKED, 'งดใช้งาน'),
    ]

    plan = models.ForeignKey(ExamSeatingPlan, on_delete=models.CASCADE, related_name='assignments', verbose_name='แผนสอบ')
    plan_room = models.ForeignKey(
        ExamSeatingPlanRoom, on_delete=models.CASCADE, related_name='assignments', verbose_name='ห้องสอบในแผน',
    )
    seat = models.ForeignKey(ExamSeat, on_delete=models.CASCADE, related_name='assignments', verbose_name='ที่นั่งสอบ')
    status = models.CharField('สถานะ', max_length=15, choices=STATUS_CHOICES, default=STATUS_ASSIGNED)
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.CASCADE,
        related_name='exam_seat_assignments', verbose_name='ผู้เข้าสอบ',
    )
    assigned_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name='+', verbose_name='จัดโดย',
    )
    assigned_at = models.DateTimeField('อัปเดตเมื่อ', auto_now=True)

    class Meta:
        verbose_name = 'การจัดที่นั่งสอบ'
        verbose_name_plural = 'การจัดที่นั่งสอบ'
        ordering = ['plan_room', 'seat__block_index', 'seat__row_index', 'seat__col_index']
        constraints = [
            models.UniqueConstraint(fields=['plan_room', 'seat'], name='uniq_exam_plan_room_seat'),
            models.UniqueConstraint(
                fields=['plan', 'user'], condition=models.Q(user__isnull=False), name='uniq_exam_plan_user',
            ),
        ]

    def __str__(self):
        if self.status == self.STATUS_ASSIGNED and self.user:
            return f'{self.plan_room.room.name} #{self.seat.seat_label} -> {self.user.camp1_code}'
        return f'{self.plan_room.room.name} #{self.seat.seat_label} ({self.get_status_display()})'
