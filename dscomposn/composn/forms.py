from django import forms
from django.contrib.auth import password_validation
from django.contrib.auth.forms import AuthenticationForm

from .models import CAMP_LEVEL_CHOICES, DEFAULT_ACADEMIC_YEARS, PROFILE_EDITABLE_CHOICES, SystemSetting, User


def _get_dynamic_academic_years(setting_obj=None, extra_year=None):
    """ดึงรายการปีการศึกษาที่ใช้ได้ในระบบ (เรียงจากปีล่าสุดลงไป)"""
    try:
        from .context_processors import get_academic_year_choices
        target_setting = setting_obj if setting_obj is not None else SystemSetting.get()
        cur_year = getattr(target_setting, 'current_academic_year', 2569)
        years = list(get_academic_year_choices(cur_year=cur_year, setting=target_setting))
    except Exception:
        years = sorted(DEFAULT_ACADEMIC_YEARS, reverse=True)
    if extra_year is not None:
        try:
            ey = int(extra_year)
            if ey not in years:
                years = sorted(set(years) | {ey}, reverse=True)
        except (TypeError, ValueError):
            pass
    return years


class BootstrapMixin:
    def _bootstrap(self):
        for name, field in self.fields.items():
            w = field.widget
            if isinstance(w, (forms.CheckboxInput,)):
                w.attrs.setdefault('class', 'form-check-input')
            elif isinstance(w, forms.CheckboxSelectMultiple):
                pass
            elif isinstance(w, forms.Select):
                w.attrs.setdefault('class', 'form-select')
            else:
                w.attrs.setdefault('class', 'form-control')


class LoginForm(BootstrapMixin, AuthenticationForm):
    username = forms.CharField(label='ชื่อผู้ใช้', max_length=10,
                               widget=forms.TextInput(attrs={'autofocus': True, 'placeholder': 'เช่น 69DS01 หรือ 68KM01'}))
    password = forms.CharField(label='รหัสผ่าน', strip=False, widget=forms.PasswordInput)

    error_messages = {
        'invalid_login': 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง',
        'inactive': 'บัญชีนี้ถูกปิดการใช้งาน กรุณาติดต่อผู้ดูแลระบบ',
    }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()


class ChangePasswordForm(BootstrapMixin, forms.Form):
    old_password = forms.CharField(label='รหัสผ่านปัจจุบัน', widget=forms.PasswordInput, strip=False)
    new_password1 = forms.CharField(label='รหัสผ่านใหม่', widget=forms.PasswordInput, strip=False)
    new_password2 = forms.CharField(label='ยืนยันรหัสผ่านใหม่', widget=forms.PasswordInput, strip=False)

    def __init__(self, user, *args, **kwargs):
        self.user = user
        super().__init__(*args, **kwargs)
        self._bootstrap()

    def clean_old_password(self):
        pw = self.cleaned_data['old_password']
        if not self.user.check_password(pw):
            raise forms.ValidationError('รหัสผ่านปัจจุบันไม่ถูกต้อง')
        return pw

    def clean(self):
        cd = super().clean()
        p1, p2 = cd.get('new_password1'), cd.get('new_password2')
        if p1 and p2:
            if p1 != p2:
                self.add_error('new_password2', 'รหัสผ่านใหม่ทั้งสองช่องไม่ตรงกัน')
            elif len(p1) < 8:
                self.add_error('new_password1', 'รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 8 ตัวอักษร')
            elif p1 == cd.get('old_password'):
                self.add_error('new_password1', 'รหัสผ่านใหม่ต้องไม่ซ้ำกับรหัสผ่านเดิม')
            elif (
                p1.upper() == self.user.camp1_code.upper()
                or (self.user.camp2_code and p1.upper() == self.user.camp2_code.upper())
                or p1 == '1234'
            ):
                self.add_error('new_password1', 'ไม่อนุญาตให้ใช้รหัสผ่านเริ่มต้นหรือรหัสค่าย')
            else:
                try:
                    password_validation.validate_password(p1, self.user)
                except forms.ValidationError as e:
                    self.add_error('new_password1', e)
        return cd


class AdminSetPasswordForm(BootstrapMixin, forms.Form):
    new_password = forms.CharField(label='รหัสผ่านใหม่', widget=forms.PasswordInput, strip=False, min_length=4)
    force_change = forms.BooleanField(label='บังคับเปลี่ยนรหัสผ่านเมื่อ Login ครั้งถัดไป', required=False,
                                      initial=True)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()


class AdminUserCreateForm(BootstrapMixin, forms.Form):
    """ฟอร์มเพิ่มผู้ดูแลระบบ (Admin) หรือ เจ้าหน้าที่ (Staff) พร้อมกำหนดสิทธิ์รายบุคคลของ Staff"""
    camp1_code = forms.CharField(label='รหัสผู้ใช้ (Username)', max_length=50, help_text='ใช้สำหรับ Login เข้าระบบ')
    role = forms.ChoiceField(
        label='บทบาท',
        choices=[(User.ROLE_STAFF, 'เจ้าหน้าที่ (Staff)'), (User.ROLE_ADMIN, 'ผู้ดูแลระบบ (Admin)')],
        initial=User.ROLE_STAFF,
    )
    password = forms.CharField(label='รหัสผ่าน', widget=forms.PasswordInput, strip=False, min_length=4)
    staff_can_view_users = forms.BooleanField(
        label='ผู้ใช้ (นักเรียน/สมาชิก) — ดูอย่างเดียว', required=False, initial=False,
    )
    staff_can_attendance = forms.BooleanField(label='บันทึกเวลาเรียน', required=False, initial=True)
    staff_can_signin_sheet = forms.BooleanField(label='แบบลงชื่อ', required=False, initial=True)
    staff_can_manage_food = forms.BooleanField(
        label='รายการอาหาร / แบบสอบถาม', required=False, initial=True,
    )
    staff_can_scores = forms.BooleanField(label='บันทึกคะแนน / เอกสารประกาศผล', required=False, initial=True)
    staff_can_exam_seating = forms.BooleanField(label='จัดที่นั่งสอบ', required=False, initial=True)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()

    def clean_camp1_code(self):
        code = self.cleaned_data['camp1_code'].strip()
        if User.objects.filter(camp1_code__iexact=code).exists():
            raise forms.ValidationError('รหัสผู้ใช้นี้มีอยู่ในระบบแล้ว')
        return code


STAFF_PERM_FIELDS = [
    'staff_can_view_users', 'staff_can_attendance', 'staff_can_signin_sheet',
    'staff_can_manage_food', 'staff_can_scores', 'staff_can_exam_seating',
]

USER_FIELDS = ['camp1_code', 'camp2_code', 'first_name', 'last_name', 'nickname', 'school_name', 'birth_date',
               'camp_center', 'academic_year', 'current_level', 'year_plus_1', 'year_plus_2', 'year_plus_3',
               'year_plus_4', 'year_plus_5', 'student_phone', 'parent_phone', 'email', 'chronic_disease',
               'food_allergy', 'is_muslim', 'role', 'is_active', 'note'] + STAFF_PERM_FIELDS


def split_full_name(full_name):
    """แยกชื่อ-นามสกุลออกจากช่องเดียวกัน (หรือเก็บใน first_name หากมีคำเดียว)"""
    s = (full_name or '').strip()
    if not s:
        return '', ''
    parts = s.split(None, 1)
    return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], '')


class UserForm(BootstrapMixin, forms.ModelForm):
    full_name = forms.CharField(label='ชื่อ-นามสกุล', required=False, max_length=300)
    password = forms.CharField(label='รหัสผ่าน (เว้นว่าง = ใช้รหัสค่าย 1 / 1234 สำหรับ admin/staff)',
                               required=False, widget=forms.PasswordInput, strip=False)
    auto_levels = forms.BooleanField(label='คำนวณระดับชั้น +1 ถึง +5 อัตโนมัติจากระดับชั้นปัจจุบัน',
                                     required=False, initial=True)

    class Meta:
        model = User
        fields = ['camp1_code', 'camp2_code', 'full_name', 'nickname', 'school_name', 'birth_date',
                  'student_phone', 'parent_phone', 'email', 'chronic_disease', 'food_allergy', 'is_muslim',
                  'academic_year', 'current_level', 'year_plus_1', 'year_plus_2', 'year_plus_3', 'year_plus_4',
                  'year_plus_5', 'role', 'is_active',
                  'staff_can_view_users', 'staff_can_attendance', 'staff_can_signin_sheet',
                  'staff_can_manage_food', 'staff_can_scores', 'staff_can_exam_seating',
                  'note']
        widgets = {
            'birth_date': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'academic_year': forms.Select(),
            'note': forms.Textarea(attrs={'rows': 2}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        extra_year = self.instance.academic_year if (self.instance and self.instance.pk) else None
        years = _get_dynamic_academic_years(extra_year=extra_year)
        self.fields['academic_year'].widget.choices = [('', '-- เลือกปีการศึกษา --')] + [
            (y, f'ปีการศึกษา {y}') for y in years
        ]
        if self.instance and self.instance.pk:
            self.fields.pop('password', None)
            self.fields['full_name'].initial = self.instance.get_full_name()
        self.fields['camp1_code'].help_text = 'Member: รูปแบบ 69DS01 (ใช้เป็น Username ตลอดไป)'
        self.fields['camp2_code'].required = False
        self._bootstrap()

    def clean_camp1_code(self):
        return (self.cleaned_data.get('camp1_code') or '').strip()

    def clean_camp2_code(self):
        return (self.cleaned_data.get('camp2_code') or '').strip().upper() or None

    def save(self, commit=True):
        obj = super().save(commit=False)
        fn, ln = split_full_name(self.cleaned_data.get('full_name', ''))
        obj.first_name = fn
        obj.last_name = ln
        if not obj.camp_center and obj.camp1_code and len(obj.camp1_code) >= 4:
            center_part = obj.camp1_code[2:4].upper()
            if center_part in ('DS', 'KM'):
                obj.camp_center = center_part
        if commit:
            obj.save()
        return obj


class ProfileForm(BootstrapMixin, forms.ModelForm):
    """ฟอร์มแก้ไขโปรไฟล์ของ Member — แสดงเฉพาะฟิลด์ที่ Admin อนุญาต (ชื่อ-นามสกุลเป็นช่องเดียว)"""
    full_name = forms.CharField(label='ชื่อ-นามสกุล', required=False, max_length=300)

    class Meta:
        model = User
        fields = [
            'full_name', 'nickname', 'birth_date',
            'student_phone', 'parent_phone', 'email',
            'chronic_disease', 'food_allergy', 'is_muslim',
            'note',
        ]
        widgets = {
            'birth_date': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'nickname': forms.TextInput(attrs={'placeholder': 'ชื่อเล่น'}),
            'student_phone': forms.TextInput(attrs={'placeholder': 'เช่น 081-234-5678'}),
            'parent_phone': forms.TextInput(attrs={'placeholder': 'เช่น 089-876-5432'}),
            'email': forms.EmailInput(attrs={'placeholder': 'example@email.com'}),
            'chronic_disease': forms.TextInput(attrs={'placeholder': 'ระบุโรคประจำตัว (หากไม่มีเว้นว่าง)'}),
            'food_allergy': forms.TextInput(attrs={'placeholder': 'ระบุอาหารที่แพ้ เช่น กุ้ง ถั่ว นม (หากไม่มีเว้นว่าง)'}),
            'note': forms.Textarea(attrs={'rows': 2}),
        }

    def __init__(self, *args, allowed_fields=None, **kwargs):
        super().__init__(*args, **kwargs)
        allowed = set(allowed_fields or [])
        if self.instance and self.instance.pk:
            self.fields['full_name'].initial = self.instance.get_full_name()
        if 'first_name' not in allowed and 'last_name' not in allowed and 'full_name' not in allowed:
            self.fields.pop('full_name', None)
        for name in ('nickname', 'birth_date', 'student_phone', 'parent_phone', 'email',
                     'chronic_disease', 'food_allergy', 'is_muslim', 'note'):
            if name not in allowed:
                self.fields.pop(name, None)
        self._bootstrap()

    def save(self, commit=True):
        obj = super().save(commit=False)
        if 'full_name' in self.cleaned_data:
            fn, ln = split_full_name(self.cleaned_data.get('full_name', ''))
            obj.first_name = fn
            obj.last_name = ln
        if commit:
            obj.save()
        return obj


class SchoolRequestForm(BootstrapMixin, forms.Form):
    new_school_name = forms.CharField(label='ชื่อโรงเรียนใหม่', max_length=255)
    reason = forms.CharField(label='เหตุผล', required=False, widget=forms.Textarea(attrs={'rows': 2}))

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()


class SystemSettingForm(BootstrapMixin, forms.ModelForm):
    class Meta:
        model = SystemSetting
        fields = ['site_title', 'site_title_line2', 'auto_logout_minutes']
        widgets = {
            'site_title_line2': forms.TextInput(attrs={'placeholder': 'บรรทัดที่ 2 (ไม่บังคับ เช่น ศูนย์อบรม สอวน. วิชาคอมพิวเตอร์)'}),
            'auto_logout_minutes': forms.NumberInput(attrs={'min': 0, 'max': 1440}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()
        for f in self.fields.values():
            f.widget.attrs['form'] = 'master-settings-form'


class OpenRoundForm(BootstrapMixin, forms.Form):
    academic_year = forms.IntegerField(label='ปีการศึกษา (พ.ศ.)', min_value=2500, max_value=2700)
    camp_level = forms.TypedChoiceField(label='ค่าย', choices=CAMP_LEVEL_CHOICES, coerce=int)
    title = forms.CharField(label='ชื่อรอบ (ไม่บังคับ)', required=False, max_length=150)
    scheduled_open_at = forms.DateTimeField(
        label='เวลาเปิดอัตโนมัติ (ไม่บังคับ)', required=False,
        widget=forms.DateTimeInput(attrs={'type': 'datetime-local'}, format='%Y-%m-%dT%H:%M'),
    )
    scheduled_close_at = forms.DateTimeField(
        label='เวลาปิดอัตโนมัติ (ไม่บังคับ)', required=False,
        widget=forms.DateTimeInput(attrs={'type': 'datetime-local'}, format='%Y-%m-%dT%H:%M'),
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._bootstrap()


class ExcelUploadForm(BootstrapMixin, forms.Form):
    file = forms.FileField(label='ไฟล์ Excel (.xlsx) หรือ CSV (.csv)')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['file'].widget.attrs['accept'] = '.xlsx,.csv'
        self._bootstrap()

    def clean_file(self):
        f = self.cleaned_data['file']
        if not f.name.lower().endswith(('.xlsx', '.csv')):
            raise forms.ValidationError('รองรับเฉพาะไฟล์ .xlsx หรือ .csv')
        return f
