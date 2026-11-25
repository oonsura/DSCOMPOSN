"""
ทดสอบ Edge Cases ตามหัวข้อ 16 ของคู่มือ
รัน: python manage.py test
"""
import io
from datetime import date

from django.conf import settings
from django.test import Client, TestCase, override_settings
from django.urls import reverse
from openpyxl import Workbook

from . import levels, rights
from .excel import import_camp2, import_students
from .models import ArchiveRecord, AuditLog, RightsRound, SystemSetting, UseRecord, User


def make_member(code='69DS01', level='ม.4', year=2569, camp2=None, **kw):
    fname = kw.pop('first_name', 'สมชาย')
    lname = kw.pop('last_name', 'ใจดี')
    u = User(camp1_code=code, camp2_code=camp2, first_name=fname, last_name=lname, school_name='รร.ทดสอบ',
             camp_center=code[2:4], academic_year=year, current_level=level, role='member',
             is_first_login=False, **kw)
    levels.apply_levels(u)
    u.set_password(code)
    u.save()
    return u


def make_xlsx(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    buf.name = 'test.xlsx'
    return buf


class LevelTests(TestCase):
    def test_ม4(self):
        self.assertEqual(levels.next_levels('ม.4'), ['ม.5', 'ม.6', 'หมดสิทธิ์', 'หมดสิทธิ์', 'หมดสิทธิ์'])

    def test_ปวช(self):
        self.assertEqual(levels.next_levels('ปวช.1'), ['ปวช.2', 'ปวช.3', 'หมดสิทธิ์', 'หมดสิทธิ์', 'หมดสิทธิ์'])

    def test_ม1(self):
        self.assertEqual(levels.next_levels('ม.1'), ['ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'])

    def test_normalize(self):
        self.assertEqual(levels.normalize_level('ม 4'), 'ม.4')
        self.assertEqual(levels.normalize_level('มัธยมศึกษาปีที่ 5'), 'ม.5')
        self.assertEqual(levels.normalize_level('ปวช 2'), 'ปวช.2')
        self.assertIsNone(levels.normalize_level('ม.7'))


class DefaultsTests(TestCase):
    def test_fixtures(self):
        admin = User.objects.get(camp1_code='admin')
        self.assertEqual(admin.role, 'admin')
        self.assertTrue(admin.check_password('1234'))
        from food.models import AbsenceReason, WarningColor
        self.assertEqual(list(AbsenceReason.objects.values_list('name', flat=True)), ['ปกติ', 'ป่วย', 'ลากิจ', 'อื่นๆ'])
        self.assertEqual(WarningColor.objects.count(), 3)
        self.assertEqual(SystemSetting.get().max_absences, 5)

    def test_camp_code_validation(self):
        from django.core.exceptions import ValidationError
        u = User(camp1_code='ABC', role='member')
        with self.assertRaises(ValidationError):
            u.full_clean()

    def test_default_password_is_camp_code(self):
        u = User.objects.create_user('69DS09', role='member')
        self.assertTrue(u.check_password('69DS09'))
        self.assertTrue(u.is_first_login)


class RightsTests(TestCase):
    def setUp(self):
        self.admin = User.objects.get(camp1_code='admin')
        SystemSetting.get()

    def open(self, camp=1, year=2569):
        return rights.open_round(year, camp, self.admin)

    def test_confirm_and_code_sequence(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        b = make_member('68DS02', year=2568, level='ม.4')
        self.open()
        r1 = rights.confirm_rights(a)
        r2 = rights.confirm_rights(b)
        self.assertEqual(r1.rights_code, '68DS01')
        self.assertEqual(r2.rights_code, '68DS02')
        self.assertEqual(a.rights_used_count(), 1)

    def test_cannot_confirm_twice_same_year(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        self.open()
        rights.confirm_rights(a)
        with self.assertRaises(rights.RightsError):
            rights.confirm_rights(a)

    def test_no_round(self):
        with self.assertRaises(rights.RightsError):
            rights.confirm_rights(make_member('68DS01', year=2568, level='ม.4'))

    def test_sixth_time_blocked(self):
        a = make_member(level='ป.6', year=2564)  # 2564..2569 => ป.6, ม.1..ม.5 (ถึง 2569 = ม.5)
        for y in range(2564, 2569):
            ArchiveRecord.objects.create(user=a, academic_year=y, camp_level=1, rights_code=a.camp1_code)
        self.assertEqual(a.rights_used_count(), 5)
        self.open()
        with self.assertRaises(rights.RightsError) as cm:
            rights.confirm_rights(a)
        self.assertIn('ครบ', str(cm.exception))

    def test_m6_expired_next_year(self):
        # รุ่นปี 2568 อยู่ ม.5 -> ปี 2569 อยู่ ม.6 (ใช้ได้) -> ปี 2570 เกิน ม.6 (หมดสิทธิ์)
        a = make_member('68DS04', year=2568, level='ม.5')
        rr = self.open(year=2570)
        elig = rights.check_eligibility(a, rr)
        self.assertFalse(elig.ok)
        self.assertEqual(elig.code, 'expired')
        # ปี 2569 (ม.6) ยังใช้ได้
        rr2 = RightsRound.objects.create(academic_year=2569, camp_level=1, is_open=True)
        self.assertTrue(rights.check_eligibility(a, rr2).ok)

    def test_pvc3_boundary(self):
        # รุ่นปี 2568 อยู่ ปวช.2 -> ปี 2569 อยู่ ปวช.3 (ใช้ได้) -> ปี 2570 หมดสิทธิ์
        a = make_member('68DS05', year=2568, level='ปวช.2')
        self.assertTrue(rights.check_eligibility(a, self.open(year=2569)).ok)
        self.assertEqual(a.level_for_year(2570), 'หมดสิทธิ์')

    def test_camp2_cannot_use_camp1(self):
        a = make_member('68DS01', year=2568, level='ม.4', camp2='68KM01')
        self.open(camp=1)
        with self.assertRaises(rights.RightsError) as cm:
            rights.confirm_rights(a)
        self.assertIn('ค่าย 2', str(cm.exception))
        rr2 = self.open(camp=2)
        rec = rights.confirm_rights(a, rr2)
        self.assertEqual(rec.rights_code, '68KM01')

    def test_camp1_only_cannot_use_camp2(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        rr = self.open(camp=2)
        self.assertEqual(rights.check_eligibility(a, rr).code, 'no_camp2')

    def test_waive_when_not_confirmed(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        b = make_member('68DS02', year=2568, level='ม.4')
        rr = self.open()
        rights.confirm_rights(a)
        waived, path = rights.close_round(rr, self.admin)
        self.assertEqual(waived, 1)
        self.assertTrue(path.exists())
        self.assertEqual(path.name, 'ใช้สิทธิ์ค่าย1_ปีการศึกษา2569.xlsx')
        rec = UseRecord.objects.get(user=b)
        self.assertFalse(rec.is_used)
        self.assertEqual(b.rights_used_count(), 0)  # สละสิทธิ์ ไม่นับเป็นการใช้
        # ปิดรอบแล้ว ยืนยันไม่ได้
        with self.assertRaises(rights.RightsError):
            rights.confirm_rights(b, rr)
        path.unlink()

    def test_consecutive_rounds_codes_unique(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        b = make_member('68DS02', year=2568, level='ม.4')
        rr1 = self.open()
        r1 = rights.confirm_rights(a)
        rights.close_round(rr1, self.admin, mark_waived=False)
        rr2 = self.open()  # เปิดรอบต่อกันในปีเดียวกัน
        r2 = rights.confirm_rights(b)
        self.assertEqual(r1.rights_code, '68DS01')
        self.assertEqual(r2.rights_code, '68DS02')
        import os
        for f in settings.BACKUP_DIR.glob('ใช้สิทธิ์ค่าย1_ปีการศึกษา2569*.xlsx'):
            os.remove(f)

    def test_code_not_reused_after_soft_delete(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        b = make_member('68DS02', year=2568, level='ม.4')
        self.open()
        r1 = rights.confirm_rights(a)
        ArchiveRecord.objects.filter(use_record=r1).delete(user=self.admin)
        r1.delete(user=self.admin)
        r2 = rights.confirm_rights(b)
        self.assertNotEqual(r1.rights_code, r2.rights_code)
        # ผู้ใช้คนเดิมที่ถูกลบรายการใช้สิทธิ์ สามารถกลับมายืนยันด้วยรหัสค่ายตนเองได้อีกครั้ง
        r1_again = rights.confirm_rights(a)
        self.assertEqual(r1_again.rights_code, '68DS01')

    def test_only_one_open_round_per_camp(self):
        self.open()
        with self.assertRaises(rights.RightsError):
            self.open()

    def test_inactive_blocked(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        a.is_active = False
        a.save()
        self.open()
        with self.assertRaises(rights.RightsError):
            rights.confirm_rights(a)

    def test_audit_logged(self):
        a = make_member('68DS01', year=2568, level='ม.4')
        self.open()
        rights.confirm_rights(a)
        self.assertTrue(AuditLog.objects.filter(action='confirm_rights', user=a).exists())
        self.assertTrue(AuditLog.objects.filter(action='open_round').exists())

    def test_rights_codes_match_camp_codes_no_reds_rekm(self):
        c1_user = make_member('68DS10', year=2568, level='ม.4')
        c2_user = make_member('68DS20', year=2568, level='ม.4', camp2='68KM20')
        rr1 = self.open(camp=1, year=2569)
        rec1 = rights.confirm_rights(c1_user, rr1)
        self.assertEqual(rec1.rights_code, '68DS10')
        self.assertEqual(ArchiveRecord.objects.get(use_record=rec1).rights_code, '68DS10')

        rr2 = self.open(camp=2, year=2569)
        rec2 = rights.confirm_rights(c2_user, rr2)
        self.assertEqual(rec2.rights_code, '68KM20')
        self.assertEqual(ArchiveRecord.objects.get(use_record=rec2).rights_code, '68KM20')

        # นำเข้าผู้ใช้สิทธิ์ค่าย 2 ด้วยรหัสค่าย 2 (camp2_code)
        c2_import_user = make_member('68DS21', year=2568, level='ม.4', camp2='68KM21')
        preview = rights.preview_camp1_rights_import(['68KM21'], academic_year=2569, camp_level=2)
        self.assertEqual(preview['valid_count'], 1)
        res = rights.import_camp1_rights(['68KM21'], self.admin, academic_year=2569, camp_level=2)
        self.assertEqual(res['imported'], 1)
        rec_imp = UseRecord.objects.get(user=c2_import_user, academic_year=2569)
        self.assertEqual(rec_imp.rights_code, '68KM21')
        self.assertEqual(ArchiveRecord.objects.get(use_record=rec_imp).rights_code, '68KM21')


class LoginTests(TestCase):
    def setUp(self):
        self.u = make_member('69DS01', is_first_login=True) if False else make_member('69DS01')
        self.u.is_first_login = True
        self.u.save()

    def test_first_login_forces_password_change(self):
        c = Client()
        r = c.post(reverse('composn:login'), {'username': '69DS01', 'password': '69DS01'})
        self.assertRedirects(r, reverse('composn:change_password'))
        r = c.get(reverse('composn:profile'))
        self.assertRedirects(r, reverse('composn:change_password'), fetch_redirect_response=False)
        r = c.post(reverse('composn:change_password'), {'old_password': '69DS01', 'new_password1': 'newpass99',
                                                        'new_password2': 'newpass99'})
        self.assertEqual(r.status_code, 302)
        self.u.refresh_from_db()
        self.assertFalse(self.u.is_first_login)
        self.assertEqual(c.get(reverse('composn:profile')).status_code, 200)

    def test_lowercase_username_ok(self):
        c = Client()
        r = c.post(reverse('composn:login'), {'username': '69ds01', 'password': '69DS01'})
        self.assertEqual(r.status_code, 302)

    def test_first_login_lowercase_camp1_password_and_forced_change(self):
        c = Client()
        r = c.post(reverse('composn:login'), {'username': '69ds01', 'password': '69ds01'})
        self.assertRedirects(r, reverse('composn:change_password'))
        self.assertRedirects(c.get(reverse('composn:profile')), reverse('composn:change_password'), fetch_redirect_response=False)
        r_chg = c.post(reverse('composn:change_password'), {
            'old_password': '69ds01',
            'new_password1': 'NewPass69!',
            'new_password2': 'NewPass69!',
        })
        self.assertEqual(r_chg.status_code, 302)
        self.u.refresh_from_db()
        self.assertFalse(self.u.is_first_login)
        self.assertEqual(c.get(reverse('composn:profile')).status_code, 200)
        # หลังเปลี่ยนรหัสผ่านแล้ว ต้องใช้รหัสเดิม (69ds01 / 69DS01) เข้าไม่ได้อีก
        c_old = Client()
        self.assertEqual(
            c_old.post(reverse('composn:login'), {'username': '69DS01', 'password': '69ds01'}).status_code, 200
        )

    def test_second_login_kicks_first(self):
        self.u.is_first_login = False
        self.u.save()
        c1, c2 = Client(), Client()
        c1.post(reverse('composn:login'), {'username': '69DS01', 'password': '69DS01'})
        self.assertEqual(c1.get(reverse('composn:profile')).status_code, 200)
        c2.post(reverse('composn:login'), {'username': '69DS01', 'password': '69DS01'})
        self.assertEqual(c2.get(reverse('composn:profile')).status_code, 200)
        r = c1.get(reverse('composn:profile'))
        self.assertEqual(r.status_code, 302)
        self.assertIn('/login/', r['Location'])

    def test_auto_logout_setting(self):
        self.assertEqual(settings.SESSION_COOKIE_AGE, 900)
        self.assertFalse(settings.SESSION_SAVE_EVERY_REQUEST)
        self.assertEqual(settings.SESSION_ENGINE, 'django.contrib.sessions.backends.cached_db')

    def test_admin_reset_password(self):
        admin = Client()
        admin.post(reverse('composn:login'), {'username': 'admin', 'password': '1234'})
        admin.post(reverse('composn:change_password'), {'old_password': '1234', 'new_password1': 'adminpw1',
                                                        'new_password2': 'adminpw1'})
        self.u.set_password('changed!')
        self.u.is_first_login = False
        self.u.save()
        r = admin.post(reverse('composn:admin_user_reset_password', args=[self.u.pk]))
        self.assertEqual(r.status_code, 302)
        self.u.refresh_from_db()
        self.assertTrue(self.u.check_password('69DS01'))
        self.assertTrue(self.u.is_first_login)

    def test_member_cannot_access_admin(self):
        self.u.is_first_login = False
        self.u.save()
        c = Client()
        c.post(reverse('composn:login'), {'username': '69DS01', 'password': '69DS01'})
        self.assertEqual(c.get(reverse('composn:admin_users')).status_code, 403)
        self.assertEqual(c.get(reverse('food:staff_attendance')).status_code, 403)

    def test_login_with_camp2_code(self):
        m = make_member('68DS01', year=2568, level='ม.4', camp2='68KM01')
        m.is_first_login = False
        m.set_password('secret99')
        m.save()

        # Login ด้วย camp1_code
        c1 = Client()
        r1 = c1.post(reverse('composn:login'), {'username': '68DS01', 'password': 'secret99'})
        self.assertRedirects(r1, reverse('composn:home'), fetch_redirect_response=False)
        self.assertEqual(int(c1.session['_auth_user_id']), m.pk)
        self.assertEqual(c1.get(reverse('composn:profile')).status_code, 200)

        # Login ด้วย camp2_code (ตัวใหญ่) -> เข้าบัญชีเดียวกัน และเตะ session เดิม
        c2 = Client()
        r2 = c2.post(reverse('composn:login'), {'username': '68KM01', 'password': 'secret99'})
        self.assertRedirects(r2, reverse('composn:home'), fetch_redirect_response=False)
        self.assertEqual(int(c2.session['_auth_user_id']), m.pk)
        self.assertEqual(c2.get(reverse('composn:profile')).status_code, 200)
        r1_after = c1.get(reverse('composn:profile'))
        self.assertEqual(r1_after.status_code, 302)
        self.assertIn('/login/', r1_after['Location'])

        # Login ด้วย camp2_code (ตัวพิมพ์เล็ก)
        c3 = Client()
        r3 = c3.post(reverse('composn:login'), {'username': '68km01', 'password': 'secret99'})
        self.assertRedirects(r3, reverse('composn:home'), fetch_redirect_response=False)
        self.assertEqual(int(c3.session['_auth_user_id']), m.pk)
        self.assertEqual(c3.get(reverse('composn:profile')).status_code, 200)

        # รหัสผ่านผิด หรือรหัสค่าย 2 ที่ไม่มีในระบบ -> เข้าไม่ได้
        c_bad = Client()
        self.assertEqual(
            c_bad.post(reverse('composn:login'), {'username': '68KM01', 'password': 'wrong'}).status_code, 200
        )
        self.assertEqual(
            c_bad.post(reverse('composn:login'), {'username': '68KM99', 'password': 'secret99'}).status_code, 200
        )

        # ผู้ใช้ที่ปิดใช้งาน (is_active=False) เข้าด้วย camp2_code ไม่ได้
        m.is_active = False
        m.save()
        c_inact = Client()
        self.assertEqual(
            c_inact.post(reverse('composn:login'), {'username': '68KM01', 'password': 'secret99'}).status_code, 200
        )


class ImportTests(TestCase):
    H = ['camp1_code', 'first_name', 'last_name', 'school_name', 'birth_date', 'camp_center', 'academic_year',
         'current_level']

    def test_create_and_overwrite(self):
        f = make_xlsx(self.H, [['69DS01', 'A', 'B', 'S1', '15/03/2555', 'DS', 2569, 'ม.4']])
        res = import_students(f)
        self.assertEqual((res['created'], res['updated']), (1, 0))
        u = User.objects.get(camp1_code='69DS01')
        self.assertTrue(u.check_password('69DS01'))
        self.assertEqual(u.birth_date, date(2012, 3, 15))
        self.assertEqual((u.year_plus_1, u.year_plus_3), ('ม.5', 'หมดสิทธิ์'))
        # ซ้ำ: ทับข้อมูล
        f = make_xlsx(self.H, [['69DS01', 'A2', 'B2', 'S2', None, 'DS', 2569, 'ม.5']])
        res = import_students(f)
        self.assertEqual((res['created'], res['updated']), (0, 1))
        u.refresh_from_db()
        self.assertEqual((u.first_name, u.school_name, u.year_plus_1), ('A2', 'โรงเรียนS2', 'ม.6'))
        self.assertEqual(User.objects.filter(camp1_code='69DS01').count(), 1)

    def test_bad_format_rows_reported(self):
        f = make_xlsx(self.H, [['BAD', 'A', 'B', 'S', None, 'DS', 2569, 'ม.4'],
                               ['69DS02', 'A', 'B', 'S', None, 'XX', 2569, 'ม.4'],
                               ['69DS03', 'A', 'B', 'S', None, 'DS', 2569, 'ม.9'],
                               ['69DS04', 'A', 'B', 'S', 'not-a-date', 'DS', 2569, 'ม.4'],
                               ['69DS05', 'A', 'B', 'S', None, 'DS', 2569, 'ม.4']])
        res = import_students(f)
        self.assertEqual(res['created'], 1)
        self.assertEqual(len(res['errors']), 4)

    def test_missing_columns(self):
        with self.assertRaises(ValueError):
            import_students(make_xlsx(['camp1_code'], [['69DS01']]))

    def test_camp2_import(self):
        make_member('69DS01')
        make_member('69DS02')
        res = import_camp2(make_xlsx(['camp1_code', 'camp2_code'], [
            ['69DS01', '69KM01'],      # ok
            ['69DS99', '69KM02'],      # ไม่มี camp1_code ตรงกัน
            ['69DS02', '69KM01'],      # camp2 ซ้ำ
            ['69DS02', 'bad'],         # format ผิด
        ]))
        self.assertEqual(res['updated'], 1)
        self.assertEqual(len(res['errors']), 3)
        self.assertEqual(User.objects.get(camp1_code='69DS01').camp2_code, '69KM01')
        self.assertEqual(User.objects.get(camp1_code='69DS01').camp_level, 2)

    def test_headerless_composn_excel_and_preview(self):
        """รองรับไฟล์รูปแบบ COMPOSN เดิม (ไม่มีหัวตาราง 14 คอลัมน์) และการ Preview ก่อนยืนยัน"""
        from .excel import preview_students
        wb = Workbook()
        ws = wb.active
        ws.append(['66DS01', 'ด.ช.ณฐกร', 'จันทรโชติ', 'สาธิต ปทุมวัน', 15, 3, 2555, 1, 'ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'])
        ws.append(['66DS02', 'สมหญิง', 'รักเรียน', 'สวนกุหลาบ', None, None, None, 1, 'ม.4', 'ม.5', 'ม.6', None, None, None])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        buf.name = 'legacy.xlsx'
        prev = preview_students(buf)
        self.assertEqual(prev['valid_count'], 2)
        res = import_students(buf)
        self.assertEqual(res['created'], 2)
        u1 = User.objects.get(camp1_code='66DS01')
        self.assertEqual(u1.academic_year, 2566)
        self.assertEqual(u1.camp_center, 'DS')
        self.assertEqual(u1.year_plus_5, 'ม.6')

    def test_thai_template_workbook_roundtrip(self):
        """ไฟล์ตัวอย่าง Template แบบภาษาไทย ดาวน์โหลดแล้วนำกลับมา Import ได้ทันที"""
        from .excel import camp2_template_workbook, students_template_workbook
        buf = io.BytesIO()
        students_template_workbook(lang='th').save(buf)
        buf.seek(0)
        buf.name = 'ตัวอย่างไฟล์นำเข้านักเรียน.xlsx'
        res = import_students(buf)
        self.assertEqual(res['created'], 2)
        self.assertEqual(len(res['errors']), 0)
        self.assertTrue(User.objects.filter(camp1_code='69DS01', first_name='ด.ช.สมชาย').exists())

        buf2 = io.BytesIO()
        camp2_template_workbook(lang='th').save(buf2)
        buf2.seek(0)
        buf2.name = 'ตัวอย่างไฟล์นำเข้ารหัสค่าย2.xlsx'
        res2 = import_camp2(buf2)
        self.assertEqual(res2['updated'], 2)

    def test_headerless_8col_standard_and_14col_mixed(self):
        """รองรับไฟล์ที่ไม่มีหัวตารางทั้งแบบ 8 คอลัมน์ (รหัส,ชื่อ,สกุล,รร,วันเกิด,ศูนย์,ปี,ระดับชั้น) และแบบ 14 คอลัมน์"""
        wb = Workbook()
        ws = wb.active
        # แถวที่ 1: แบบ 8 คอลัมน์ไม่มีหัวตาราง (ระดับชั้นอยู่ที่คอลัมน์ 8)
        ws.append(['66DS01', 'ด.ช.ณัชพล', 'ตั้งศรีสกุล', 'สามเสนวิทยาลัย', None, 'DS', 2566, 'ม.2'])
        # แถวที่ 2: แบบ ปวช. เว้นวรรคในรูปแบบ 8 คอลัมน์
        ws.append(['68DS48', 'นางสาวภิญญาพัชญ์', 'ขมหวาน', 'สถาบันโคเซ็น สจล.', None, 'DS', 2568, 'ปวช. 1'])
        # แถวที่ 3: แบบ 14 คอลัมน์ COMPOSN เดิม (ระดับชั้นอยู่ที่คอลัมน์ 9)
        ws.append(['67DS01', 'นายเทรุกิ ณัฐนนท์', 'ซาโต้', 'สถาบันโคเซ็น สจล.', None, None, None, 1, 'ปวช.1', 'ปวช.2', 'ปวช.3', None, None, None])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        buf.name = 'mixed.xlsx'
        res = import_students(buf)
        self.assertEqual(res['created'], 3)
        self.assertEqual(User.objects.get(camp1_code='66DS01').current_level, 'ม.2')
        self.assertEqual(User.objects.get(camp1_code='66DS01').year_plus_1, 'ม.3')
        self.assertEqual(User.objects.get(camp1_code='68DS48').current_level, 'ปวช.1')
        self.assertEqual(User.objects.get(camp1_code='68DS48').year_plus_2, 'ปวช.3')
        self.assertEqual(User.objects.get(camp1_code='67DS01').current_level, 'ปวช.1')

    def test_headerless_3col_sets_camp1_code_password_not_fullname(self):
        """นำเข้านักเรียนแบบไม่มีหัวตาราง 2-5 คอลัมน์ ต้องตั้งรหัสผ่านเริ่มต้นเป็นรหัสค่าย 1 (ไม่ใช่ชื่อ-นามสกุล)"""
        wb = Workbook()
        ws = wb.active
        ws.append(['69DS22', 'นายกวิน เพียรไพศาลสกุล', 'โรงเรียนเทพศิรินทร์'])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        buf.name = '3col_students.xlsx'
        res = import_students(buf)
        self.assertEqual(res['created'], 1)
        u = User.objects.get(camp1_code='69DS22')
        self.assertEqual(u.first_name, 'นายกวิน')
        self.assertEqual(u.last_name, 'เพียรไพศาลสกุล')
        self.assertEqual(u.school_name, 'โรงเรียนเทพศิรินทร์')
        self.assertTrue(u.is_first_login)
        self.assertTrue(u.check_password('69DS22'))
        self.assertTrue(u.check_password('69ds22'))
        self.assertFalse(u.check_password('นายกวิน เพียรไพศาลสกุล'))

        # Login ด้วย 69DS22 ต้องเข้าได้และบังคับไปเปลี่ยนรหัสผ่าน
        c = Client()
        r_log = c.post(reverse('composn:login'), {'username': '69DS22', 'password': '69DS22'})
        self.assertRedirects(r_log, reverse('composn:change_password'))

        # หากผู้ใช้เดิมยังเป็น is_first_login=True แต่รหัสผ่านเพี้ยนอยู่ การนำเข้าซ้ำต้องซ่อมกลับเป็น camp1_code
        u.set_password('นายกวิน เพียรไพศาลสกุล')
        u.save()
        buf.seek(0)
        res2 = import_students(buf)
        self.assertEqual(res2['updated'], 1)
        u.refresh_from_db()
        self.assertTrue(u.check_password('69DS22'))

        # แต่หากเคยเปลี่ยนรหัสผ่านแล้ว (is_first_login=False) การนำเข้าซ้ำห้ามทับรหัสผ่านผู้ใช้
        u.set_password('CustomPass99!')
        u.is_first_login = False
        u.save()
        buf.seek(0)
        import_students(buf)
        u.refresh_from_db()
        self.assertTrue(u.check_password('CustomPass99!'))
        self.assertFalse(u.check_password('69DS22'))


class AttendanceTests(TestCase):
    def setUp(self):
        from food.models import Attendance  # noqa
        self.staff = User.objects.create_user('staff1', password='x', role='staff', is_first_login=False)
        self.m = make_member()

    def test_absence_count_and_warning(self):
        from food.models import absence_status
        from food.services import save_attendance
        for i in range(1, 5):
            save_attendance(self.m, date(2026, 7, i), 'morning', 'absent', 'ป่วย', recorded_by=self.staff)
        st = absence_status(self.m)
        self.assertEqual(st['count'], 4)
        self.assertTrue(st['near_max'])
        self.assertEqual(st['text'], 'ขาดอบรม 4 จาก 5 ภาคเวลา (ใกล้ถึงจำนวนที่ขาดอบรมสูงสุด)')
        self.assertEqual(st['color'], '#fd7e14')  # สีส้ม (>=3)
        save_attendance(self.m, date(2026, 7, 5), 'morning', 'absent', recorded_by=self.staff)
        st = absence_status(self.m)
        self.assertTrue(st['reached'])
        self.assertEqual(st['color'], '#dc3545')

    def test_change_present_absent_both_ways(self):
        from food.models import Attendance, absence_status
        from food.services import save_attendance
        d = date(2026, 7, 1)
        save_attendance(self.m, d, 'morning', 'present', recorded_by=self.staff)
        self.assertEqual(absence_status(self.m)['count'], 0)
        save_attendance(self.m, d, 'morning', 'absent', 'ลากิจ', recorded_by=self.staff)
        self.assertEqual(absence_status(self.m)['count'], 1)
        save_attendance(self.m, d, 'morning', 'present', recorded_by=self.staff)
        self.assertEqual(absence_status(self.m)['count'], 0)
        self.assertEqual(Attendance.objects.filter(user=self.m).count(), 1)  # ไม่สร้างซ้ำ

    def test_soft_delete_attendance(self):
        from food.models import Attendance, absence_status
        from food.services import delete_attendance, save_attendance
        att, _ = save_attendance(self.m, date(2026, 7, 1), 'morning', 'absent', recorded_by=self.staff)
        self.assertEqual(absence_status(self.m)['count'], 1)
        delete_attendance(att, self.staff)
        self.assertEqual(absence_status(self.m)['count'], 0)
        self.assertEqual(Attendance.all_objects.filter(pk=att.pk, is_deleted=True).count(), 1)
        # บันทึกใหม่ในวัน/ช่วงเดิมได้
        save_attendance(self.m, date(2026, 7, 1), 'morning', 'late', recorded_by=self.staff)
        self.assertEqual(Attendance.objects.filter(user=self.m).count(), 1)


class ScoreTests(TestCase):
    def test_total_formula(self):
        from decimal import Decimal
        from food.models import ScoreCategory, total_score
        from food.services import save_score
        year = SystemSetting.get().current_academic_year
        c1 = ScoreCategory.objects.create(name='สอบ', weight=Decimal('60'), academic_year=year)
        c2 = ScoreCategory.objects.create(name='งาน', weight=Decimal('40'), academic_year=year)
        m = make_member()
        save_score(m, c1, 1, Decimal('45'), Decimal('50'))   # 0.9*60 = 54
        save_score(m, c2, 1, Decimal('15'), Decimal('20'))   # 0.75*40 = 30
        total, _ = total_score(m)
        self.assertEqual(total, Decimal('84.00'))

    def test_score_over_max_rejected(self):
        from decimal import Decimal
        from food.models import ScoreCategory
        from food.services import FoodError, save_score
        c = ScoreCategory.objects.create(name='x', weight=10, academic_year=2569)
        with self.assertRaises(FoodError):
            save_score(make_member(), c, 1, Decimal('11'), Decimal('10'))


class DisplayTests(TestCase):
    def test_requires_login_and_no_sensitive_data(self):
        admin = User.objects.get(camp1_code='admin')
        a = make_member('68DS01', year=2568, level='ม.4')
        b = make_member('69DS02', first_name='วิชัย', last_name='รักดี')
        a.birth_date = date(2010, 1, 1)
        a.save()
        rights.open_round(2569, 1, admin)
        rights.confirm_rights(a)
        c = Client()
        # ต้องเข้าสู่ระบบก่อนจึงจะดูหน้าแสดงผลได้
        self.assertEqual(c.get('/display/').status_code, 302)
        c.force_login(b)
        for url in ('/display/', '/display/ds/', '/display/km/'):
            r = c.get(url)
            self.assertEqual(r.status_code, 200, url)
        html = c.get('/display/ds/').content.decode()
        self.assertIn('68DS01', html)          # แสดงรหัสค่ายที่ใช้สิทธิ์
        self.assertNotIn('ใจดี', html)          # ไม่แสดงนามสกุลเต็มของผู้อื่น
        self.assertNotIn('รร.ทดสอบ', html)      # ไม่แสดงโรงเรียน
        self.assertNotIn('2553', html)         # ไม่แสดงวันเกิด
        self.assertNotIn('68DS01', c.get('/display/km/').content.decode())

    def test_display_download_admin_only(self):
        self.assertEqual(Client().get('/display/download/').status_code, 302)


class SoftDeleteTests(TestCase):
    def test_use_record_soft_delete(self):
        admin = User.objects.get(camp1_code='admin')
        a = make_member('68DS01', year=2568, level='ม.4')
        rights.open_round(2569, 1, admin)
        rec = rights.confirm_rights(a)
        ArchiveRecord.objects.filter(use_record=rec).delete(user=admin)
        rec.delete(user=admin)
        self.assertEqual(a.rights_used_count(), 0)
        self.assertTrue(UseRecord.all_objects.filter(pk=rec.pk, is_deleted=True).exists())
        self.assertFalse(UseRecord.objects.filter(pk=rec.pk).exists())


class SchoolRequestTests(TestCase):
    def test_one_request_until_cleared(self):
        m = make_member()
        c = Client()
        c.post(reverse('composn:login'), {'username': '69DS01', 'password': '69DS01'})
        c.post(reverse('composn:school_request'), {'new_school_name': 'ใหม่1'})
        c.post(reverse('composn:school_request'), {'new_school_name': 'ใหม่2'})
        self.assertEqual(m.school_requests.count(), 1)
        admin = Client()
        admin.post(reverse('composn:login'), {'username': 'admin', 'password': '1234'})
        admin.post(reverse('composn:change_password'), {'old_password': '1234', 'new_password1': 'adminpw1',
                                                        'new_password2': 'adminpw1'})
        req = m.school_requests.first()
        admin.post(reverse('composn:admin_school_requests'), {'id': req.pk, 'action': 'approve'})
        m.refresh_from_db()
        self.assertEqual(m.school_name, 'โรงเรียนใหม่1')
        admin.post(reverse('composn:admin_school_requests'), {'id': req.pk, 'action': 'clear'})
        c.post(reverse('composn:school_request'), {'new_school_name': 'ใหม่2'})
        self.assertEqual(m.school_requests.count(), 2)

class SmokeTests(TestCase):
    """เปิดทุกหน้าด้วยแต่ละบทบาท ต้องไม่ error 500"""

    def test_all_pages_render(self):
        from food.models import FoodMenu, FoodOrder
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        staff = User.objects.create_user('staff1', password='x', role='staff', is_first_login=False)
        m = make_member('68DS01', year=2568, level='ม.4', camp2='68KM01')
        make_member('69DS02')
        rr = rights.open_round(2569, 2, admin_u)
        rights.confirm_rights(m)
        o = FoodOrder.objects.create(food_date=date.today(), restaurant_name='ร้าน', max_selection_count=2,
                                     allow_current_students=True, allow_rights_students=True)
        FoodMenu.objects.create(food_order=o, menu_name='ข้าวผัด')
        admin_pages = [reverse(n) for n in (
            'composn:admin_dashboard', 'composn:admin_users', 'composn:admin_users_settings', 'composn:admin_admins',
            'composn:admin_admin_create', 'composn:admin_user_create', 'composn:admin_import_students',
            'composn:admin_import_camp2', 'composn:admin_calculate_level', 'composn:admin_rounds',
            'composn:admin_rights_settings', 'composn:admin_use_records',
            'composn:admin_school_requests', 'composn:admin_settings', 'composn:admin_backups', 'composn:admin_exports',
            'composn:admin_audit_logs', 'composn:admin_trash', 'composn:admin_template_students',
            'composn:admin_template_camp2', 'composn:display_download', 'composn:admin_year_report',
            'food:staff_attendance', 'food:attendance_summary', 'food:signin_sheet',
            'food:staff_scores', 'food:score_summary', 'food:score_announcement',
            'food:score_categories', 'food:food_settings', 'food:orders', 'food:order_create',
            'composn:display', 'composn:home')]
        admin_pages += [
            reverse('food:signin_sheet') + '?mode=exam&camp=2',
            reverse('composn:admin_user_detail', args=[m.pk]), reverse('composn:admin_user_edit', args=[m.pk]),
            reverse('composn:admin_round_detail', args=[rr.pk]), reverse('food:attendance_history', args=[m.pk]),
            reverse('food:order_detail', args=[o.pk]), reverse('food:order_edit', args=[o.pk]),
            reverse('food:order_copy', args=[o.pk]), reverse('food:order_report', args=[o.pk]),
            reverse('food:order_print', args=[o.pk]), reverse('food:order_export', args=[o.pk]),
            reverse('composn:admin_exports') + '?kind=students', reverse('composn:admin_exports') + '?kind=all_history',
            reverse('composn:admin_exports') + '?kind=history&code=69DS01',
            reverse('composn:admin_exports') + '?kind=absence_monthly&month=2026-07',
            reverse('composn:admin_users') + '?export=xlsx', reverse('composn:admin_round_detail', args=[rr.pk]) + '?export=xlsx',
            reverse('food:attendance_summary') + '?export=xlsx', reverse('food:score_summary') + '?export=xlsx',
            '/display/ds/', '/display/km/']
        c = Client()
        c.force_login(admin_u)
        for url in admin_pages:
            r = c.get(url)
            self.assertIn(r.status_code, (200, 302), f'{url} -> {r.status_code}')
        # staff
        c = Client()
        c.force_login(staff)
        for url in (reverse('food:staff_attendance'), reverse('food:staff_scores'), reverse('food:orders'),
                    reverse('composn:home')):
            self.assertIn(c.get(url).status_code, (200, 302), url)
        # member
        c = Client()
        c.force_login(m)
        for url in (reverse('composn:profile'), reverse('food:my_food'), reverse('food:my_food_detail', args=[o.pk]),
                    reverse('composn:my_history_export')):
            self.assertEqual(c.get(url).status_code, 200, url)
        self.assertRedirects(c.get(reverse('food:my_selections')), reverse('food:my_food'))
        # ลงเวลาผ่าน API
        c = Client()
        c.force_login(staff)
        r = c.post(reverse('food:attendance_save'), {'user_id': m.pk, 'date': '2026-07-01', 'period': 'morning',
                                                      'status': 'absent', 'reason': 'ป่วย'})
        self.assertEqual(r.json()['absence']['count'], 1)
        # member เลือกอาหาร
        c = Client()
        c.force_login(m)
        mid = o.menus.first().pk
        c.post(reverse('food:my_food_detail', args=[o.pk]), {f'qty_{mid}': '1'})
        self.assertEqual(o.selections.count(), 1)

    def test_score_visibility_and_training_date_range(self):
        from decimal import Decimal
        from food.models import ScoreCategory, TrainingDate
        from food.services import save_score
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        m = make_member('69DS10')
        cat_visible = ScoreCategory.objects.create(name='สอบย่อย 1', weight=Decimal('40'), academic_year=2569,
                                                   is_visible=True)
        cat_hidden = ScoreCategory.objects.create(name='สอบปลายค่าย', weight=Decimal('60'), academic_year=2569,
                                                  is_visible=False)
        save_score(m, cat_visible, 1, Decimal('80'), Decimal('100'), recorded_by=admin_u)
        save_score(m, cat_hidden, 1, Decimal('90'), Decimal('100'), recorded_by=admin_u)

        # หน้าข้อมูลการอบรมและคะแนนของนักเรียนเห็นเฉพาะหมวดที่ is_visible=True
        mc = Client()
        mc.force_login(m)
        prof_html = mc.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('สอบย่อย 1', prof_html)
        self.assertNotIn('สอบปลายค่าย', prof_html)

        # Admin สร้างช่วงวันที่อบรม 3 วันรวด (2026-10-10 ถึง 2026-10-12)
        ac = Client()
        ac.force_login(admin_u)
        ac.post(reverse('food:score_categories'), {
            'action': 'add_training_date',
            'year': '2569',
            'date_start': '2026-10-10',
            'date_end': '2026-10-12',
            'event_type': 'training',
            'title': 'อบรมค่าย 1 ช่วงเช้า-บ่าย',
        })
        self.assertEqual(TrainingDate.objects.filter(academic_year=2569).count(), 3)

        # Admin สร้างหัวข้อใหญ่ (คะแนนสอบทฤษฎี) และหัวข้อย่อย (ข้อ1, ข้อ2, ข้อ3, ข้อ4)
        ac.post(reverse('food:score_categories'), {
            'action': 'create',
            'year': '2569',
            'group_type': ScoreCategory.GROUP_THEORY,
            'name': 'คะแนนสอบทฤษฎี',
            'weight': '50',
            'default_max_score': '100',
            'order': '2',
            'is_visible': '1',
        })
        main_theory = ScoreCategory.objects.get(academic_year=2569, name='คะแนนสอบทฤษฎี', parent__isnull=True)
        ac.post(reverse('food:score_categories'), {
            'action': 'create_sub',
            'year': '2569',
            'parent_id': str(main_theory.pk),
            'name': 'ข้อ1, ข้อ2, ข้อ3, ข้อ4',
            'weight': '0',
            'default_max_score': '25',
            'order': '1',
            'is_visible': '1',
        })
        subs = list(main_theory.subcategories.order_by('order', 'id'))
        self.assertEqual([s.name for s in subs], ['ข้อ1', 'ข้อ2', 'ข้อ3', 'ข้อ4'])
        for s in subs:
            save_score(m, s, 1, Decimal('20'), Decimal('25'), recorded_by=admin_u)
        prof_html2 = mc.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('คะแนนสอบทฤษฎี', prof_html2)
        self.assertIn('ข้อ1', prof_html2)

        # ทดสอบบันทึกที่ปุ่มหัวข้อใหญ่ที่เดียว (อัปเดตทั้งหัวข้อใหญ่และหัวข้อย่อยทั้งหมดพร้อมกัน)
        post_main = {
            'action': 'update',
            'id': str(main_theory.pk),
            'year': '2569',
            'name': 'คะแนนสอบทฤษฎี',
            'group_type': ScoreCategory.GROUP_THEORY,
            'weight': '50',
            'default_max_score': '100',
            'order': '2',
            # ไม่ส่ง is_visible และ is_open ของทั้งหัวข้อใหญ่และหัวข้อย่อย = ปิดทั้งหมด
        }
        for idx, s in enumerate(subs, 1):
            post_main[f'sub_name_{s.pk}'] = s.name
            post_main[f'sub_score_{s.pk}'] = '15'
            post_main[f'sub_order_{s.pk}'] = str(idx)
        ac.post(reverse('food:score_categories'), post_main)
        main_theory.refresh_from_db()
        self.assertFalse(main_theory.is_visible)
        self.assertFalse(main_theory.is_open)
        self.assertEqual(main_theory.default_max_score, Decimal('60'))  # 4 ข้อย่อย x 15 คะแนน = 60
        for s in main_theory.subcategories.all():
            self.assertEqual(s.default_max_score, Decimal('15'))
            self.assertFalse(s.is_visible)
            self.assertFalse(s.is_open)

        # ทดสอบกรอกคะแนนแบบตาราง Excel พร้อมกันทุกหัวข้อย่อย (score_<user_id>_<sub_id>)
        main_theory.is_open = True
        main_theory.save()
        main_theory.subcategories.update(is_open=True)
        grid_post = {
            'action': 'save_manual',
            'main_id': str(main_theory.pk),
            'exam_no': '1',
        }
        for s in subs:
            grid_post[f'score_{m.pk}_{s.pk}'] = '12'
        ac.post(reverse('food:staff_scores'), grid_post)
        for s in subs:
            self.assertEqual(s.scores.get(user=m, exam_no=1).score, Decimal('12'))

        # ทดสอบค่าเริ่มต้นเวลาเรียน = เข้าอบรม (present) สำหรับวันที่มาถึงแล้ว และไม่มีปุ่มเลือกทั้งหมด/เอาออกทั้งหมด
        att_html = ac.get(reverse('food:staff_attendance') + '?year=2569&date=2026-10-10&period=morning').content.decode('utf-8')
        from food.models import Attendance
        self.assertTrue(Attendance.objects.filter(user=m, date='2026-10-10', session_period='morning', status='present').exists())
        self.assertNotIn('mark_all_present_force', att_html)
        self.assertNotIn('clear_all_attendance', att_html)

        # ทดสอบว่าหน้ากรอกคะแนนปี 2569 แสดงเฉพาะนักเรียนปี 69 (69DS10) และผู้ขอใช้สิทธิ์ปี 2569 หลังปิดรอบการใช้สิทธิ์
        m_old_unused = make_member('68DS01', year=2568, level='ม.2')
        m_old_used = make_member('68DS02', year=2568, level='ม.2')
        rr69 = rights.open_round(2569, 1, admin_u)
        rights.confirm_rights(m_old_used, rr69)

        # นักเรียนที่ยืนยันใช้สิทธิ์ (68DS02) จะถูกดึงเข้าตารางคะแนนปี 2569 ทันที ส่วนผู้ไม่ได้ใช้สิทธิ์ (68DS01) จะไม่แสดง
        scores_html_before = ac.get(reverse('food:staff_scores') + '?year=2569').content.decode('utf-8')
        self.assertIn('69DS10', scores_html_before)
        self.assertNotIn('68DS01', scores_html_before)
        self.assertIn('68DS02', scores_html_before)

        # หลังปิดระบบการใช้สิทธิ์: นักเรียนปีเก่าที่ใช้สิทธิ์ (68DS02) ยังคงอยู่ในหน้ากรอกคะแนนปี 2569 ส่วน 68DS01 ที่สละสิทธิ์ไม่แสดง
        rights.close_round(rr69, admin_u, mark_waived=True)
        scores_html_after = ac.get(reverse('food:staff_scores') + '?year=2569').content.decode('utf-8')
        self.assertIn('69DS10', scores_html_after)
        self.assertIn('68DS02', scores_html_after)
        self.assertNotIn('68DS01', scores_html_after)

    def test_student_can_confirm_and_cancel_rights_anytime_before_deadline(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        m = make_member('68DS05', year=2568, level='ม.3')
        rr = rights.open_round(2569, 1, admin_u)

        mc = Client()
        mc.force_login(m)

        # มีสิทธิ์ในปีการศึกษานี้ เห็นปุ่มยืนยันใช้สิทธิ์
        html1 = mc.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('ยืนยันใช้สิทธิ์', html1)

        # กดยืนยันใช้สิทธิ์
        mc.post(reverse('composn:confirm_rights'))
        self.assertEqual(m.rights_used_count(), 1)
        html2 = mc.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('ยกเลิกการใช้สิทธิ์', html2)

        # กดยกเลิกการใช้สิทธิ์ก่อนหมดเวลา
        mc.post(reverse('composn:cancel_rights'))
        self.assertEqual(m.rights_used_count(), 0)
        self.assertIsNone(rights.year_record(m, 2569))

        # สามารถกลับมากดยืนยันใช้สิทธิ์ใหม่ได้อีกก่อนหมดเวลา
        mc.post(reverse('composn:confirm_rights'))
        self.assertEqual(m.rights_used_count(), 1)

        # เมื่อปิดรอบแล้ว (หมดเวลา) ไม่สามารถยกเลิกได้
        rights.close_round(rr, admin_u, mark_waived=False)
        with self.assertRaises(rights.RightsError):
            rights.cancel_rights(m)

    def test_admin_import_camp1_rights_with_preview_validation(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()

        # 1. นักเรียนที่มีสิทธิ์ค่าย 1 ปี 2569
        m_valid = make_member('68DS11', year=2568, level='ม.2')
        # 2. นักเรียนที่มีรหัสค่าย 2 แล้ว -> ใช้สิทธิ์ค่าย 1 ไม่ได้
        m_camp2 = make_member('68DS12', year=2568, level='ม.2', camp2='68KM12')
        # 3. นักเรียนที่ระดับชั้นเกิน ม.6 ในปี 2569 -> หมดสิทธิ์
        m_expired = make_member('68DS13', year=2568, level='ม.6')

        ac = Client()
        ac.force_login(admin_u)

        # ทดสอบ Preview ตรวจสอบความถูกต้อง
        resp = ac.post(reverse('composn:admin_import_rights'), {
            'action': 'preview',
            'raw_codes': '68DS11\n68DS12\n68DS13\n99DS99',
        })
        self.assertEqual(resp.status_code, 200)
        preview = resp.context['preview']
        self.assertEqual(preview['total'], 4)
        self.assertEqual(preview['valid_count'], 1)
        self.assertEqual(preview['invalid_count'], 3)

        by_code = {r['camp1_code']: r for r in preview['rows']}
        self.assertTrue(by_code['68DS11']['valid'])
        self.assertFalse(by_code['68DS12']['valid'])
        self.assertIn('ค่าย 2', by_code['68DS12']['reason'])
        self.assertFalse(by_code['68DS13']['valid'])
        self.assertIn('เกิน', by_code['68DS13']['reason'])
        self.assertFalse(by_code['99DS99']['valid'])

        # ยืนยันนำเข้าเฉพาะรหัสที่ผ่านการตรวจสอบ
        resp2 = ac.post(reverse('composn:admin_import_rights'), {
            'action': 'confirm_import',
            'valid_codes': ['68DS11'],
        })
        self.assertEqual(resp2.status_code, 200)
        self.assertEqual(m_valid.rights_used_count(), 1)
        self.assertEqual(m_camp2.rights_used_count(), 0)
        self.assertEqual(m_expired.rights_used_count(), 0)

    def test_menu_order_and_labels(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)
        html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        for numbered in ('1. แดชบอร์ด', '1. ผู้ใช้', '2. เวลาเรียน / คะแนน', '3. อาหาร', '4. การใช้สิทธิ์', '5. ตั้งค่าระบบ'):
            self.assertNotIn(numbered, html)
        for lbl in ('แดชบอร์ด', 'ผู้ใช้', 'เวลาเรียน / คะแนน', 'อาหาร', 'การใช้สิทธิ์', 'ตั้งค่าระบบ'):
            self.assertIn(lbl, html)
        self.assertIn('calc(0.78rem + 2px)', html)
        self.assertIn('calc(0.84rem + 2px)', html)
        self.assertIn('calc(0.86rem + 2px)', html)
        self.assertNotIn('กรอกคะแนน (รายคน', html)
        self.assertNotIn('กรอกคะแนน (รายบุคคล', html)
        self.assertNotIn('ดูหน้าแสดงผล (/display)', html)

    def test_year_grouped_scores_approval_and_camp_colors(self):
        from decimal import Decimal
        from food.models import ScoreCategory, approve_year_scores, member_year_results
        from food.services import save_score

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()

        # นักเรียนรุ่นปี 2568 (ปีที่เข้าอบรม = 2568) ขอใช้สิทธิ์ในปี 2569 (ปีที่ใช้สิทธิ์ = 2569)
        m1 = make_member('68DS21', year=2568, level='ม.3')
        m2 = make_member('68DS22', year=2568, level='ม.3', camp2='68KM22')

        cat68 = ScoreCategory.objects.create(name='สอบปี 68', weight=Decimal('100'), academic_year=2568, is_visible=True)
        cat69 = ScoreCategory.objects.create(name='สอบใช้สิทธิ์ปี 69', weight=Decimal('100'), academic_year=2569, is_visible=True)
        save_score(m1, cat68, 1, Decimal('75'), Decimal('100'), recorded_by=admin_u)
        save_score(m1, cat69, 1, Decimal('85'), Decimal('100'), recorded_by=admin_u)

        rr1 = rights.open_round(2569, 1, admin_u)
        rr2 = rights.open_round(2569, 2, admin_u)

        mc1 = Client()
        mc1.force_login(m1)
        prof1 = mc1.get(reverse('composn:profile')).content.decode('utf-8')
        # ไม่มีจำนวนการใช้สิทธิ์ 0/5 หรือ สิทธิ์คงเหลือ
        self.assertNotIn('ใช้สิทธิ์แล้ว (ครั้ง)', prof1)
        self.assertNotIn('สิทธิ์คงเหลือ', prof1)
        # แจ้งมีสิทธิ์ค่ายที่ 1 + วันเวลาเปิด-ปิด
        self.assertIn('มีสิทธิ์ใช้สิทธิ์ค่ายที่ 1', prof1)
        self.assertIn('เปิดให้ใช้สิทธิ์:', prof1)
        self.assertIn('ปิดระบบใช้สิทธิ์:', prof1)

        # นักเรียนค่าย 2 ใช้พื้นหลังสีส้มเข้ม (#d84315)
        mc2 = Client()
        mc2.force_login(m2)
        prof2 = mc2.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('มีสิทธิ์ใช้สิทธิ์ค่ายที่ 2', prof2)
        self.assertIn('#d84315', prof2)

        # m1 กดใช้สิทธิ์ปี 2569 แต่ยังไม่ปิดรอบและ Admin ยังไม่อนุมัติ -> คะแนนปี 2569 ยังไม่แสดง
        rights.confirm_rights(m1, rr1)
        yr_list = member_year_results(m1)
        self.assertEqual([y['year'] for y in yr_list], [2569, 2568])  # ปีล่าสุดอยู่บนสุด
        self.assertFalse(yr_list[0]['is_training_year'])              # 2569 = ปีที่ใช้สิทธิ์ (ไม่มีเวลาเรียน)
        self.assertIsNone(yr_list[0]['absence'])
        self.assertFalse(yr_list[0]['show_scores'])
        self.assertTrue(yr_list[1]['is_training_year'])               # 2568 = ปีที่เข้าอบรม (มีเวลาเรียน)
        self.assertIsNotNone(yr_list[1]['absence'])
        self.assertTrue(yr_list[1]['show_scores'])

        # ปิดระบบใช้สิทธิ์แล้ว แต่ Admin ยังไม่อนุมัติ -> คะแนนปี 2569 ยังไม่แสดง
        rights.close_round(rr1, admin_u, mark_waived=False)
        yr_list = member_year_results(m1)
        self.assertFalse(yr_list[0]['show_scores'])

        # Admin อนุมัติประกาศผลปี 2569 -> แสดงคะแนนปี 2569 + ผลการคัดเลือกเข้าค่าย 2
        approve_year_scores(2569, admin_u)
        m1.refresh_from_db()
        self.assertEqual(m1.camp2_code, '69KM01')
        yr_list = member_year_results(m1)
        self.assertTrue(yr_list[0]['show_scores'])
        self.assertTrue(yr_list[0]['show_result'])
        self.assertEqual(yr_list[0]['total'], Decimal('85.00'))
        self.assertTrue(yr_list[0]['result'].passed)

        # ยกเลิกอนุมัติ -> รหัสค่าย 2 ที่สร้างของปีนี้ถูกเคลียร์ออก
        from food.models import unapprove_year_scores
        unapprove_year_scores(2569)
        m1.refresh_from_db()
        self.assertIsNone(m1.camp2_code)

    def test_profile_extra_fields_dashboard_summary_display_and_camp_surveys(self):
        from food.models import CampSurvey

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()

        m = make_member('69DS31', year=2569, level='ม.2')
        self.assertFalse(m.is_muslim)  # ค่าเริ่มต้นคือไม่เลือก

        # Admin เปิดให้แก้ไขฟิลด์ทั้งหมด
        ac = Client()
        ac.force_login(admin_u)
        ac.post(reverse('composn:admin_users_settings'), {
            'editable_fields': [
                'student_phone', 'parent_phone', 'email',
                'chronic_disease', 'food_allergy', 'is_muslim',
            ],
            'save_dashboard_fields': '1',
            'dashboard_summary_fields': ['is_muslim', 'food_allergy', 'chronic_disease'],
        })

        # นักเรียนบันทึกข้อมูลส่วนตัว
        mc = Client()
        mc.force_login(m)
        mc.post(reverse('composn:profile'), {
            'action': 'profile',
            'student_phone': '0811112222',
            'parent_phone': '0899998888',
            'email': 'student@example.com',
            'chronic_disease': 'ภูมิแพ้อากาศ',
            'food_allergy': 'แพ้กุ้งและถั่วลิสง',
            'is_muslim': 'on',
        })
        m.refresh_from_db()
        self.assertEqual(m.student_phone, '0811112222')
        self.assertEqual(m.parent_phone, '0899998888')
        self.assertEqual(m.email, 'student@example.com')
        self.assertEqual(m.chronic_disease, 'ภูมิแพ้อากาศ')
        self.assertEqual(m.food_allergy, 'แพ้กุ้งและถั่วลิสง')
        self.assertTrue(m.is_muslim)

        # หน้าสรุปการใช้สิทธิ์ แสดงข้อความ "สรุปการใช้สิทธิ์ สอวน.คอมพิวเตอร์ ศูนย์อบรมโรงเรียนเทพศิรินทร์" และไม่มีแยกศูนย์ DS/KM หรือจำนวนค่าย 1/2
        disp_html = mc.get(reverse('composn:display')).content.decode('utf-8')
        self.assertIn('สรุปการใช้สิทธิ์ สอวน.คอมพิวเตอร์ ศูนย์อบรมโรงเรียนเทพศิรินทร์', disp_html)
        self.assertNotIn('ศูนย์ DS', disp_html)
        self.assertNotIn('สมาชิกทั้งหมด', disp_html)

        # Admin สร้างแบบสอบถามสำรวจการสั่งอาหารในค่าย
        ac.post(reverse('food:surveys_manage'), {'action': 'create_default_food_survey'})
        survey = CampSurvey.objects.filter(academic_year=2569).first()
        self.assertIsNotNone(survey)
        q1 = survey.questions.order_by('sort_order').first()

        # นักเรียนทำแบบสอบถาม
        mc.post(reverse('food:member_surveys'), {
            'survey_id': str(survey.pk),
            f'q_{q1.pk}': 'อาหารฮาลาล (อิสลาม)',
        })

        # ตรวจสอบบน Dashboard ของ Admin ว่าเห็นสรุปศาสนาอิสลาม แพ้อาหาร และผลแบบสอบถาม
        dash_html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        self.assertIn('ศาสนาอิสลาม (อาหารฮาลาล)', dash_html)
        self.assertIn('แพ้กุ้งและถั่วลิสง', dash_html)
        self.assertIn('นักเรียนตอบแล้ว 1 คน', dash_html)

    def test_all_save_buttons_across_system(self):
        from food.models import AbsenceReason, ScoreCategory, TrainingDate, WarningColor

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # 1. ปุ่มบันทึกการตั้งค่าระบบทั้งหมด (/admin/settings/)
        r_obj = AbsenceReason.objects.first()
        c_obj = WarningColor.objects.first()
        post_settings = {
            'action': 'save_all_settings',
            'site_title': 'ระบบ DS-COMPOSN ทดสอบ',
            'display_title': 'สรุปการใช้สิทธิ์ สอวน.คอมพิวเตอร์ ศูนย์อบรมโรงเรียนเทพศิรินทร์',
            'current_academic_year': '2569',
            'max_rights': '5',
            'max_absences': '5',
            'color_monday': '#f9a825',
            'member_theme_primary': '#1b5e20',
            'login_subtitle': 'กรุณาเข้าสู่ระบบ',
        }
        if r_obj:
            post_settings[f'reason_name_{r_obj.pk}'] = r_obj.name
            post_settings[f'reason_order_{r_obj.pk}'] = str(r_obj.order)
            post_settings[f'reason_active_{r_obj.pk}'] = '1'
        if c_obj:
            post_settings[f'color_name_{c_obj.pk}'] = c_obj.name
            post_settings[f'color_code_{c_obj.pk}'] = '#ffc107'
            post_settings[f'color_min_{c_obj.pk}'] = str(c_obj.min_absences)

        resp1 = ac.post(reverse('composn:admin_settings'), post_settings)
        self.assertEqual(resp1.status_code, 302)
        self.assertEqual(SystemSetting.get().site_title, 'ระบบ DS-COMPOSN ทดสอบ')

        # 2. ปุ่มบันทึกการตั้งค่าหน้าอาหาร (/admin/food/settings/)
        resp2 = ac.post(reverse('food:food_settings'), {
            'action': 'save_all_food_settings',
            'food_system_open': '1',
            'food_closed_message': 'ปิดระบบอาหารชั่วคราว',
            'color_monday': '#f9a825',
        })
        self.assertEqual(resp2.status_code, 302)
        self.assertTrue(SystemSetting.get().food_system_open)

        # 3. ปุ่มบันทึกการตั้งค่าหน้าการใช้สิทธิ์ (/admin/rights/settings/)
        resp3 = ac.post(reverse('composn:admin_rights_settings'), {
            'max_rights': '5',
            'display_title': 'สรุปการใช้สิทธิ์ สอวน.คอมพิวเตอร์ ศูนย์อบรมโรงเรียนเทพศิรินทร์',
            'rights_warning_text': 'กรุณาตรวจสอบวันสอบ',
            'exam_date_camp1': '10 ต.ค. 2569',
            'exam_date_camp2': '20 ต.ค. 2569',
        })
        self.assertEqual(resp3.status_code, 302)

        # 4. ปุ่มบันทึกการตั้งค่าคะแนน + วันอบรม + เกณฑ์คัดเลือก (/admin/score-categories/)
        mc = ScoreCategory.objects.create(name='คะแนนเก็บ', weight=40, academic_year=2569, default_max_score=20)
        sc = ScoreCategory.objects.create(parent=mc, name='งาน 1', weight=20, academic_year=2569, default_max_score=20)
        resp4 = ac.post(reverse('food:score_categories'), {
            'action': 'save_all_categories',
            'tab': 'scores',
            f'group_type_{mc.pk}': ScoreCategory.GROUP_COLLECT,
            f'name_{mc.pk}': 'คะแนนเก็บแก้ไข',
            f'weight_{mc.pk}': '45',
            f'default_max_score_{mc.pk}': '25',
            f'order_{mc.pk}': '1',
            f'is_visible_{mc.pk}': '1',
            f'is_open_{mc.pk}': '1',
            f'name_{sc.pk}': 'งาน 1 แก้ไข',
            f'default_max_score_{sc.pk}': '25',
            f'order_{sc.pk}': '1',
            f'is_visible_{sc.pk}': '1',
            f'is_open_{sc.pk}': '1',
        })
        self.assertEqual(resp4.status_code, 302)
        mc.refresh_from_db()
        sc.refresh_from_db()
        self.assertEqual(mc.name, 'คะแนนเก็บแก้ไข')
        self.assertEqual(sc.name, 'งาน 1 แก้ไข')

        td = TrainingDate.objects.create(academic_year=2569, date=date(2026, 10, 15), event_type='training', title='วันแรก')
        resp5 = ac.post(reverse('food:score_categories'), {
            'action': 'save_all_training_dates',
            'tab': 'dates',
            f'td_type_{td.pk}': 'exam',
            f'td_title_{td.pk}': 'สอบท้ายค่าย',
        })
        self.assertEqual(resp5.status_code, 302)
        td.refresh_from_db()
        self.assertEqual(td.event_type, 'exam')
        self.assertEqual(td.title, 'สอบท้ายค่าย')

        resp6 = ac.post(reverse('food:score_categories'), {
            'action': 'save_qualification',
            'tab': 'qualification',
            'passing_score': '60',
            'max_qualified_count': '30',
            'announcement_title': 'ประกาศผลปี 2569',
            'announcement_note': 'หมายเหตุทดสอบ',
        })
        self.assertEqual(resp6.status_code, 302)
        self.assertEqual(int(SystemSetting.get().passing_score), 60)

    def test_admin_users_table_triangles_circles_dropdown_and_current_year_level(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # นักเรียนรุ่นปี 2568 (ม.3 -> ปี 2569 เป็น ม.4) ใช้สิทธิ์ในปี 2569
        m = make_member('68DS35', year=2568, level='ม.3')
        m.nickname = 'ต้นกล้า'
        m.student_phone = '0812345678'
        m.parent_phone = '0898765432'
        m.email = 'tonkla@example.com'
        m.chronic_disease = 'หอบหืด'
        m.food_allergy = 'แพ้อาหารทะเล'
        m.is_first_login = True
        m.save()

        rights.open_round(2569, 1, admin_u)
        rights.confirm_rights(m)

        html = ac.get(reverse('composn:admin_users') + '?year=2569&tab=rights').content.decode('utf-8')
        # แสดงชื่อเล่น ข้อมูลติดต่อ โรคประจำตัว (สามเหลี่ยมสีชมพู #e91e63) และแพ้อาหาร (สามเหลี่ยมสีม่วงอ่อน #b388ff)
        self.assertIn('ต้นกล้า', html)
        self.assertIn('0812345678', html)
        self.assertIn('0898765432', html)
        self.assertIn('tonkla@example.com', html)
        self.assertIn('หอบหืด', html)
        self.assertIn('แพ้อาหารทะเล', html)
        self.assertIn('#e91e63', html)
        self.assertIn('tri-pink', html)
        self.assertIn('#b388ff', html)
        # ตรวจสอบการเลือกปีการศึกษาที่มีข้อมูลในฟอร์มค้นหา
        self.assertIn('<select class="form-select form-select-sm" name="year"', html)
        # ระดับชั้นของนักเรียนที่ใช้สิทธิ์ในปี 2569 ต้องแสดงเป็น ม.4 (ตรงกับปีการศึกษา 2569)
        self.assertIn('<td class="text-center fw-semibold">ม.4</td>', html)

    def test_admin_create_minimal_fields_only(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # ตรวจสอบหน้าฟอร์มเพิ่ม Admin/Staff ว่ามีเฉพาะ รหัสผู้ใช้ บทบาท และรหัสผ่าน
        resp = ac.get(reverse('composn:admin_admin_create'))
        self.assertEqual(resp.status_code, 200)
        self.assertIn('name="camp1_code"', resp.content.decode('utf-8'))
        self.assertIn('name="role"', resp.content.decode('utf-8'))
        self.assertIn('name="password"', resp.content.decode('utf-8'))
        # ต้องไม่มีฟิลด์นักเรียน เช่น school_name, current_level, birth_date
        self.assertNotIn('name="school_name"', resp.content.decode('utf-8'))
        self.assertNotIn('name="current_level"', resp.content.decode('utf-8'))

        # บันทึกสร้างเจ้าหน้าที่ใหม่
        resp2 = ac.post(reverse('composn:admin_admin_create'), {
            'camp1_code': 'staff_test1',
            'role': User.ROLE_STAFF,
            'password': 'password123',
        })
        self.assertEqual(resp2.status_code, 302)
        new_staff = User.objects.get(camp1_code='staff_test1')
        self.assertEqual(new_staff.role, User.ROLE_STAFF)
        self.assertTrue(new_staff.check_password('password123'))

    def test_users_filter_camp2_and_advanced_filters(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        m1 = make_member('69DS81', year=2569, level='ม.4')
        m2 = make_member('69DS82', year=2569, level='ม.4', camp2='69KM82')
        m1.chronic_disease = 'ภูมิแพ้'
        m1.save()

        # กรองเฉพาะคนที่มีรหัสค่าย 2
        html_camp2 = ac.get(reverse('composn:admin_users') + '?year=2569&camp=2').content.decode('utf-8')
        self.assertIn('69DS82', html_camp2)
        self.assertNotIn('69DS81', html_camp2)

        # กรองเฉพาะค่าย 1 (ไม่มีรหัสค่าย 2)
        html_camp1 = ac.get(reverse('composn:admin_users') + '?year=2569&camp=1').content.decode('utf-8')
        self.assertIn('69DS81', html_camp1)
        self.assertNotIn('69DS82', html_camp1)

        # กรองตามข้อมูลโรคประจำตัว
        html_disease = ac.get(reverse('composn:admin_users') + '?year=2569&health=disease').content.decode('utf-8')
        self.assertIn('69DS81', html_disease)
        self.assertNotIn('69DS82', html_disease)

    def test_delete_camp2_preview_and_confirm_with_password(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.set_password('adminpass123')
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        u1 = make_member('69DS88', year=2569, level='ม.4', camp2='69KM88')
        u2 = make_member('69DS89', year=2569, level='ม.4', camp2='69KM89')

        # 1. Preview
        resp_prev = ac.post(reverse('composn:admin_delete_camp2'), {
            'action': 'preview',
            'camp2_codes_text': '69KM88\n69KM999',  # 69KM88 มีในระบบ, 69KM999 ไม่มี
        })
        self.assertEqual(resp_prev.status_code, 200)
        html = resp_prev.content.decode('utf-8')
        self.assertIn('69KM88', html)
        self.assertIn('69KM999', html)
        self.assertIn('1 คนที่ตรงกัน', html)

        # 2. Confirm แต่ใส่รหัสผ่าน Admin ผิด -> ไม่ลบ
        resp_fail = ac.post(reverse('composn:admin_delete_camp2'), {
            'action': 'confirm_delete',
            'confirmed_codes': '69KM88',
            'admin_password': 'wrongpassword',
        })
        u1.refresh_from_db()
        self.assertEqual(u1.camp2_code, '69KM88')  # ยังอยู่

        # 3. Confirm ด้วยรหัสผ่าน Admin ที่ถูกต้อง -> ลบรหัสค่าย 2
        resp_ok = ac.post(reverse('composn:admin_delete_camp2'), {
            'action': 'confirm_delete',
            'confirmed_codes': '69KM88',
            'admin_password': 'adminpass123',
        })
        self.assertEqual(resp_ok.status_code, 302)
        u1.refresh_from_db()
        self.assertIsNone(u1.camp2_code)  # ลบสำเร็จแล้ว

    def test_auto_exam_date_camp1_theory_and_practice(self):
        from food.models import TrainingDate

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # สร้างวันสอบทฤษฎีและปฏิบัติในปี 2569
        TrainingDate.objects.create(
            academic_year=2569, date=date(2026, 10, 15),
            event_type=TrainingDate.TYPE_EXAM, title='สอบทฤษฎีท้ายค่าย 1',
        )
        TrainingDate.objects.create(
            academic_year=2569, date=date(2026, 10, 16),
            event_type=TrainingDate.TYPE_EXAM, title='สอบปฏิบัติต้ายค่าย 1',
        )

        setting = SystemSetting.get()
        setting.current_academic_year = '2569'
        setting.exam_date_camp1 = ''
        setting.save()

        auto_text = setting.get_auto_exam_date_camp1('2569')
        self.assertIn('สอบทฤษฎี: 15 ต.ค. 2569', auto_text)
        self.assertIn('สอบปฏิบัติ: 16 ต.ค. 2569', auto_text)

        # เปิดหน้าตั้งค่าการใช้สิทธิ์ (/admin/rounds/) จะนำข้อความมาลงเป็นค่าเริ่มต้นอัตโนมัติ
        resp = ac.get(reverse('composn:admin_rounds'))
        self.assertEqual(resp.status_code, 200)
        setting.refresh_from_db()
        self.assertIn('สอบทฤษฎี: 15 ต.ค. 2569', setting.exam_date_camp1)
        self.assertIn('สอบปฏิบัติ: 16 ต.ค. 2569', setting.exam_date_camp1)

    def test_users_tabs_current_blue_rights_green_and_ordering(self):
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # นักเรียนปัจจุบันปี 2569
        m_curr = make_member('69DS91', year=2569, level='ม.4')
        # นักเรียนรุ่นปี 2568 ที่มาใช้สิทธิ์ในปี 2569
        m_rights = make_member('68DS92', year=2568, level='ม.4')
        rights.open_round(2569, 1, admin_u)
        rights.confirm_rights(m_rights)

        # 1. แท็บนักเรียนปัจจุบัน (ค่าเริ่มต้น หรือ ?tab=current)
        html_curr = ac.get(reverse('composn:admin_users') + '?year=2569&tab=current').content.decode('utf-8')
        self.assertIn('69DS91', html_curr)
        self.assertNotIn('68DS92', html_curr)
        self.assertIn('code-blue', html_curr)
        self.assertIn('นักเรียนปัจจุบัน', html_curr)

        # 2. แท็บนักเรียนใช้สิทธิ์ (?tab=rights)
        html_rights = ac.get(reverse('composn:admin_users') + '?year=2569&tab=rights').content.decode('utf-8')
        self.assertIn('68DS92', html_rights)
        self.assertNotIn('69DS91', html_rights)
        self.assertIn('code-green', html_rights)
        self.assertIn('นักเรียนใช้สิทธิ์', html_rights)

        # 3. แท็บทั้งหมด (?tab=all) -> แสดงทั้งสองคน โดยเรียงลำดับให้นักเรียนปัจจุบันขึ้นก่อน
        html_all = ac.get(reverse('composn:admin_users') + '?year=2569&tab=all').content.decode('utf-8')
        self.assertIn('69DS91', html_all)
        self.assertIn('68DS92', html_all)
        idx_curr = html_all.find('69DS91')
        idx_rights = html_all.find('68DS92')
        self.assertLess(idx_curr, idx_rights)  # นักเรียนปัจจุบันขึ้นก่อนเสมอ

    def test_staff_scores_includes_confirmed_rights_students_with_cohort_colors_and_ordering(self):
        """หน้ากรอกคะแนน (/admin/scores/) รวมนักเรียนใช้สิทธิ์ที่ยืนยันแล้ว และแสดงสีฟ้า/เขียว พร้อมเรียงนักเรียนปัจจุบันขึ้นก่อน"""
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # สร้างหัวข้อคะแนนใหญ่และย่อยปี 2569
        from food.models import ScoreCategory
        cat_main = ScoreCategory.objects.create(name='ปฏิบัติการ', academic_year=2569, weight=50, default_max_score=100)
        cat_sub = ScoreCategory.objects.create(name='Lab 1', parent=cat_main, academic_year=2569, weight=50, default_max_score=100)

        # นักเรียนปัจจุบันปี 2569
        m_curr = make_member('69DS01', year=2569, level='ม.4')
        # นักเรียนใช้สิทธิ์จากปี 2568
        m_rights = make_member('68DS02', year=2568, level='ม.4')

        # เปิดรอบและยืนยันใช้สิทธิ์ปี 2569
        rights.open_round(2569, 1, admin_u)
        rights.confirm_rights(m_rights)

        res = ac.get(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}')
        self.assertEqual(res.status_code, 200)
        html = res.content.decode('utf-8')

        # ทั้ง 69DS01 และ 68DS02 ต้องปรากฏในหน้ากรอกคะแนน
        self.assertIn('69DS01', html)
        self.assertIn('68DS02', html)

        # ตรวจสอบสีรหัสค่าย
        # นักเรียนปัจจุบัน: code-blue
        self.assertIn('code-blue', html)
        self.assertIn('69DS01', html)

        # นักเรียนใช้สิทธิ์: code-green
        self.assertIn('code-green', html)
        self.assertIn('68DS02', html)

        # ลำดับ: นักเรียนปัจจุบัน (69DS01) ต้องแสดงก่อนนักเรียนใช้สิทธิ์ (68DS02)
        idx_curr = html.find('69DS01')
        idx_rights = html.find('68DS02')
        self.assertLess(idx_curr, idx_rights)

    def test_paste_scores_vertical_and_horizontal(self):
        """ทดสอบการคัดลอกจาก Excel มาวางคะแนนทั้งแบบแนวตั้งและแนวนอน"""
        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        from decimal import Decimal
        from food.models import Score, ScoreCategory
        cat_main = ScoreCategory.objects.create(name='ทฤษฎี', academic_year=2569, weight=50, default_max_score=100)
        cat_sub1 = ScoreCategory.objects.create(name='Quiz 1', parent=cat_main, academic_year=2569, weight=25, default_max_score=50)
        cat_sub2 = ScoreCategory.objects.create(name='Quiz 2', parent=cat_main, academic_year=2569, weight=25, default_max_score=50)

        m1 = make_member('69DS71', year=2569, level='ม.4')
        m2 = make_member('69DS72', year=2569, level='ม.4')

        # 1. วางแบบแนวตั้ง (คอลัมน์คะแนน 2 แถว แถวละ 1 ตัวเลข ลง Quiz 1)
        res_v = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'paste_excel',
            'main_id': cat_main.pk,
            'exam_no': 1,
            'category': cat_sub1.pk,
            'paste_orientation': 'vertical',
            'paste_data': '45.5\n48.0',
        })
        self.assertEqual(res_v.status_code, 302)
        s1 = Score.objects.filter(user=m1, category=cat_sub1, exam_no=1).first()
        s2 = Score.objects.filter(user=m2, category=cat_sub1, exam_no=1).first()
        self.assertIsNotNone(s1)
        self.assertEqual(s1.score, Decimal('45.5'))
        self.assertIsNotNone(s2)
        self.assertEqual(s2.score, Decimal('48.0'))

        # 2. วางแบบแนวนอน (แถวเดียว 2 ตัวเลขคั่นด้วย Tab ข้ามหัวข้อย่อย Quiz 1 และ Quiz 2 ของ m1)
        res_h = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'paste_excel',
            'main_id': cat_main.pk,
            'exam_no': 2,
            'category': cat_sub1.pk,
            'paste_orientation': 'horizontal',
            'paste_data': '35.0\t42.5',
        })
        self.assertEqual(res_h.status_code, 302)
        s1_q1 = Score.objects.filter(user=m1, category=cat_sub1, exam_no=2).first()
        s1_q2 = Score.objects.filter(user=m1, category=cat_sub2, exam_no=2).first()
        self.assertIsNotNone(s1_q1)
        self.assertEqual(s1_q1.score, Decimal('35.0'))
        self.assertIsNotNone(s1_q2)
        self.assertEqual(s1_q2.score, Decimal('42.5'))

    def test_score_excel_import_preview_by_code_and_name_with_column_matching(self):
        """นำเข้าคะแนนด้วยไฟล์ Excel ตรวจสอบคอลัมน์ จับคู่ด้วยรหัสหรือชื่อ และแสดง Preview ปัญหาก่อนนำเข้า"""
        import io
        from decimal import Decimal
        import openpyxl
        from django.core.files.uploadedfile import SimpleUploadedFile
        from food.models import Score, ScoreCategory

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        cat_main = ScoreCategory.objects.create(name='ปฏิบัติการวิจัย', academic_year=2569, weight=60, default_max_score=30)
        cat_sub1 = ScoreCategory.objects.create(name='Lab 1', parent=cat_main, academic_year=2569, weight=10, default_max_score=10)
        cat_sub2 = ScoreCategory.objects.create(name='Lab 2', parent=cat_main, academic_year=2569, weight=20, default_max_score=20)

        # นักเรียนในระบบ
        m1 = make_member('69DS61', year=2569, level='ม.4', first_name='สมชาย', last_name='ใจดี')
        m2 = make_member('69DS62', year=2569, level='ม.4', first_name='ด.ช.วิชัย', last_name='รักเรียน')

        # สร้างไฟล์ Excel จำลอง
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['รหัสค่าย 1', 'ชื่อ-นามสกุล', 'Lab 1 (เต็ม 10)', 'Lab 2', 'หมายเหตุ'])
        # แถว 1: รหัส 69DS61 -> จับคู่ด้วยรหัสค่าย, คะแนนถูกต้อง
        ws.append(['69DS61', 'สมชาย ใจดี', '9.5', '18.0', 'ดีมาก'])
        # แถว 2: ไม่มีรหัสค่าย แต่มีชื่อ "วิชัย รักเรียน" (ตัด ด.ช. ออก) -> จับคู่ด้วยชื่อ, แต่คะแนน Lab 2 เกินเต็ม (25 > 20)
        ws.append(['', 'วิชัย รักเรียน', '8.0', '25.0', ''])
        # แถว 3: รหัสและชื่อไม่มีในระบบ -> ไม่พบนักเรียน
        ws.append(['69DS99', 'ไม่มีในระบบ', '5.0', '10.0', ''])

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        uploaded = SimpleUploadedFile('scores_test.xlsx', bio.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # 1. อัปโหลดไฟล์เพื่อทำ Preview
        res_post = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'import_excel',
            'main_id': cat_main.pk,
            'exam_no': 1,
            'excel_file': uploaded,
        })
        self.assertEqual(res_post.status_code, 302)
        self.assertIn('show_preview=1', res_post.url)

        # 2. ตรวจสอบหน้า Preview (GET)
        res_preview = ac.get(res_post.url)
        self.assertEqual(res_preview.status_code, 200)
        html_preview = res_preview.content.decode('utf-8')

        # ตรวจสอบว่าแสดงหัวข้อ Preview และสรุปการตรวจสอบคอลัมน์
        self.assertIn('ผลการตรวจสอบไฟล์คะแนน Excel ก่อนนำเข้า (Preview)', html_preview)
        self.assertIn('Lab 1', html_preview)
        self.assertIn('Lab 2', html_preview)
        self.assertIn('นำเข้าได้: 1 คน', html_preview)
        self.assertIn('พบปัญหา: 2 คน', html_preview)

        # ตรวจสอบการรายงานปัญหา: เกินเต็ม และไม่พบในระบบ
        self.assertIn('เกินเต็ม', html_preview)
        self.assertIn('ไม่พบนักเรียน', html_preview)
        # ตรวจสอบว่ามีปุ่มยืนยันนำเข้า
        self.assertIn('ยืนยันนำเข้าคะแนน', html_preview)

        # 3. กดยืนยันนำเข้าคะแนน (confirm_import_excel)
        res_confirm = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'confirm_import_excel',
            'main_id': cat_main.pk,
            'exam_no': 1,
        })
        self.assertEqual(res_confirm.status_code, 302)

        # ตรวจสอบคะแนนในฐานข้อมูล: เฉพาะ m1 ที่ข้อมูลถูกต้องถูกบันทึก
        sc1 = Score.objects.filter(user=m1, category=cat_sub1, exam_no=1).first()
        sc2 = Score.objects.filter(user=m1, category=cat_sub2, exam_no=1).first()
        self.assertIsNotNone(sc1)
        self.assertEqual(sc1.score, Decimal('9.5'))
        self.assertIsNotNone(sc2)
        self.assertEqual(sc2.score, Decimal('18.0'))

        # m2 ที่มีคะแนนเกินเต็ม ไม่ถูกบันทึก (ข้ามแถวที่มีปัญหา)
        sc_m2 = Score.objects.filter(user=m2, category=cat_sub2, exam_no=1).first()
        self.assertIsNone(sc_m2)

    def test_login_username_label_and_field_errors(self):
        """หน้าเข้าสู่ระบบแสดง label 'ชื่อผู้ใช้' (ไม่ใช่ 'รหัสค่าย 1 / ชื่อผู้ใช้') และแสดงข้อความผิดพลาด"""
        c = Client()
        html = c.get(reverse('composn:login')).content.decode('utf-8')
        self.assertIn('>ชื่อผู้ใช้<', html)
        self.assertNotIn('รหัสค่าย 1 / ชื่อผู้ใช้', html)
        res = c.post(reverse('composn:login'), {'username': 'toolongusername12345', 'password': '123'})
        self.assertEqual(res.status_code, 200)
        self.assertIn('alert-danger', res.content.decode('utf-8'))

    def test_score_excel_import_headerless_topic_label_override_and_2d_paste(self):
        """ทดสอบนำเข้าไฟล์ Excel ที่ไม่มีหัวตาราง (ข้อ 1..ข้อ 12), แสดงหัวข้อคะแนนในการ์ด, จับคู่ชื่อใกล้เคียง/ข้ามแถวว่าง, และวางตาราง 2D"""
        import io
        from decimal import Decimal
        import openpyxl
        from django.core.files.uploadedfile import SimpleUploadedFile
        from food.models import Score, ScoreCategory

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        cat_main = ScoreCategory.objects.create(name='สอบทฤษฎี 1/2569', academic_year=2569, weight=20, default_max_score=60)
        max_scores = [6, 9, 6, 5, 4, 3, 3, 4, 4, 6, 4, 6]
        subs = [
            ScoreCategory.objects.create(
                name=f'ข้อ {i}', parent=cat_main, academic_year=2569,
                weight=Decimal(ms), default_max_score=Decimal(ms), order=i,
            )
            for i, ms in enumerate(max_scores, start=1)
        ]

        m1 = make_member('69DS81', year=2569, level='ม.4', first_name='นายพีรวัศ', last_name='จิรชัยพิพัฒน์')
        m2 = make_member('69DS82', year=2569, level='ม.4', first_name='นายศิวัชย์พสิษฐ์', last_name='เมฆินทรางกูร')
        m3 = make_member('69DS83', year=2569, level='ม.4', first_name='นายฌัฐพงศ์', last_name='โชติเสรีวิทย์')
        make_member('68DS84', year=2568, level='ม.4', first_name='นายวรภาคย์', last_name='ชลานุเคราะห์')  # ปีอื่น ยังไม่ใช้สิทธิ์ปี 2569

        # ตรวจว่าหน้ากรอกคะแนนแสดงชื่อหัวข้อคะแนนต่อจาก "นำเข้าคะแนนด้วยไฟล์ Excel (.xlsx)"
        page_html = ac.get(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}').content.decode('utf-8')
        self.assertIn('นำเข้าคะแนนด้วยไฟล์ Excel (.xlsx)', page_html)
        self.assertIn('คะแนนสอบทฤษฎี 1/2569', page_html)

        # สร้างไฟล์ Excel ไม่มีแถวหัวตาราง
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['นายพีรวัศ จิรชัยพิพัฒน์', 6, 8, 6, 5, 4, 3, 3, 4, 4, 6, 4, 0])
        ws.append(['นาย ศิวัชย์พสิษฐ์ เมฆินทรางกูร', None, None, None, None, None, None, None, None, None, None, None, None])
        ws.append(['นายณัฐพงศ์ โชติเสรีวิทย์', 6, 9, 4, 5, 4, 3, 3, 4, 4, 6, 4, 4])  # สะกด ณ แทน ฌ
        ws.append(['นายวรภาคย์ ชลานุเคราะห์', 6, 0, 0, 5, 4, 0, 3, 0, 0, 6, 4, 0.5])  # ปี 2568 ยังไม่ใช้สิทธิ์

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        uploaded = SimpleUploadedFile('theory1.xlsx', bio.read(),
                                      content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # แม้อยู่ในหน้าที่มีตัวกรองค้นหา q=... ก็ต้องจับคู่นักเรียนทั้งหมดในปีการศึกษาได้
        res_post = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}&q=filter_something', {
            'action': 'import_excel',
            'year': 2569,
            'main_id': cat_main.pk,
            'exam_no': 1,
            'excel_file': uploaded,
        })
        self.assertEqual(res_post.status_code, 302)
        res_prev = ac.get(res_post.url)
        html_prev = res_prev.content.decode('utf-8')
        self.assertIn('นำเข้าได้: 1 คน', html_prev)
        self.assertIn('พบปัญหา: 2 คน', html_prev)
        self.assertIn('ไม่มีคะแนน (ข้าม): 1 คน', html_prev)
        self.assertIn('ยังไม่ได้ยืนยันใช้สิทธิ์ในปีการศึกษา 2569', html_prev)
        self.assertIn('69DS83 นายฌัฐพงศ์ โชติเสรีวิทย์', html_prev)
        self.assertIn('name="override_user_3"', html_prev)

        # ยืนยันนำเข้าพร้อมเลือกจับคู่แถวที่ 3 (m3) เอง
        res_confirm = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'confirm_import_excel',
            'year': 2569,
            'main_id': cat_main.pk,
            'exam_no': 1,
            'override_user_3': str(m3.pk),
        })
        self.assertEqual(res_confirm.status_code, 302)
        self.assertEqual(Score.objects.get(user=m1, category=subs[0], exam_no=1).score, Decimal('6'))
        self.assertEqual(Score.objects.get(user=m1, category=subs[1], exam_no=1).score, Decimal('8'))
        self.assertEqual(Score.objects.get(user=m1, category=subs[11], exam_no=1).score, Decimal('0'))
        self.assertEqual(Score.objects.get(user=m3, category=subs[2], exam_no=1).score, Decimal('4'))
        self.assertEqual(Score.objects.filter(user=m2, exam_no=1).count(), 0)

        # Export Excel ของหัวข้อที่มีเครื่องหมาย "/" ในชื่อต้องไม่เกิด 500
        res_exp = ac.get(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}&export=xlsx')
        self.assertEqual(res_exp.status_code, 200)

        # วางตาราง 2D ในโหมด auto ต้องกระจายข้ามหัวข้อย่อย (ไม่ยัดรวมลงข้อ 1)
        res_2d = ac.post(reverse('food:staff_scores') + f'?year=2569&main_id={cat_main.pk}', {
            'action': 'paste_excel',
            'year': 2569,
            'main_id': cat_main.pk,
            'exam_no': 2,
            'category': subs[0].pk,
            'paste_orientation': 'auto',
            'paste_data': '5\t7\n4\t6',
        })
        self.assertEqual(res_2d.status_code, 302)
        self.assertEqual(Score.objects.get(user=m1, category=subs[0], exam_no=2).score, Decimal('5'))
        self.assertEqual(Score.objects.get(user=m1, category=subs[1], exam_no=2).score, Decimal('7'))
        self.assertEqual(Score.objects.get(user=m2, category=subs[0], exam_no=2).score, Decimal('4'))
        self.assertEqual(Score.objects.get(user=m2, category=subs[1], exam_no=2).score, Decimal('6'))

    def test_remember_page_state_middleware_and_round_close_fixes(self):
        """ระบบจดจำสถานะหน้าล่าสุด (query/tab) เมื่อย้อนกลับมาอยู่ที่เดิม + ปิด/เปิดรอบจากหน้ารอบไม่ 500"""
        from datetime import timedelta
        from django.utils import timezone
        from food.models import ScoreCategory

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        c1 = ScoreCategory.objects.create(name='หัวข้อ A', academic_year=2569, weight=50, default_max_score=50)
        c2 = ScoreCategory.objects.create(name='หัวข้อ B', academic_year=2569, weight=50, default_max_score=50)

        # 1. เปิดหน้ากรอกคะแนนที่หัวข้อ B ครั้งที่ 2
        url_scores = reverse('food:staff_scores')
        res1 = ac.get(f'{url_scores}?year=2569&main_id={c2.pk}&exam_no=2')
        self.assertEqual(res1.status_code, 200)

        # ย้ายไปหน้าอื่น แล้วกลับมาที่หน้ากรอกคะแนนโดยไม่มี query -> พากลับมาที่หัวข้อ B ครั้งที่ 2
        ac.get(reverse('composn:admin_dashboard'))
        res_back = ac.get(url_scores)
        self.assertEqual(res_back.status_code, 302)
        self.assertIn(f'main_id={c2.pk}', res_back.url)
        self.assertIn('exam_no=2', res_back.url)

        # กดล้าง (?reset=1) -> ลืมค่าที่จำไว้
        res_reset = ac.get(f'{url_scores}?reset=1')
        self.assertEqual(res_reset.status_code, 200)
        res_fresh = ac.get(url_scores)
        self.assertEqual(res_fresh.status_code, 200)

        # 2. ซิงก์แท็บฝั่งเบราว์เซอร์ผ่าน X-Sync-Page-State -> กลับมามี ?tab=dates
        url_cats = reverse('food:score_categories')
        res_sync = ac.get(f'{url_cats}?year=2569&tab=dates', HTTP_X_SYNC_PAGE_STATE='1')
        self.assertEqual(res_sync.status_code, 204)
        res_cat_back = ac.get(url_cats)
        self.assertEqual(res_cat_back.status_code, 302)
        self.assertIn('tab=dates', res_cat_back.url)

        # 3. ปิดรอบและเปิดรอบใหม่โดยส่ง next='composn:admin_rounds' ต้องไม่เกิด 500 NoReverseMatch
        # และปิดรอบหลังเลย scheduled_close_at ต้องบันทึกสละสิทธิ์ได้ถูกต้อง
        m_wait = make_member('68DS91', year=2568, level='ม.4')
        rr = rights.open_round(2569, 1, admin_u)
        rr.scheduled_close_at = timezone.now() - timedelta(hours=1)
        rr.save(update_fields=['scheduled_close_at'])

        res_close = ac.post(reverse('composn:admin_close_round', args=[rr.pk]), {
            'mark_waived': '1',
            'next': 'composn:admin_rounds',
        })
        self.assertEqual(res_close.status_code, 302)
        self.assertEqual(res_close.url, reverse('composn:admin_rounds'))
        self.assertTrue(UseRecord.objects.filter(user=m_wait, rights_round=rr, is_used=False).exists())

        res_reopen = ac.post(reverse('composn:admin_reopen_round', args=[rr.pk]), {
            'next': 'composn:admin_rounds',
        })
        self.assertEqual(res_reopen.status_code, 302)
        self.assertEqual(res_reopen.url, reverse('composn:admin_rounds'))

    def test_auto_logout_minutes_qualification_count_food_groups_notifications_and_attendance_rules(self):
        """ทดสอบ: ตั้งเวลาตัดการเชื่อมต่อ (นาที), คัดเลือกตามจำนวนคน, กลุ่มสั่งอาหาร+แจ้งเตือน, และกติกาบันทึกเวลาเรียนใหม่"""
        import time
        from datetime import timedelta
        from decimal import Decimal
        from django.utils import timezone
        from food.models import Attendance, FoodMenu, FoodOrder, ScoreCategory, TrainingDate, absence_status, compute_year_ranking
        from food.services import save_score

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        ac = Client()
        ac.force_login(admin_u)

        # 1. ตั้งค่าเวลาตัดการเชื่อมต่อเป็น 15 นาที
        res_set = ac.post(reverse('composn:admin_settings'), {
            'action': 'save_all_settings',
            'site_title': 'ระบบ DS-COMPOSN',
            'display_title': 'สรุป',
            'current_academic_year': '2569',
            'max_absences': '5',
            'auto_logout_minutes': '15',
        })
        self.assertEqual(res_set.status_code, 302)
        self.assertEqual(SystemSetting.get().auto_logout_minutes, 15)
        dash_html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        self.assertIn('var limit = 900 * 1000', dash_html)
        self.assertIn('บันทึกเวลาเรียน', dash_html)
        self.assertNotIn('ตรวจสอบเวลาเรียน', dash_html)

        # ทดสอบตัดการเชื่อมต่อฝั่งเซิร์ฟเวอร์เมื่อครบเวลา
        sess = ac.session
        sess['_last_activity'] = time.time() - 1000  # เกิน 900 วินาที
        sess.save()
        res_timeout = ac.get(reverse('composn:admin_dashboard'))
        self.assertEqual(res_timeout.status_code, 302)
        self.assertIn('timeout=1', res_timeout.url)
        ac.force_login(admin_u)

        # 2. ผู้ผ่านการคัดเลือก ระบุเป็นจำนวน คน (Top 2 แม้คะแนน < 50)
        m_curr = make_member('69DS91', year=2569, level='ม.4')
        m_curr2 = make_member('69DS92', year=2569, level='ม.4')
        m_curr3 = make_member('69DS93', year=2569, level='ม.4')
        cat = ScoreCategory.objects.create(name='สอบรวม', academic_year=2569, weight=Decimal('100'), default_max_score=Decimal('100'))
        save_score(m_curr, cat, 1, Decimal('45'), Decimal('100'), recorded_by=admin_u)
        save_score(m_curr2, cat, 1, Decimal('40'), Decimal('100'), recorded_by=admin_u)
        save_score(m_curr3, cat, 1, Decimal('30'), Decimal('100'), recorded_by=admin_u)

        ac.post(reverse('food:score_categories'), {
            'action': 'save_qualification',
            'tab': 'qualification',
            'passing_score': '0',
            'max_qualified_count': '2',
            'announcement_title': 'ประกาศผล',
            'announcement_note': '',
        })
        ranked = compute_year_ranking(2569)
        passed_codes = [r['user'].camp1_code for r in ranked if r['passed']]
        self.assertEqual(passed_codes, ['69DS91', '69DS92'])

        # 3. รายการอาหาร: checkbox นักเรียนปัจจุบัน / นักเรียนใช้สิทธิ์ (ค่าเริ่มต้นนักเรียนปัจจุบันอย่างเดียว) + แจ้งเตือน
        m_rights = make_member('68DS95', year=2568, level='ม.4')
        rr = rights.open_round(2569, 1, admin_u)
        rights.confirm_rights(m_rights, rr)

        today = timezone.localdate()
        # สร้างรายการอาหารค่าเริ่มต้น (นักเรียนปัจจุบันอย่างเดียว)
        o_curr = FoodOrder.objects.create(food_date=today, restaurant_name='ร้านเฉพาะปัจจุบัน', max_selection_count=1)
        menu_curr = FoodMenu.objects.create(food_order=o_curr, menu_name='ข้าวผัด')
        self.assertTrue(o_curr.allow_current_students)
        self.assertFalse(o_curr.allow_rights_students)

        # นักเรียนปัจจุบันต้องเห็นรายการและได้รับการแจ้งเตือน
        mc_curr = Client()
        mc_curr.force_login(m_curr)
        prof_curr = mc_curr.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('แจ้งเตือนการเลือกอาหาร', prof_curr)
        self.assertIn('ร้านเฉพาะปัจจุบัน', prof_curr)

        # นักเรียนใช้สิทธิ์ต้องไม่เห็นรายการนี้และไม่ถูกแจ้งเตือน
        mc_rights = Client()
        mc_rights.force_login(m_rights)
        prof_rights = mc_rights.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('ร้านเฉพาะปัจจุบัน', prof_rights)

        # สร้างรายการที่เปิดให้เฉพาะนักเรียนใช้สิทธิ์ -> นักเรียนใช้สิทธิ์ต้องได้รับการแจ้งเตือน
        o_rights = FoodOrder.objects.create(
            food_date=today, restaurant_name='ร้านสำหรับใช้สิทธิ์', max_selection_count=1,
            allow_current_students=False, allow_rights_students=True,
        )
        FoodMenu.objects.create(food_order=o_rights, menu_name='กะเพรา')
        prof_rights2 = mc_rights.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('แจ้งเตือนการเลือกอาหาร', prof_rights2)
        self.assertIn('ร้านสำหรับใช้สิทธิ์', prof_rights2)

        # เมื่อนักเรียนปัจจุบันกดเลือกอาหารแล้ว การแจ้งเตือนต้องหายไป
        mc_curr.post(reverse('food:my_food_detail', args=[o_curr.pk]), {'action': 'select', 'menu_id': menu_curr.pk})
        prof_curr2 = mc_curr.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('ร้านเฉพาะปัจจุบัน', prof_curr2)

        # 4. บันทึกเวลาเรียน: ห้ามกรอกวันอนาคต, กรอกวันนี้/ย้อนหลังได้, สี วันหยุดแดง/วันสอบเขียว/วันอบรมดำ, ขาด/ลาป่วย/ลากิจ นับเป็นขาดอบรม
        yesterday = today - timedelta(days=1)
        tomorrow = today + timedelta(days=2)
        exam_day = today - timedelta(days=3)
        TrainingDate.objects.create(academic_year=2569, date=exam_day, event_type='exam', title='สอบ')
        TrainingDate.objects.create(academic_year=2569, date=yesterday, event_type='training', title='อบรม')
        TrainingDate.objects.create(academic_year=2569, date=today, event_type='training', title='อบรมวันนี้')
        TrainingDate.objects.create(academic_year=2569, date=tomorrow, event_type='holiday', title='หยุด')
        future_train = today + timedelta(days=3)
        TrainingDate.objects.create(academic_year=2569, date=future_train, event_type='training', title='อบรมอนาคต')

        att_page = ac.get(reverse('food:staff_attendance') + f'?year=2569&date={today.isoformat()}&period=morning').content.decode('utf-8')
        # ตรวจสีป้ายวันหยุด(ดำ)/วันสอบ(แดง)/วันอบรม(น้ำเงิน) และไม่มีปุ่มเลือกทั้งหมด/เอาออกทั้งหมด
        self.assertIn('th-day-holiday', att_page)
        self.assertIn('th-day-exam', att_page)
        self.assertIn('th-day-training', att_page)
        self.assertIn('bg-dark border border-light">วันหยุด</span>', att_page)
        self.assertIn('bg-danger border border-light">วันสอบ</span>', att_page)
        self.assertIn('bg-primary border border-light">วันอบรม</span>', att_page)
        self.assertIn('ยังไม่ถึงวัน', att_page)
        self.assertNotIn('ลงเวลาเรียนทั้งหมด', att_page)
        self.assertNotIn('เอาออกทั้งหมด', att_page)

        # ค่าเริ่มต้นต้องเป็น present (เข้าอบรม) สำหรับวันวานและวันนี้ แต่ไม่สร้างสำหรับวันอนาคตหรือวันหยุด
        self.assertTrue(Attendance.objects.filter(user=m_curr, date=yesterday, session_period='morning', status='present').exists())
        self.assertTrue(Attendance.objects.filter(user=m_curr, date=today, session_period='morning', status='present').exists())
        self.assertFalse(Attendance.objects.filter(user=m_curr, date=future_train, session_period='morning').exists())
        self.assertFalse(Attendance.objects.filter(user=m_curr, date=tomorrow, session_period='morning').exists())

        # ห้ามบันทึกวันในอนาคต
        r_fut = ac.post(reverse('food:attendance_save'), {
            'user_id': m_curr.pk, 'date': future_train.isoformat(), 'period': 'morning', 'status': 'absent', 'year': 2569,
        })
        self.assertEqual(r_fut.status_code, 400)

        # บันทึก ขาดอบรม + ลาป่วย + ลากิจ (ย้อนหลังและวันนี้) -> ต้องนับรวมเป็น 3 ครั้งขาดอบรม
        day_minus_2 = today - timedelta(days=2)
        r1 = ac.post(reverse('food:attendance_save'), {
            'user_id': m_curr.pk, 'date': day_minus_2.isoformat(), 'period': 'morning', 'status': 'absent', 'year': 2569,
        })
        self.assertEqual(r1.status_code, 200)
        r2 = ac.post(reverse('food:attendance_save'), {
            'user_id': m_curr.pk, 'date': yesterday.isoformat(), 'period': 'morning', 'status': 'sick', 'year': 2569,
        })
        self.assertEqual(r2.status_code, 200)
        r3 = ac.post(reverse('food:attendance_save'), {
            'user_id': m_curr.pk, 'date': today.isoformat(), 'period': 'morning', 'status': 'personal', 'year': 2569,
        })
        self.assertEqual(r3.status_code, 200)
        self.assertEqual(r3.json()['absence']['count'], 3)
        self.assertEqual(absence_status(m_curr, 2569)['count'], 3)

    def test_batch3_food_cards_scores_posn_logo_signin_sheet_and_two_slot_attendance(self):
        import base64
        import tempfile
        from datetime import date, timedelta
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test import override_settings
        from django.utils import timezone
        from food.models import (Attendance, FoodMenu, FoodOrder, FoodSelection,
                                 ScoreCategory, TrainingDate, absence_status)

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        s = SystemSetting.get()
        s.current_academic_year = 2569
        s.food_system_open = True
        s.save()

        # สร้างนักเรียน 52 คน เพื่อทดสอบทั้งการ์ดอาหารและใบลงชื่อ 50-52 คน/แผ่น
        members = [make_member(f'69DS{i:02d}', year=2569, level='ม.4') for i in range(1, 53)]
        m1, m2, m3 = members[0], members[1], members[2]

        ac = Client()
        ac.force_login(admin_u)
        mc = Client()
        mc.force_login(m1)

        # 1. การ์ดอาหาร: เรียงตามวันที่น้อยไปมาก + แสดง "เลือกแล้ว X คน / 52 คน" + ปุ่มกลับหน้าหลักรายการอาหาร
        today = timezone.localdate()
        o_late = FoodOrder.objects.create(
            academic_year=2569, food_date=today + timedelta(days=5), restaurant_name='ร้านวันหลัง',
            max_selection_count=1, edit_toggle_open=True, allow_current_students=True, allow_rights_students=False,
        )
        o_early = FoodOrder.objects.create(
            academic_year=2569, food_date=today + timedelta(days=1), restaurant_name='ร้านวันแรก',
            max_selection_count=1, edit_toggle_open=True, allow_current_students=True, allow_rights_students=False,
        )
        menu_e = FoodMenu.objects.create(food_order=o_early, menu_name='ข้าวมันไก่')
        FoodMenu.objects.create(food_order=o_late, menu_name='ก๋วยเตี๋ยว')
        for stu in (m1, m2, m3):
            FoodSelection.objects.create(food_order=o_early, food_menu=menu_e, user=stu, quantity=1)

        orders_html = ac.get(reverse('food:orders') + '?year=2569').content.decode('utf-8')
        self.assertLess(orders_html.index('ร้านวันแรก'), orders_html.index('ร้านวันหลัง'))
        self.assertIn('เลือกแล้ว 3 คน / 52 คน', orders_html)
        self.assertIn('เลือกแล้ว 0 คน / 52 คน', orders_html)

        my_food_resp = mc.get(reverse('food:my_food'))
        self.assertEqual([o.restaurant_name for o in my_food_resp.context['orders']], ['ร้านวันแรก', 'ร้านวันหลัง'])
        my_food_html = my_food_resp.content.decode('utf-8')
        cards_section = my_food_html[my_food_html.index('class="card food-card'):]
        self.assertLess(cards_section.index('ร้านวันแรก'), cards_section.index('ร้านวันหลัง'))
        self.assertIn('เลือกแล้ว 3 คน / 52 คน', my_food_html)

        # ตรวจปุ่มกลับหน้ารายการอาหารในหน้าย่อยทุกหน้า
        for url in [
            reverse('food:order_detail', args=[o_early.pk]),
            reverse('food:order_report', args=[o_early.pk]),
            reverse('food:order_create'),
            reverse('food:order_edit', args=[o_early.pk]),
            reverse('food:order_copy', args=[o_early.pk]),
        ]:
            sub_html = ac.get(url).content.decode('utf-8')
            self.assertIn('กลับหน้ารายการอาหาร', sub_html)
            self.assertIn(reverse('food:orders'), sub_html)
        for url in [
            reverse('food:my_food_detail', args=[o_early.pk]),
        ]:
            sub_html = mc.get(url).content.decode('utf-8')
            self.assertIn('กลับหน้ารายการอาหาร', sub_html)
            self.assertIn(reverse('food:my_food'), sub_html)
        self.assertRedirects(mc.get(reverse('food:my_selections')), reverse('food:my_food'))

        # 2. กรอกคะแนน: นำช่อง "ครั้งที่สอบ" ออกจาก UI แล้ว
        ScoreCategory.objects.create(academic_year=2569, name='ทฤษฎี', default_max_score=100)
        scores_html = ac.get(reverse('food:staff_scores') + '?year=2569').content.decode('utf-8')
        self.assertNotIn('ครั้งที่สอบ', scores_html)

        # 3. อัปโหลดตราโรงเรียน + ตรา สอวน. ในหน้าตั้งค่าระบบ และตรวจใบลงชื่อ A4 แนวนอน (52 คน/แผ่น, 2 แผ่น/วัน)
        tiny_png = base64.b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            with override_settings(MEDIA_ROOT=tmpdir):
                r_set = ac.post(reverse('composn:admin_settings'), {
                    'action': 'save_all_settings',
                    'current_academic_year': '2569',
                    'max_absences': '3',
                    'auto_logout_minutes': '60',
                    'logo_school': SimpleUploadedFile('debsirin.png', tiny_png, content_type='image/png'),
                    'logo_posn': SimpleUploadedFile('posn.png', tiny_png, content_type='image/png'),
                })
                self.assertEqual(r_set.status_code, 302)
                s.refresh_from_db()
                self.assertTrue(bool(s.logo_school))
                self.assertTrue(bool(s.logo_posn))

                settings_html = ac.get(reverse('composn:admin_settings')).content.decode('utf-8')
                self.assertIn('ตรา สอวน.', settings_html)
                self.assertIn('name="logo_posn"', settings_html)

                # ใบลงชื่อเข้าอบรม: 52 คน -> ค่าเริ่มต้น (period=both) ต้องได้ 2 แผ่น (ภาคเช้า + ภาคบ่าย) แผ่นละ 26 แถว (ซ้าย-ขวา)
                r_sheet = ac.get(reverse('food:signin_sheet') + f'?year=2569&date={today.isoformat()}')
                self.assertEqual(r_sheet.status_code, 200)
                sheets = r_sheet.context['sheets']
                self.assertEqual(len(sheets), 2)
                self.assertEqual([sh['period_key'] for sh in sheets], ['morning', 'afternoon'])
                self.assertEqual(len(sheets[0]['paired_rows']), 26)
                self.assertEqual(sheets[0]['count'], 52)
                sheet_html = r_sheet.content.decode('utf-8')
                self.assertIn('size: A4 landscape', sheet_html)
                self.assertIn('ภาคเช้า', sheet_html)
                self.assertIn('ภาคบ่าย', sheet_html)
                self.assertIn(s.logo_school.url, sheet_html)
                self.assertIn(s.logo_posn.url, sheet_html)
                self.assertLess(sheet_html.index(s.logo_school.url), sheet_html.index(s.logo_posn.url))
                self.assertIn('ศูนย์ส่งเสริมโอลิมปิกวิชาการและพัฒนามาตรฐานวิทยาศาสตร์ศึกษา', sheet_html)

                # เลือกเฉพาะภาคเช้า -> ได้ 1 แผ่น
                r_sheet_m = ac.get(reverse('food:signin_sheet') + f'?year=2569&date={today.isoformat()}&period=morning')
                self.assertEqual(len(r_sheet_m.context['sheets']), 1)

                # ล้างรูปตรา สอวน. ได้
                ac.post(reverse('composn:admin_settings'), {
                    'action': 'save_all_settings',
                    'current_academic_year': '2569',
                    'max_absences': '3',
                    'auto_logout_minutes': '60',
                    'clear_logo_posn': '1',
                })
                s.refresh_from_db()
                self.assertFalse(bool(s.logo_posn))

        # 4. บันทึกเวลาเรียน 1 วันมี 2 ช่อง (เช้า + บ่าย)
        yesterday = today - timedelta(days=1)
        future_day = today + timedelta(days=3)
        holiday_day = today + timedelta(days=2)
        exam_day = today - timedelta(days=2)
        TrainingDate.objects.create(academic_year=2569, date=exam_day, event_type='exam', title='วันสอบ')
        TrainingDate.objects.create(academic_year=2569, date=yesterday, event_type='training', title='วันอบรม 1')
        TrainingDate.objects.create(academic_year=2569, date=today, event_type='training', title='วันอบรม 2')
        TrainingDate.objects.create(academic_year=2569, date=holiday_day, event_type='holiday', title='วันหยุด')
        TrainingDate.objects.create(academic_year=2569, date=future_day, event_type='training', title='วันอนาคต')

        att_resp = ac.get(reverse('food:staff_attendance') + f'?year=2569&date={today.isoformat()}&period=morning')
        self.assertEqual(att_resp.status_code, 200)
        att_html = att_resp.content.decode('utf-8')
        # ตารางต้องแสดงหัวคอลัมน์ย่อย เช้า และ บ่าย + ปุ่ม matrix ที่มีทั้ง data-period="morning" และ data-period="afternoon"
        self.assertIn('colspan="2"', att_html)
        self.assertIn('data-period="morning"', att_html)
        self.assertIn('data-period="afternoon"', att_html)
        # ค่าเริ่มต้น present ต้องถูกสร้างทั้งภาคเช้าและภาคบ่ายสำหรับวันที่มาถึงแล้ว (และไม่สร้างให้วันอนาคต/วันหยุด)
        self.assertTrue(Attendance.objects.filter(user=m1, date=yesterday, session_period='morning', status='present').exists())
        self.assertTrue(Attendance.objects.filter(user=m1, date=yesterday, session_period='afternoon', status='present').exists())
        self.assertTrue(Attendance.objects.filter(user=m1, date=today, session_period='morning', status='present').exists())
        self.assertTrue(Attendance.objects.filter(user=m1, date=today, session_period='afternoon', status='present').exists())
        self.assertFalse(Attendance.objects.filter(user=m1, date=future_day).exists())
        self.assertFalse(Attendance.objects.filter(user=m1, date=holiday_day).exists())

        # บันทึกภาคเช้าและภาคบ่ายแยกกันได้อิสระ และคงหมายเหตุเดิมเมื่ออัปเดตสถานะจากตาราง Matrix
        r_am = ac.post(reverse('food:attendance_save'), {
            'user_id': m1.pk, 'date': today.isoformat(), 'period': 'morning',
            'status': 'sick', 'reason': 'ป่วย', 'note': 'มีใบรับรองแพทย์', 'year': 2569,
        })
        self.assertEqual(r_am.status_code, 200)
        self.assertEqual(r_am.json()['absence']['count'], 1)
        r_pm = ac.post(reverse('food:attendance_save'), {
            'user_id': m1.pk, 'date': today.isoformat(), 'period': 'afternoon',
            'status': 'absent', 'year': 2569,
        })
        self.assertEqual(r_pm.status_code, 200)
        self.assertEqual(r_pm.json()['absence']['count'], 2)
        # เปลี่ยนภาคเช้าจาก sick -> personal โดยไม่ส่ง note (เหมือนกดปุ่มใน Matrix) ต้องยังคง note เดิมไว้
        r_am2 = ac.post(reverse('food:attendance_save'), {
            'user_id': m1.pk, 'date': today.isoformat(), 'period': 'morning',
            'status': 'personal', 'year': 2569,
        })
        self.assertEqual(r_am2.status_code, 200)
        att_am = Attendance.objects.get(user=m1, date=today, session_period='morning')
        att_pm = Attendance.objects.get(user=m1, date=today, session_period='afternoon')
        self.assertEqual(att_am.status, 'personal')
        self.assertEqual(att_am.note, 'มีใบรับรองแพทย์')
        self.assertEqual(att_pm.status, 'absent')
        self.assertEqual(absence_status(m1, 2569)['count'], 2)

    def test_batch4_day_colors_settings_dedup_face_photo_staff_perms_and_sidebar(self):
        import io
        import tempfile
        from datetime import timedelta
        from unittest.mock import patch
        import cv2
        import numpy as np
        from PIL import Image
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test import override_settings
        from django.utils import timezone
        from composn.face_photo import FacePhotoError, MAX_OUTPUT_BYTES, process_and_validate_profile_photo
        from food.models import FoodOrder, ScoreCategory, TrainingDate

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        s = SystemSetting.get()
        s.current_academic_year = 2569
        s.save()

        m1 = make_member('69DS11', year=2569, level='ม.4', first_name='สมชาย', last_name='ใจดี')
        ac = Client()
        ac.force_login(admin_u)

        # 1. ตรวจสีพื้นวันอบรม(น้ำเงิน) / วันสอบ(แดง) / วันหยุด(ดำ) และนำเหตุผลการขาดเรียนออก
        today = timezone.localdate()
        TrainingDate.objects.create(academic_year=2569, date=today - timedelta(days=1), event_type='training', title='อบรม')
        TrainingDate.objects.create(academic_year=2569, date=today, event_type='exam', title='สอบ')
        TrainingDate.objects.create(academic_year=2569, date=today + timedelta(days=1), event_type='holiday', title='หยุด')

        att_html = ac.get(reverse('food:staff_attendance') + f'?year=2569&date={today.isoformat()}').content.decode('utf-8')
        self.assertIn('th-day-training', att_html)
        self.assertIn('th-day-exam', att_html)
        self.assertIn('th-day-holiday', att_html)
        self.assertIn('td-day-holiday', att_html)
        self.assertNotIn('reason-sel', att_html)
        self.assertNotIn('เหตุผลเพิ่มเติม', att_html)

        cat_html = ac.get(reverse('food:score_categories') + '?year=2569&tab=dates').content.decode('utf-8')
        self.assertIn("'training': 'bg-primary'", cat_html)
        self.assertIn("'exam': 'bg-danger'", cat_html)
        self.assertIn("'holiday': 'bg-dark'", cat_html)

        # 2. ตั้งค่าระบบ: ไม่มีเหตุผลการขาดเรียน, ไม่มีฟิลด์ซ้ำ (color_monday / editable_fields / display_title), และมีสีแถบข้าง Staff
        set_html = ac.get(reverse('composn:admin_settings')).content.decode('utf-8')
        self.assertNotIn('เหตุผลการขาดเรียน', set_html)
        self.assertNotIn('name="color_monday"', set_html)
        self.assertNotIn('name="editable_fields"', set_html)
        self.assertNotIn('name="display_title"', set_html)
        self.assertIn('name="staff_theme_primary"', set_html)
        self.assertIn('name="staff_theme_gradient"', set_html)

        # บันทึกสีแถบข้าง Staff
        r_save_set = ac.post(reverse('composn:admin_settings'), {
            'action': 'save_all_settings',
            'site_title': 'ระบบ DS-COMPOSN',
            'current_academic_year': '2569',
            'max_absences': '5',
            'auto_logout_minutes': '15',
            'staff_theme_primary': '#123456',
            'staff_theme_gradient': '#654321',
        })
        self.assertEqual(r_save_set.status_code, 302)
        s.refresh_from_db()
        self.assertEqual(s.staff_theme_primary, '#123456')
        self.assertEqual(s.staff_theme_gradient, '#654321')

        # 3. ระบบตรวจใบหน้า ครอปเฉพาะใบหน้า <= 1MB ตรวจหน้ากาก แว่นดำ หมวก + อัปโหลดและ Lightbox
        def _make_portrait_bytes(mask=False, sunglasses=False, hat_color=None, blank=False):
            img = np.full((600, 600, 3), (220, 235, 250), dtype=np.uint8)
            if not blank:
                cv2.ellipse(img, (300, 230), (115, 115), 0, 180, 360, (35, 30, 28), -1)
                cv2.ellipse(img, (300, 290), (105, 130), 0, 0, 360, (165, 195, 235), -1)
                for ex in (258, 342):
                    cv2.ellipse(img, (ex, 265), (18, 9), 0, 0, 360, (245, 245, 245), -1)
                    cv2.circle(img, (ex, 265), 7, (45, 35, 30), -1)
                    cv2.ellipse(img, (ex, 244), (22, 5), 0, 180, 360, (40, 35, 30), 3)
                cv2.line(img, (300, 275), (300, 320), (140, 165, 205), 3)
                cv2.ellipse(img, (300, 355), (26, 9), 0, 0, 180, (125, 125, 195), -1)
                if sunglasses:
                    cv2.rectangle(img, (218, 240), (290, 290), (18, 18, 18), -1)
                    cv2.rectangle(img, (310, 240), (382, 290), (18, 18, 18), -1)
                if mask:
                    cv2.rectangle(img, (215, 305), (385, 405), (240, 210, 150), -1)
                if hat_color is not None:
                    cv2.rectangle(img, (185, 110), (415, 215), hat_color, -1)
            rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            buf = io.BytesIO()
            Image.fromarray(rgb).save(buf, format='JPEG', quality=92)
            return buf.getvalue()

        # ภาพไม่มีใบหน้า -> ต้องถูกปฏิเสธ
        with self.assertRaises(FacePhotoError):
            process_and_validate_profile_photo(io.BytesIO(_make_portrait_bytes(blank=True)), '69DS11')

        # ตรวจหน้ากาก / แว่นดำ / หมวก และครอปภาพปกติ <= 1MB
        with patch('composn.face_photo._detect_primary_face', return_value=(195, 180, 210, 220)):
            with self.assertRaises(FacePhotoError) as ctx_sg:
                process_and_validate_profile_photo(io.BytesIO(_make_portrait_bytes(sunglasses=True)), '69DS11')
            self.assertIn('แว่น', str(ctx_sg.exception))

            with self.assertRaises(FacePhotoError) as ctx_mk:
                process_and_validate_profile_photo(io.BytesIO(_make_portrait_bytes(mask=True)), '69DS11')
            self.assertIn('หน้ากาก', str(ctx_mk.exception))

            with self.assertRaises(FacePhotoError) as ctx_ht:
                process_and_validate_profile_photo(io.BytesIO(_make_portrait_bytes(hat_color=(35, 35, 215))), '69DS11')
            self.assertIn('หมวก', str(ctx_ht.exception))

            cf, info = process_and_validate_profile_photo(io.BytesIO(_make_portrait_bytes()), '69DS11')
            self.assertLessEqual(info['size_bytes'], MAX_OUTPUT_BYTES)
            self.assertEqual(info['width'], info['height'])
            self.assertLessEqual(info['width'], 600)

        # ทดสอบ Member อัปโหลดรูปจริงผ่านหน้า /profile/ และแสดงผลพร้อม Lightbox ในหน้าผู้ใช้
        mc = Client()
        mc.force_login(m1)
        with tempfile.TemporaryDirectory() as tmpdir:
            with override_settings(MEDIA_ROOT=tmpdir):
                # อัปโหลดภาพเปล่า -> ไม่บันทึก
                mc.post(reverse('composn:profile'), {
                    'action': 'upload_photo',
                    'profile_photo': SimpleUploadedFile('blank.jpg', _make_portrait_bytes(blank=True), content_type='image/jpeg'),
                })
                m1.refresh_from_db()
                self.assertFalse(bool(m1.profile_photo))

                # อัปโหลดภาพใบหน้าผ่านเกณฑ์ -> บันทึกสำเร็จและแสดงในหน้าผู้ใช้พร้อม Lightbox ย่อ/ขยาย
                with patch('composn.face_photo._detect_primary_face', return_value=(195, 180, 210, 220)):
                    mc.post(reverse('composn:profile'), {
                        'action': 'upload_photo',
                        'profile_photo': SimpleUploadedFile('face.jpg', _make_portrait_bytes(), content_type='image/jpeg'),
                    })
                m1.refresh_from_db()
                self.assertTrue(bool(m1.profile_photo))
                self.assertLessEqual(m1.profile_photo.size, MAX_OUTPUT_BYTES)

                users_html = ac.get(reverse('composn:admin_users') + '?year=2569').content.decode('utf-8')
                self.assertIn(m1.profile_photo.url, users_html)
                self.assertIn('openPhotoLightbox', users_html)
                self.assertIn('id="photoZoomModal"', users_html)

                # 4. สิทธิ์รายบุคคลของเจ้าหน้าที่ (Staff) + สีแถบข้าง Staff
                # 4.1 สร้าง Staff ที่มีเฉพาะสิทธิ์ ดูผู้ใช้ (อ่านอย่างเดียว) + พิมพ์ใบลงชื่อ
                r_create_st = ac.post(reverse('composn:admin_admin_create'), {
                    'camp1_code': 'staff_view',
                    'role': 'staff',
                    'password': 'password123',
                    'staff_can_view_users': 'on',
                    'staff_can_signin_sheet': 'on',
                })
                self.assertEqual(r_create_st.status_code, 302)
                st_view = User.objects.get(camp1_code='staff_view')
                st_view.is_first_login = False
                st_view.save()
                self.assertTrue(st_view.staff_can_view_users)
                self.assertTrue(st_view.staff_can_signin_sheet)
                self.assertFalse(st_view.staff_can_attendance)
                self.assertFalse(st_view.staff_can_manage_food)
                self.assertFalse(st_view.staff_can_scores)

                # หน้าจัดการแอดมินต้องแสดงป้ายสิทธิ์ของ Staff
                admins_html = ac.get(reverse('composn:admin_admins')).content.decode('utf-8')
                self.assertIn('ผู้ใช้ (นักเรียน/สมาชิก)', admins_html)
                self.assertIn('แบบลงชื่อ', admins_html)

                sc = Client()
                sc.force_login(st_view)
                # home redirect ต้องพาไปหน้าที่มีสิทธิ์ (signin_sheet) ไม่ติด 403
                self.assertRedirects(sc.get(reverse('composn:home')), reverse('food:signin_sheet'))

                # สีแถบข้าง Staff ต้องแสดงตรงตามที่ตั้งค่า
                sheet_resp = sc.get(reverse('food:signin_sheet'))
                self.assertEqual(sheet_resp.status_code, 200)
                sheet_html = sheet_resp.content.decode('utf-8')
                self.assertIn('#123456', sheet_html)
                self.assertIn('#654321', sheet_html)
                # Sidebar แสดงเฉพาะเมนูที่มีสิทธิ์
                self.assertIn('ผู้ใช้ (นักเรียน/สมาชิก)', sheet_html)
                self.assertIn('แบบลงชื่อ', sheet_html)
                self.assertNotIn('บันทึกคะแนน', sheet_html)
                self.assertNotIn('บันทึกเวลาเรียน', sheet_html)
                self.assertNotIn('รายการอาหาร', sheet_html)

                # ดูผู้ใช้ + ค้นหา + ดูรูป + ดูรายละเอียดได้ แต่แก้ไข/ลบ/Export ไม่ได้
                st_users = sc.get(reverse('composn:admin_users') + '?year=2569&q=69DS11')
                self.assertEqual(st_users.status_code, 200)
                st_users_html = st_users.content.decode('utf-8')
                self.assertIn('69DS11', st_users_html)
                self.assertIn(m1.profile_photo.url, st_users_html)
                self.assertIn('openPhotoLightbox', st_users_html)
                self.assertNotIn(reverse('composn:admin_user_edit', args=[m1.pk]), st_users_html)
                self.assertNotIn(reverse('composn:admin_user_create'), st_users_html)

                st_detail = sc.get(reverse('composn:admin_user_detail', args=[m1.pk]))
                self.assertEqual(st_detail.status_code, 200)
                st_detail_html = st_detail.content.decode('utf-8')
                self.assertIn(m1.profile_photo.url, st_detail_html)
                self.assertNotIn(reverse('composn:admin_user_edit', args=[m1.pk]), st_detail_html)

                # ห้ามเข้าหน้าแก้ไข/ลบ/Export/ดูบัญชี Admin/และเมนูที่ไม่ได้รับสิทธิ์
                self.assertEqual(sc.get(reverse('composn:admin_user_edit', args=[m1.pk])).status_code, 403)
                self.assertEqual(sc.post(reverse('composn:admin_user_delete', args=[m1.pk])).status_code, 403)
                self.assertEqual(sc.get(reverse('composn:admin_users') + '?export=xlsx').status_code, 403)
                self.assertEqual(sc.get(reverse('composn:admin_user_detail', args=[admin_u.pk])).status_code, 403)
                self.assertEqual(sc.get(reverse('food:staff_attendance')).status_code, 403)
                self.assertEqual(sc.get(reverse('food:staff_scores')).status_code, 403)
                self.assertEqual(sc.get(reverse('food:orders')).status_code, 403)

                # 4.2 สร้าง Staff ที่มีเฉพาะสิทธิ์ จัดการอาหาร (เพิ่ม/ลบ/แก้ไข เหมือน Admin)
                st_food = User.objects.create_user(
                    'staff_food', 'password123', role='staff', is_first_login=False,
                    staff_can_view_users=False, staff_can_attendance=False,
                    staff_can_signin_sheet=False, staff_can_manage_food=True, staff_can_scores=False,
                )
                fc = Client()
                fc.force_login(st_food)
                self.assertEqual(fc.get(reverse('food:orders')).status_code, 200)
                r_fc = fc.post(reverse('food:order_create'), {
                    'food_date': today.isoformat(),
                    'restaurant_name': 'ร้าน Staff สร้าง',
                    'max_selection_count': '1',
                    'selection_mode': 'unique',
                    'allow_current_students': 'on',
                    'edit_toggle_open': 'on',
                })
                self.assertEqual(r_fc.status_code, 302)
                fo = FoodOrder.objects.get(restaurant_name='ร้าน Staff สร้าง')
                self.assertEqual(fc.post(reverse('food:order_action', args=[fo.pk]), {'action': 'delete'}).status_code, 302)
                self.assertFalse(FoodOrder.objects.filter(pk=fo.pk).exists())
                # ถูกบล็อกจากหน้าอื่น
                self.assertEqual(fc.get(reverse('composn:admin_users')).status_code, 403)
                self.assertEqual(fc.get(reverse('food:staff_scores')).status_code, 403)
                self.assertEqual(fc.get(reverse('food:staff_attendance')).status_code, 403)
                self.assertEqual(fc.get(reverse('food:signin_sheet')).status_code, 403)

                # 4.3 Staff เฉพาะกรอกคะแนนสอบ
                ScoreCategory.objects.create(academic_year=2569, name='ทฤษฎี', default_max_score=100)
                st_score = User.objects.create_user(
                    'staff_score', 'password123', role='staff', is_first_login=False,
                    staff_can_view_users=False, staff_can_attendance=False,
                    staff_can_signin_sheet=False, staff_can_manage_food=False, staff_can_scores=True,
                )
                scc = Client()
                scc.force_login(st_score)
                self.assertEqual(scc.get(reverse('food:staff_scores')).status_code, 200)
                self.assertEqual(scc.get(reverse('food:orders')).status_code, 403)
                self.assertEqual(scc.get(reverse('food:staff_attendance')).status_code, 403)

    def test_batch5_signin_sheet_pdf_date_range_and_exam_group_split(self):
        from datetime import timedelta
        from django.utils import timezone
        from food.models import TrainingDate

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        s = SystemSetting.get()
        s.current_academic_year = 2569
        s.save()

        # สร้างนักเรียนปัจจุบัน 3 คน (ปี 2569) และนักเรียนใช้สิทธิ์ 2 คน (ปีเก่า 2568 ยืนยันใช้สิทธิ์ปี 2569 ค่าย 1)
        cur1 = make_member('69DS01', year=2569, level='ม.4', first_name='ปัจจุบันหนึ่ง', last_name='ทดสอบ')
        cur2 = make_member('69DS02', year=2569, level='ม.4', first_name='ปัจจุบันสอง', last_name='ทดสอบ')
        cur3 = make_member('69DS03', year=2569, level='ม.4', first_name='ปัจจุบันสาม', last_name='ทดสอบ')
        r_stu1 = make_member('68DS01', year=2568, level='ม.3', first_name='ใช้สิทธิ์หนึ่ง', last_name='ปีเก่า')
        r_stu2 = make_member('68DS02', year=2568, level='ม.3', first_name='ใช้สิทธิ์สอง', last_name='ปีเก่า')
        rr = rights.open_round(2569, 1, admin_u)
        rec1 = rights.confirm_rights(r_stu1, rr)
        rec2 = rights.confirm_rights(r_stu2, rr)

        # ตั้งค่าตาราง 3 วันติดกัน: วันอบรม -> วันหยุด (ต้องถูกข้าม) -> วันสอบ
        d1 = timezone.localdate()
        d2 = d1 + timedelta(days=1)
        d3 = d1 + timedelta(days=2)
        TrainingDate.objects.create(academic_year=2569, date=d1, event_type='training', title='อบรมวันแรก')
        TrainingDate.objects.create(academic_year=2569, date=d2, event_type='holiday', title='วันหยุด')
        TrainingDate.objects.create(academic_year=2569, date=d3, event_type='exam', title='สอบท้ายค่าย')

        ac = Client()
        ac.force_login(admin_u)

        # 1. ทดสอบหน้าเว็บช่วงวันที่ d1..d3 (mode=auto, period=both)
        q_range = f'?year=2569&camp=1&mode=auto&period=both&date_start={d1.isoformat()}&date_end={d3.isoformat()}'
        resp = ac.get(reverse('food:signin_sheet') + q_range)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode('utf-8')

        # หัวข้อหน้าต้องไม่มี "(A4 แนวนอน · 50–52 คน/แผ่น)" และเปลี่ยนปุ่มเป็น Export PDF
        self.assertIn('แบบลงชื่อ', html)
        self.assertNotIn('แบบลงชื่อเข้าอบรม / วันสอบท้ายค่าย', html)
        self.assertNotIn('(A4 แนวนอน · 50–52 คน/แผ่น)', html)
        self.assertIn('Export PDF', html)
        self.assertNotIn('พิมพ์ใบลงชื่อ A4 แนวนอน', html)
        self.assertIn('name="date_start"', html)
        self.assertIn('name="date_end"', html)

        # ตรวจจำนวนแผ่น:
        # - d1 (วันอบรม): 2 แผ่น (เช้า/บ่าย) เฉพาะนักเรียนปัจจุบัน 3 คน (ไม่มีนักเรียนใช้สิทธิ์)
        # - d2 (วันหยุด): ข้ามอัตโนมัติ (skipped_holidays == 1)
        # - d3 (วันสอบ): 4 แผ่น (เช้า: ปัจจุบัน 1 แผ่น + ใช้สิทธิ์ 1 แผ่น; บ่าย: ปัจจุบัน 1 แผ่น + ใช้สิทธิ์ 1 แผ่น)
        # รวมทั้งหมด = 6 แผ่น!
        self.assertEqual(resp.context['skipped_holidays'], 1)
        sheets = resp.context['sheets']
        self.assertEqual(len(sheets), 6)

        train_sheets = [sh for sh in sheets if sh['day_mode'] == 'training']
        exam_cur_sheets = [sh for sh in sheets if sh['day_mode'] == 'exam' and sh['group_key'] == 'current']
        exam_rights_sheets = [sh for sh in sheets if sh['day_mode'] == 'exam' and sh['group_key'] == 'rights']
        self.assertEqual(len(train_sheets), 2)
        self.assertEqual(len(exam_cur_sheets), 2)
        self.assertEqual(len(exam_rights_sheets), 2)

        # ยืนยันว่าวันอบรมไม่มีนักเรียนใช้สิทธิ์เลย
        for sh in train_sheets:
            self.assertEqual(sh['group_key'], 'current')
            self.assertEqual(sh['count'], 3)
            codes = []
            for p in sh['paired_rows']:
                if p['left']:
                    codes.append(p['left']['code'])
                if p['right']:
                    codes.append(p['right']['code'])
            self.assertEqual(set(codes), {'69DS01', '69DS02', '69DS03'})
            self.assertNotIn(rec1.rights_code, codes)
            self.assertNotIn('68DS01', codes)

        # ยืนยันว่าวันสอบแยกแผ่นนักเรียนปัจจุบัน (3 คน) กับนักเรียนใช้สิทธิ์ (2 คน) คนละแผ่นกันชัดเจน
        for sh in exam_cur_sheets:
            self.assertEqual(sh['count'], 3)
            self.assertIn('นักเรียนปัจจุบัน', sh['title_line4'])
        for sh in exam_rights_sheets:
            self.assertEqual(sh['count'], 2)
            self.assertIn('นักเรียนใช้สิทธิ์', sh['title_line4'])
            codes = []
            for p in sh['paired_rows']:
                if p['left']:
                    codes.append(p['left']['code'])
                if p['right']:
                    codes.append(p['right']['code'])
            self.assertEqual(set(codes), {rec1.rights_code, rec2.rights_code})

        # 2. ทดสอบ Export PDF จริงตามช่วงวันที่ -> ต้องได้ไฟล์ PDF 6 หน้า
        pdf_resp = ac.get(reverse('food:signin_sheet') + q_range + '&export=pdf')
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertEqual(pdf_resp['Content-Type'], 'application/pdf')
        self.assertTrue(pdf_resp.content.startswith(b'%PDF-'))
        # ตรวจจำนวนหน้าใน PDF (ReportLab เขียน /Count 6 ใน Pages dictionary)
        self.assertIn(b'/Count 6', pdf_resp.content)

    def test_batch6_current_student_no_rights_inline_edit_school_modal_and_dropdown_results(self):
        from datetime import timedelta
        from decimal import Decimal
        from django.utils import timezone
        from food.models import Attendance, Score, ScoreCategory

        admin_u = User.objects.get(camp1_code='admin')
        admin_u.is_first_login = False
        admin_u.save()
        s = SystemSetting.get()
        s.current_academic_year = 2569
        s.set_editable_fields(['student_phone', 'nickname', 'is_muslim'])
        s.save()

        # นักเรียนปัจจุบัน (รุ่น 2569) vs นักเรียนรุ่นปีก่อน (รุ่น 2568)
        m_curr = make_member('69DS21', year=2569, level='ม.4')
        m_curr.school_name = 'โรงเรียนเทพศิรินทร์'
        m_curr.save()
        m_prev = make_member('68DS21', year=2568, level='ม.3')
        m_prev.school_name = 'โรงเรียนเก่า'
        m_prev.save()

        rr = rights.open_round(2569, 1, admin_u)

        # 1. นักเรียนปัจจุบันไม่ต้องยืนยันการใช้สิทธิ์ และถูกปฏิเสธหากพยายามยืนยัน (และไม่ถูก mark สละสิทธิ์)
        elig_curr = rights.check_eligibility(m_curr, rr)
        self.assertFalse(elig_curr.ok)
        self.assertEqual(elig_curr.code, 'current_student')
        with self.assertRaises(rights.RightsError):
            rights.confirm_rights(m_curr, rr)
        self.assertNotIn(m_curr, rights.eligible_unconfirmed_users(rr))

        # นักเรียนปีก่อน (2568) ยังใช้สิทธิ์ปี 2569 ได้ตามปกติ
        self.assertTrue(rights.check_eligibility(m_prev, rr).ok)

        # สร้างคะแนนและเวลาเรียนให้นักเรียนปัจจุบัน
        mc_cat = ScoreCategory.objects.create(academic_year=2569, name='คะแนนเก็บ', weight=Decimal('25'), default_max_score=Decimal('20'))
        sub1 = ScoreCategory.objects.create(academic_year=2569, parent=mc_cat, name='ส่งงาน', default_max_score=Decimal('10'))
        sub2 = ScoreCategory.objects.create(academic_year=2569, parent=mc_cat, name='จิตสาธารณะ', default_max_score=Decimal('10'))
        Score.objects.create(user=m_curr, category=sub1, exam_no=1, score=Decimal('9'), max_score=Decimal('10'))
        Score.objects.create(user=m_curr, category=sub2, exam_no=1, score=Decimal('8'), max_score=Decimal('10'))

        today = timezone.localdate()
        Attendance.objects.create(user=m_curr, academic_year=2569, date=today - timedelta(days=1),
                                  session_period='morning', status='present')
        Attendance.objects.create(user=m_curr, academic_year=2569, date=today - timedelta(days=1),
                                  session_period='afternoon', status='sick')
        Attendance.objects.create(user=m_curr, academic_year=2569, date=today,
                                  session_period='morning', status='present')

        c_curr = Client()
        c_curr.force_login(m_curr)
        html_curr = c_curr.get(reverse('composn:profile')).content.decode('utf-8')

        # 1b. หน้าโปรไฟล์ของนักเรียนปัจจุบัน: ไม่มีรายละเอียดการยืนยันสิทธิ์ / สิทธิ์ COMPOSN
        self.assertNotIn('สิทธิ์ COMPOSN', html_curr)
        self.assertNotIn('ยืนยันใช้สิทธิ์', html_curr)
        self.assertNotIn('ประวัติการใช้สิทธิ์', html_curr)

        # หน้าโปรไฟล์ของนักเรียนปีก่อน: ยังคงเห็นส่วนสิทธิ์ COMPOSN และปุ่มยืนยันใช้สิทธิ์
        c_prev = Client()
        c_prev.force_login(m_prev)
        html_prev = c_prev.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('สิทธิ์ COMPOSN', html_prev)
        self.assertIn('ยืนยันใช้สิทธิ์', html_prev)

        # 2. แก้ไขข้อมูลส่วนตัวแบบ Inline: ปุ่มปากกาข้างข้อมูลส่วนตัว + แถบเหลืองอ่อนเฉพาะช่องที่แก้ไขได้ + ไม่มีการ์ดแยก
        self.assertIn('id="btn-toggle-edit-profile"', html_curr)
        self.assertIn('bi-pencil-fill', html_curr)
        self.assertIn('#fff9db', html_curr)
        self.assertIn('pf-editable-row', html_curr)
        self.assertNotIn('<h6><i class="bi bi-pencil-square me-1"></i>แก้ไขข้อมูลส่วนตัว</h6>', html_curr)
        # บันทึกข้อมูลแบบ inline ได้จริง
        c_curr.post(reverse('composn:profile'), {
            'action': 'profile',
            'student_phone': '0891234567',
            'nickname': 'เจมส์',
            'is_muslim': 'on',
        })
        m_curr.refresh_from_db()
        self.assertEqual(m_curr.student_phone, '0891234567')
        self.assertEqual(m_curr.nickname, 'เจมส์')
        self.assertTrue(m_curr.is_muslim)

        # 3. ปุ่มคำร้องขอเปลี่ยนชื่อโรงเรียน + Modal แสดงชื่อโรงเรียนเดิมและช่องกรอกชื่อโรงเรียนใหม่
        self.assertIn('คำร้องขอเปลี่ยนชื่อโรงเรียน', html_curr)
        self.assertIn('id="schoolRequestModal"', html_curr)
        self.assertIn('ชื่อโรงเรียนเดิม', html_curr)
        self.assertIn('โรงเรียนเทพศิรินทร์', html_curr)
        self.assertIn('ชื่อโรงเรียนใหม่', html_curr)
        self.assertIn('name="new_school_name"', html_curr)

        # 4. ข้อมูลการอบรม และคะแนน: ย้ายมาที่เมนูย่อยตามปีการศึกษา, คะแนนอยู่บนสุด, คะแนนรวมติดชื่อหัวข้อใหญ่ (ไม่มีคำว่า "เต็ม"), เวลาเรียนแบบ Dropdown
        html_train = c_curr.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('id="yr-body-2569"', html_train)
        self.assertIn('id="yr-scores-2569"', html_train)
        self.assertIn('id="yr-att-2569"', html_train)
        # ส่วนคะแนนต้องอยู่เหนือส่วนเวลาเรียน
        self.assertLess(html_train.index('id="yr-scores-2569"'), html_train.index('id="yr-att-2569"'))
        # หัวข้อใหญ่มีคะแนนติดอยู่ข้างชื่อ และต้องไม่มีคำว่า "เต็ม"
        scores_chunk = html_train[html_train.index('id="yr-scores-2569"'):html_train.index('id="yr-att-2569"')]
        self.assertIn('1. คะแนนเก็บ', scores_chunk)
        self.assertIn('21.25 / 25.00', scores_chunk)
        self.assertIn('9.00 / 10.00', scores_chunk)
        self.assertNotIn('เต็ม', scores_chunk)
        # เวลาเรียนมีสรุป % และชิปรายวัน (เช้า/บ่าย) ใน Dropdown
        att_chunk = html_train[html_train.index('id="yr-att-2569"') - 1200:]
        self.assertIn('att-day-chip', att_chunk)
        self.assertIn('เช้า:', att_chunk)
        self.assertIn('บ่าย:', att_chunk)

    def test_batch8_online_certificates_and_official_announcement(self):
        """
        ทดสอบ Batch 8:
        1. นำ (ตั้งค่าที่ข้อ 6) ออกจากระบบทั้งหมด
        2. เปลี่ยนป้าย 'จำนวนครั้งขาดเรียนสูงสุด' -> 'จำนวนที่ขาดอบรมสูงสุด (ภาคเวลา)'
        3. สร้างเกียรติบัตรออนไลน์:
           - นักเรียนปัจจุบันที่ขาดอบรมไม่เกินที่กำหนด และเข้าสอบครบทุกครั้ง ได้รับเกียรติบัตร
           - นักเรียนปัจจุบันที่ขาดอบรมเกิน หรือขาดสอบ ไม่ได้รับเกียรติบัตร
           - นักเรียนใช้สิทธิ์ ไม่ต้องพิจารณาส่วนนี้ (ไม่ได้รับเกียรติบัตร)
           - ทุกคนที่ได้รับหรือไม่ได้รับเกียรติบัตร ได้รับสิทธิ์สอบท้ายค่ายทุกคน
           - สร้างและดาวน์โหลดเกียรติบัตร PDF ทั้งเดี่ยวและรวมเล่ม
        4. เอกสารประกาศผล:
           - ตั้งค่าเปลี่ยนชื่อผู้อำนวยการ, ตำแหน่ง, โรงเรียน, ศูนย์อบรม, จำนวนผู้ผ่านการคัดเลือก
           - แสดงผลเอกสารประกาศผลตามไฟล์ Word (ตราครุฑ) และรายชื่อแนบท้ายตามไฟล์ Excel (4 โลโก้)
           - Export PDF และ Export Excel
        """
        from food.certificate_pdf import check_certificate_eligibility, build_certificate_pdf
        from food.announcement_pdf import build_announcement_pdf
        from food.models import TrainingDate, Attendance

        # 1. ตรวจสอบป้ายชื่อฟิลด์ใน Model
        f_max_abs = SystemSetting._meta.get_field('max_absences')
        self.assertEqual(f_max_abs.verbose_name, 'จำนวนที่ขาดอบรมสูงสุด (ภาคเวลา)')
        f_max_qual = SystemSetting._meta.get_field('max_qualified_count')
        self.assertEqual(f_max_qual.verbose_name, 'จำนวนผู้ที่ผ่านการคัดเลือก (คน)')

        # ตรวจสอบการนำ (ตั้งค่าที่ข้อ 6) ออกจากหน้า score_categories
        admin = User.objects.create_superuser('adm_b8', password='pass', is_first_login=False)
        ac = Client()
        ac.force_login(admin)
        resp_sc = ac.get(reverse('food:score_categories') + '?year=2569')
        self.assertEqual(resp_sc.status_code, 200)
        self.assertNotIn('ข้อ 6', resp_sc.content.decode('utf-8'))
        self.assertNotIn('ข้อ6', resp_sc.content.decode('utf-8'))

        # 2. ตั้งค่าระบบและเอกสารประกาศผล
        setting = SystemSetting.get()
        setting.max_absences = 4
        setting.max_qualified_count = 31
        setting.announcement_school_name = 'โรงเรียนเทพศิรินทร์'
        setting.announcement_director_name = 'นายวิธาน พรหมสินธุศักดิ์'
        setting.announcement_director_position = 'ผู้อำนวยการโรงเรียนเทพศิรินทร์\nประธานศูนย์โรงเรียนเทพศิรินทร์ สอวน.วิชาคอมพิวเตอร์ ค่าย ๑'
        setting.announcement_center_name = 'ศูนย์โรงเรียนเทพศิรินทร์-สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหาร ลาดกระบัง'
        setting.announcement_date_text = 'ประกาศ ณ วันที่ ๒ มกราคม พ.ศ. ๒๕๗๐'
        setting.save()

        # 3. เตรียมข้อมูลนักเรียนและเวลาเรียน
        # นักเรียนปัจจุบัน 1: เข้าครบ สอบครบ -> ได้รับเกียรติบัตร
        m_good = make_member('69DS91')
        m_good.first_name, m_good.last_name = 'สมชาย', 'เรียนดี'
        m_good.save()

        # นักเรียนปัจจุบัน 2: ขาดอบรมเกิน 4 ครั้ง -> ไม่ได้รับเกียรติบัตร แต่มีสิทธิ์สอบท้ายค่าย
        m_abs = make_member('69DS92')
        m_abs.first_name, m_abs.last_name = 'สมศักดิ์', 'ขาดบ่อย'
        m_abs.save()

        # นักเรียนปัจจุบัน 3: ขาดสอบ 1 ช่วง -> ไม่ได้รับเกียรติบัตร แต่มีสิทธิ์สอบท้ายค่าย
        m_miss_exam = make_member('69DS93')
        m_miss_exam.first_name, m_miss_exam.last_name = 'สมหมาย', 'ขาดสอบ'
        m_miss_exam.save()

        # นักเรียนใช้สิทธิ์ (รุ่น 68) -> ไม่พิจารณาเกียรติบัตร แต่มีสิทธิ์สอบท้ายค่าย
        m_rights = make_member('68DS99')
        m_rights.first_name, m_rights.last_name = 'สมศรี', 'ใช้สิทธิ์'
        m_rights.academic_year = 2568
        m_rights.save()
        UseRecord.objects.create(
            user=m_rights, academic_year=2569, camp_level=1, rights_code=m_rights.camp1_code, is_used=True,
        )

        # สร้างวันอบรม 2 วัน และวันสอบ 1 วัน (เช้า+บ่าย)
        d_train1 = date(2026, 10, 1)
        d_train2 = date(2026, 10, 2)
        d_exam = date(2026, 10, 3)
        TrainingDate.objects.create(academic_year=2569, date=d_train1, event_type='training', title='อบรม 1')
        TrainingDate.objects.create(academic_year=2569, date=d_train2, event_type='training', title='อบรม 2')
        TrainingDate.objects.create(academic_year=2569, date=d_exam, event_type='exam', exam_period='both', title='สอบปลายค่าย')

        # บันทึกเวลาเรียน: m_good เข้าครบทุกวันและเข้าสอบทั้งเช้า/บ่าย
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_train1, session_period='morning', status='present')
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_train1, session_period='afternoon', status='present')
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_train2, session_period='morning', status='present')
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_train2, session_period='afternoon', status='present')
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_exam, session_period='morning', status='present')
        Attendance.objects.create(user=m_good, academic_year=2569, date=d_exam, session_period='afternoon', status='present')

        # บันทึกเวลาเรียน: m_abs ขาดอบรม 5 ครั้ง (เกิน 4 ครั้ง) แต่เข้าสอบครบ
        for i in range(1, 6):
            dt_dummy = date(2026, 9, 20 + i)
            TrainingDate.objects.create(academic_year=2569, date=dt_dummy, event_type='training')
            Attendance.objects.create(user=m_abs, academic_year=2569, date=dt_dummy, session_period='morning', status='absent')
        Attendance.objects.create(user=m_abs, academic_year=2569, date=d_exam, session_period='morning', status='present')
        Attendance.objects.create(user=m_abs, academic_year=2569, date=d_exam, session_period='afternoon', status='present')

        # บันทึกเวลาเรียน: m_miss_exam ไม่ขาดอบรม แต่ขาดสอบภาคบ่าย
        Attendance.objects.create(user=m_miss_exam, academic_year=2569, date=d_train1, session_period='morning', status='present')
        Attendance.objects.create(user=m_miss_exam, academic_year=2569, date=d_exam, session_period='morning', status='present')
        Attendance.objects.create(user=m_miss_exam, academic_year=2569, date=d_exam, session_period='afternoon', status='absent')

        # 4. ทดสอบฟังก์ชัน check_certificate_eligibility
        elig_good = check_certificate_eligibility(m_good, 2569)
        self.assertTrue(elig_good['is_eligible'])
        self.assertTrue(elig_good['final_exam_eligible'])

        elig_abs = check_certificate_eligibility(m_abs, 2569)
        self.assertFalse(elig_abs['is_eligible'])
        self.assertFalse(elig_abs['absence_passed'])
        self.assertTrue(elig_abs['final_exam_eligible'])  # กฎสำคัญ: ต้องมีสิทธิ์สอบท้ายค่าย

        elig_miss_exam = check_certificate_eligibility(m_miss_exam, 2569)
        self.assertFalse(elig_miss_exam['is_eligible'])
        self.assertFalse(elig_miss_exam['exams_passed'])
        self.assertTrue(elig_miss_exam['final_exam_eligible'])  # กฎสำคัญ: ต้องมีสิทธิ์สอบท้ายค่าย

        elig_rights = check_certificate_eligibility(m_rights, 2569)
        self.assertFalse(elig_rights['is_eligible'])
        self.assertFalse(elig_rights['is_current'])
        self.assertTrue(elig_rights['final_exam_eligible'])  # กฎสำคัญ: ต้องมีสิทธิ์สอบท้ายค่าย

        # 5. ทดสอบสร้าง PDF เกียรติบัตร
        pdf_single = build_certificate_pdf([m_good], 2569, setting)
        self.assertGreater(len(pdf_single), 1000)

        pdf_batch = build_certificate_pdf([m_good], 2569, setting)
        self.assertGreater(len(pdf_batch), 1000)

        # 6. ทดสอบหน้าเว็บ Admin: admin_certificates (ชื่อหน้า "เกียรติบัตร", จัดกลุ่มตามสิทธิ์, ไม่มีคอลัมน์สิทธิ์สอบท้ายค่าย)
        resp_admin_cert = ac.get(reverse('food:admin_certificates') + '?year=2569')
        self.assertEqual(resp_admin_cert.status_code, 200)
        content_admin = resp_admin_cert.content.decode('utf-8')
        self.assertIn('เกียรติบัตร', content_admin)
        self.assertNotIn('สร้างเกียรติบัตรออนไลน์', content_admin)
        self.assertNotIn('สิทธิ์สอบท้ายค่าย', content_admin)
        self.assertNotIn('มีสิทธิ์สอบ', content_admin)
        self.assertNotIn('>ประเภท<', content_admin)
        self.assertNotIn('>สิทธิ์รับเกียรติบัตร<', content_admin)
        self.assertIn('กลุ่มนักเรียนปัจจุบัน — ได้รับเกียรติบัตร', content_admin)
        self.assertIn('กลุ่มนักเรียนปัจจุบัน — ไม่ได้รับเกียรติบัตร', content_admin)
        self.assertIn('กลุ่มนักเรียนใช้สิทธิ์ — ไม่พิจารณาเกียรติบัตร', content_admin)
        grp_elig_chunk = content_admin.split('id="cert-grp-eligible"')[1].split('id="cert-grp-ineligible"')[0]
        grp_inelig_chunk = content_admin.split('id="cert-grp-ineligible"')[1].split('id="cert-grp-rights"')[0]
        grp_rights_chunk = content_admin.split('id="cert-grp-rights"')[1]
        self.assertIn('69DS91', grp_elig_chunk)
        self.assertNotIn('69DS92', grp_elig_chunk)
        self.assertIn('69DS92', grp_inelig_chunk)
        self.assertIn('69DS93', grp_inelig_chunk)
        self.assertIn('68DS99', grp_rights_chunk)

        # Export Excel รายชื่อเกียรติบัตร (แยกแผ่นตามกลุ่ม และไม่มีคอลัมน์สิทธิ์สอบท้ายค่าย)
        resp_cert_xlsx = ac.get(reverse('food:admin_certificates') + '?year=2569&export=xlsx')
        self.assertEqual(resp_cert_xlsx.status_code, 200)
        self.assertEqual(resp_cert_xlsx['Content-Type'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        import openpyxl
        wb_cert = openpyxl.load_workbook(io.BytesIO(resp_cert_xlsx.content))
        self.assertEqual(wb_cert.sheetnames, ['ได้รับเกียรติบัตร', 'ไม่ได้รับเกียรติบัตร', 'นักเรียนใช้สิทธิ์'])
        for ws_c in wb_cert.worksheets:
            for row_c in ws_c.iter_rows(values_only=True):
                for cell_c in row_c:
                    self.assertNotIn('สิทธิ์สอบท้ายค่าย', str(cell_c or ''))

        # ดาวน์โหลดเกียรติบัตรเดี่ยวผ่าน View
        resp_dl = ac.get(reverse('food:certificate_download', args=[m_good.pk]) + '?year=2569')
        self.assertEqual(resp_dl.status_code, 200)
        self.assertEqual(resp_dl['Content-Type'], 'application/pdf')

        # ดาวน์โหลดรวมเล่ม Batch Export ผ่าน View
        resp_batch = ac.get(reverse('food:certificates_export_all') + '?year=2569')
        self.assertEqual(resp_batch.status_code, 200)
        self.assertEqual(resp_batch['Content-Type'], 'application/pdf')

        # 7. ทดสอบนักเรียนดาวน์โหลดเกียรติบัตรตนเอง (my_certificate)
        # ค่าเริ่มต้นปีการศึกษาคือ ปิดการแสดง -> นักเรียนยังต้องไม่เห็นปุ่มและโหลดตรงไม่ได้ (302)
        mc_good = Client()
        mc_good.force_login(m_good)
        resp_unpub = mc_good.get(reverse('food:my_certificate') + '?year=2569')
        self.assertEqual(resp_unpub.status_code, 302)
        prof_unpub = mc_good.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('ดาวน์โหลดเกียรติบัตร (PDF)', prof_unpub)

        # ผู้ดูแลระบบเปิดการแสดงเกียรติบัตรปีการศึกษา 2569
        resp_pub_toggle = ac.post(reverse('food:admin_certificates'), {
            'action': 'toggle_cert_publish',
            'year': '2569',
            'tab': 'students',
        })
        self.assertEqual(resp_pub_toggle.status_code, 302)

        resp_my_cert = mc_good.get(reverse('food:my_certificate') + '?year=2569')
        self.assertEqual(resp_my_cert.status_code, 200)
        self.assertEqual(resp_my_cert['Content-Type'], 'application/pdf')

        # นักเรียนที่ขาดอบรมเกิน เข้าโหลดเกียรติบัตรต้อง redirect พร้อมข้อความแจ้งเตือน
        mc_abs = Client()
        mc_abs.force_login(m_abs)
        resp_my_abs = mc_abs.get(reverse('food:my_certificate') + '?year=2569')
        self.assertEqual(resp_my_abs.status_code, 302)

        # ตรวจสอบหน้าข้อมูลส่วนตัวต้องมีปุ่มดาวน์โหลดเกียรติบัตร และยังไม่แสดงสิทธิ์สอบท้ายค่ายจนกว่าจะประกาศผล
        prof_good = mc_good.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('ดาวน์โหลดเกียรติบัตร (PDF)', prof_good)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 1', prof_good)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 2', prof_good)

        # 8. ทดสอบเอกสารประกาศผล (score_announcement)
        resp_ann = ac.get(reverse('food:score_announcement') + '?year=2569')
        self.assertEqual(resp_ann.status_code, 200)
        ann_html = resp_ann.content.decode('utf-8')
        self.assertIn('ประกาศโรงเรียนเทพศิรินทร์', ann_html)
        self.assertIn('นายวิธาน พรหมสินธุศักดิ์', ann_html)
        self.assertIn('ประกาศ ณ วันที่ 2 มกราคม พ.ศ. 2570', ann_html)
        self.assertIn('ศูนย์โรงเรียนเทพศิรินทร์-สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหาร ลาดกระบัง', ann_html)

        # Export PDF ประกาศผล
        resp_ann_pdf = ac.get(reverse('food:score_announcement') + '?year=2569&export=pdf')
        self.assertEqual(resp_ann_pdf.status_code, 200)
        self.assertEqual(resp_ann_pdf['Content-Type'], 'application/pdf')

        # Export Excel ประกาศผล
        resp_ann_xlsx = ac.get(reverse('food:score_announcement') + '?year=2569&export=xlsx')
        self.assertEqual(resp_ann_xlsx.status_code, 200)
        self.assertEqual(resp_ann_xlsx['Content-Type'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    def test_batch9_certificate_settings_background_and_lecturer_committee_crud(self):
        """
        ทดสอบ Batch 9:
        1. ตั้งค่าเกียรติบัตร: รูปพื้นหลัง + ปรับตำแหน่งแกน Y / ขนาดตัวอักษรทั้ง 4 ส่วน (TH Sarabun) + ค่าเริ่มต้นข้อ 3-4
        2. สิทธิ์: Staff ที่ไม่ใช่ Admin ห้ามบันทึก save_settings (403)
        3. รีเซ็ตตำแหน่ง/ขนาด (reset_settings) โดยไม่ลบรูปพื้นหลัง
        4. CRUD รายชื่อวิทยากรและคณะกรรมการ (เพิ่ม/แก้ไข/ลบ/ล้าง) และใช้ข้อ 4 ร่วมกัน
        5. นำเข้าจาก Excel (.xlsx) และ CSV (.csv) + ดาวน์โหลดไฟล์ตัวอย่าง + Export Excel
        6. สร้าง PDF จริงทั้งตอนมีรูปพื้นหลังและหลังลบรูปพื้นหลัง (นักเรียน / วิทยากร / คณะกรรมการ / Preview PDF)
        7. เมนู "ตั้งค่าเกียรติบัตร" แสดงใน Admin navbar/sidebar และหน้าตั้งค่าระบบหลัก
        """
        import io
        import tempfile
        from PIL import Image
        from django.core.files.uploadedfile import SimpleUploadedFile
        from food.models import CertificateRecipient, CertificateSetting

        admin = User.objects.create_superuser('adm_b9', password='pass', is_first_login=False)
        staff = User.objects.create_user(
            'stf_b9', password='pass', role='staff', is_first_login=False, staff_can_scores=True
        )
        m_stu = make_member('69DS95')
        m_stu.first_name, m_stu.last_name, m_stu.school_name = 'เด็กชายนิธิดล', 'ธรรมภูษิต', 'โรงเรียนสวนกุหลาบวิทยาลัย'
        m_stu.save()

        ac = Client()
        ac.force_login(admin)
        sc = Client()
        sc.force_login(staff)

        with tempfile.TemporaryDirectory() as tmp_media:
            with self.settings(MEDIA_ROOT=tmp_media):
                # สร้างรูปพื้นหลังจำลอง 1485x1050 (อัตราส่วน A4 แนวนอน)
                img_buf = io.BytesIO()
                Image.new('RGB', (1485, 1050), color=(250, 248, 240)).save(img_buf, format='PNG')
                bg_file = SimpleUploadedFile('cert_bg.png', img_buf.getvalue(), content_type='image/png')

                # 1. บันทึกตั้งค่าเกียรติบัตรผ่าน Admin
                r_save = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
                    'action': 'save_settings',
                    'tab': 'settings',
                    'year': 2569,
                    'background_image': bg_file,
                    'line1_y': '96.5',
                    'line1_font_size': '36.0',
                    'line1_bold': '1',
                    'line1_color': '#112233',
                    'line2_y': '114.0',
                    'line2_font_size': '24.0',
                    'line2_bold': '1',
                    'line2_color': '#223344',
                    'student_default_province': 'กรุงเทพมหานคร',
                    'line3_y': '131.5',
                    'line3_font_size': '20.5',
                    'line3_line_spacing': '10.5',
                    'line3_bold': '1',
                    'line3_color': '#000000',
                    'student_line3_text': 'ได้ผ่านโครงการโอลิมปิกวิชาการ สอวน. ค่าย ๑ ปีการศึกษา {year_thai}\nณ ศูนย์โรงเรียนเทพศิรินทร์',
                    'lecturer_line3_text': 'ได้เข้าร่วมเป็นวิทยากรอบรมนักเรียน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\nณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์ จังหวัดกรุงเทพมหานคร',
                    'committee_line3_text': 'ได้เข้าร่วมเป็นคณะกรรมการดำเนินงาน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\nระหว่างวันที่ ๒๘ กันยายน - ๑๔ ตุลาคม {year_thai}\nณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์ จังหวัดกรุงเทพมหานคร',
                    'line4_text': 'ให้ไว้ ณ วันที่ ๑๔ ตุลาคม {year_thai}',
                    'line4_y': '164.0',
                    'line4_font_size': '21.0',
                    'line4_bold': '1',
                    'line4_color': '#000000',
                })
                self.assertEqual(r_save.status_code, 302)

                cs = CertificateSetting.get()
                self.assertTrue(bool(cs.background_image))
                self.assertAlmostEqual(cs.line1_y, 96.5)
                self.assertAlmostEqual(cs.line1_font_size, 36.0)
                self.assertEqual(cs.line1_color, '#112233')
                self.assertAlmostEqual(cs.line2_y, 114.0)
                self.assertAlmostEqual(cs.line3_y, 131.5)
                self.assertAlmostEqual(cs.line3_line_spacing, 10.5)
                self.assertAlmostEqual(cs.line4_y, 164.0)
                self.assertEqual(cs.render_text(cs.line4_text, 2569), 'ให้ไว้ ณ วันที่ ๑๔ ตุลาคม ๒๕๖๙')
                self.assertEqual(cs.format_student_school('โรงเรียนสวนกุหลาบวิทยาลัย'), 'โรงเรียนสวนกุหลาบวิทยาลัย กรุงเทพมหานคร')

                # 2. Staff ที่ไม่ใช่ Admin ห้าม save_settings / reset_settings (403)
                r_forbidden = sc.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
                    'action': 'save_settings', 'tab': 'settings', 'year': 2569,
                })
                self.assertEqual(r_forbidden.status_code, 403)
                r_tab_staff = sc.get(reverse('food:admin_certificates') + '?year=2569&tab=settings')
                self.assertEqual(r_tab_staff.status_code, 302)

                # 3. รีเซ็ตแกน Y/ขนาดตัวอักษร ต้องกลับค่าเริ่มต้นแต่ไม่ลบรูปพื้นหลัง
                r_reset = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
                    'action': 'reset_settings', 'tab': 'settings', 'year': 2569,
                })
                self.assertEqual(r_reset.status_code, 302)
                cs.refresh_from_db()
                self.assertTrue(bool(cs.background_image))
                self.assertAlmostEqual(cs.line1_y, 94.0)
                self.assertAlmostEqual(cs.line1_font_size, 34.0)

                # 4. CRUD วิทยากร & คณะกรรมการ
                r_add_lec = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
                    'action': 'add_recipient',
                    'year': 2569,
                    'recipient_type': 'lecturer',
                    'full_name': 'นายณัฏฐ์กวี ลืนภูเขียว',
                    'school_province': 'โรงเรียนไตรมิตรวิทยาลัย กรุงเทพมหานคร',
                    'description': '',  # เว้นว่างเพื่อใช้ค่าเริ่มต้นของกลุ่มวิทยากร
                })
                self.assertEqual(r_add_lec.status_code, 302)
                lec = CertificateRecipient.objects.get(full_name='นายณัฏฐ์กวี ลืนภูเขียว')
                self.assertEqual(lec.recipient_type, 'lecturer')
                self.assertIn('วิทยากรอบรมนักเรียน', lec.resolved_description(cs))

                r_add_com = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=committee', {
                    'action': 'add_recipient',
                    'year': 2569,
                    'recipient_type': 'committee',
                    'full_name': 'นายอรรถธรรม ตั้งปริมณฑล',
                    'school_province': 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
                    'description': 'ได้เข้าร่วมเป็นคณะกรรมการดำเนินงาน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\nระหว่างวันที่ ๒๘ กันยายน - ๑๔ ตุลาคม {year_thai}',
                })
                self.assertEqual(r_add_com.status_code, 302)
                com = CertificateRecipient.objects.get(full_name='นายอรรถธรรม ตั้งปริมณฑล')

                # แก้ไขข้อมูล
                r_edit = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
                    'action': 'edit_recipient',
                    'year': 2569,
                    'recipient_id': lec.pk,
                    'recipient_type': 'lecturer',
                    'full_name': 'นายณัฏฐ์กวี ลืนภูเขียว (แก้ไข)',
                    'school_province': 'โรงเรียนไตรมิตรวิทยาลัย กรุงเทพมหานคร',
                    'description': 'ได้เข้าร่วมเป็นวิทยากรพิเศษ',
                })
                self.assertEqual(r_edit.status_code, 302)
                lec.refresh_from_db()
                self.assertEqual(lec.full_name, 'นายณัฏฐ์กวี ลืนภูเขียว (แก้ไข)')
                self.assertEqual(lec.resolved_description(cs), 'ได้เข้าร่วมเป็นวิทยากรพิเศษ')

                # 5. นำเข้า Excel (.xlsx) และ CSV (.csv) + Template + Export Excel
                xlsx_bytes = make_xlsx(
                    ['ประเภท', 'ชื่อ-นามสกุล', 'โรงเรียน จังหวัด', 'รายละเอียด'],
                    [
                        ['วิทยากร', 'อ.สมปอง วิทยากร', 'จุฬาลงกรณ์มหาวิทยาลัย กรุงเทพมหานคร', ''],
                        ['คณะกรรมการ', 'นายสมหมาย กรรมการ', 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร', ''],
                    ],
                )
                r_imp_xlsx = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
                    'action': 'import_recipients',
                    'year': 2569,
                    'default_type': 'lecturer',
                    'import_file': SimpleUploadedFile('rec.xlsx', xlsx_bytes.getvalue(),
                                                      content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                })
                self.assertEqual(r_imp_xlsx.status_code, 302)
                self.assertTrue(CertificateRecipient.objects.filter(full_name='อ.สมปอง วิทยากร', recipient_type='lecturer').exists())
                self.assertTrue(CertificateRecipient.objects.filter(full_name='นายสมหมาย กรรมการ', recipient_type='committee').exists())

                csv_bytes = 'ชื่อ-นามสกุล,โรงเรียน,จังหวัด,รายละเอียด\nนายมานะ ใจดี,โรงเรียนสวนกุหลาบวิทยาลัย,กรุงเทพมหานคร,คณะกรรมการฝ่ายวิชาการ\n'.encode('utf-8-sig')
                r_imp_csv = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=committee', {
                    'action': 'import_recipients',
                    'year': 2569,
                    'default_type': 'committee',
                    'import_file': SimpleUploadedFile('rec.csv', csv_bytes, content_type='text/csv'),
                })
                self.assertEqual(r_imp_csv.status_code, 302)
                c_csv = CertificateRecipient.objects.get(full_name='นายมานะ ใจดี')
                self.assertEqual(c_csv.recipient_type, 'committee')
                self.assertEqual(c_csv.school_province, 'โรงเรียนสวนกุหลาบวิทยาลัย กรุงเทพมหานคร')

                r_tpl = ac.get(reverse('food:certificate_recipients_template'))
                self.assertEqual(r_tpl.status_code, 200)
                self.assertEqual(r_tpl['Content-Type'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                r_exp_xlsx = ac.get(reverse('food:certificate_recipients_export_xlsx') + '?year=2569&type=committee')
                self.assertEqual(r_exp_xlsx.status_code, 200)
                self.assertEqual(r_exp_xlsx['Content-Type'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                # 6. ทดสอบสร้าง PDF จริงตอนมีรูปพื้นหลัง
                for url in (
                    reverse('food:certificate_download', args=[m_stu.pk]) + '?year=2569',
                    reverse('food:certificates_export_all') + '?year=2569',
                    reverse('food:certificate_recipient_download', args=[lec.pk]),
                    reverse('food:certificate_recipients_export_pdf') + '?year=2569&type=committee',
                    reverse('food:certificate_preview_pdf') + '?year=2569&group=lecturer&line1_y=92',
                    reverse('food:certificate_preview_pdf') + '?year=2569&group=committee',
                    reverse('food:certificate_preview_pdf') + '?year=2569&group=student',
                ):
                    r_pdf = ac.get(url)
                    self.assertEqual(r_pdf.status_code, 200, url)
                    self.assertEqual(r_pdf['Content-Type'], 'application/pdf', url)
                    self.assertTrue(r_pdf.content.startswith(b'%PDF-'), url)
                    self.assertGreater(len(r_pdf.content), 1000, url)

                # ทดสอบลบรูปพื้นหลังแล้วสร้าง PDF fallback
                r_clear = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
                    'action': 'save_settings',
                    'tab': 'settings',
                    'year': 2569,
                    'clear_background_image': '1',
                    'line1_y': '94.0', 'line1_font_size': '34.0', 'line1_bold': '1', 'line1_color': '#000000',
                    'line2_y': '111.0', 'line2_font_size': '22.0', 'line2_bold': '1', 'line2_color': '#000000',
                    'line3_y': '127.0', 'line3_font_size': '21.0', 'line3_line_spacing': '11.0', 'line3_bold': '1', 'line3_color': '#000000',
                    'line4_y': '162.0', 'line4_font_size': '21.0', 'line4_bold': '1', 'line4_color': '#000000',
                    'line4_text': 'ให้ไว้ ณ วันที่ ๑๔ ตุลาคม {year_thai}',
                })
                self.assertEqual(r_clear.status_code, 302)
                cs.refresh_from_db()
                self.assertFalse(bool(cs.background_image))

                r_fb_pdf = ac.get(reverse('food:certificate_recipient_download', args=[com.pk]))
                self.assertEqual(r_fb_pdf.status_code, 200)
                self.assertTrue(r_fb_pdf.content.startswith(b'%PDF-'))

                # ลบรายชื่อเดี่ยว
                r_del = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=committee', {
                    'action': 'delete_recipient', 'year': 2569, 'recipient_id': c_csv.pk,
                })
                self.assertEqual(r_del.status_code, 302)
                self.assertFalse(CertificateRecipient.objects.filter(pk=c_csv.pk).exists())

                # 7. ตรวจสอบหน้า Settings Tab และเมนู Admin
                r_settings_tab = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=settings')
                self.assertEqual(r_settings_tab.status_code, 200)
                html_st = r_settings_tab.content.decode('utf-8')
                self.assertIn('THSarabunNewCert', html_st)
                self.assertIn('cert-preview-stage', html_st)
                self.assertIn('ตั้งค่าเกียรติบัตร', html_st)

                r_sys = ac.get(reverse('composn:admin_settings'))
                self.assertEqual(r_sys.status_code, 200)
                self.assertIn(reverse('food:admin_certificates'), r_sys.content.decode('utf-8'))

    def test_batch10_transfer_certificate_recipients_from_previous_year(self):
        """
        ทดสอบการโอนรายชื่อวิทยากรและคณะกรรมการจากปีที่ผ่านมา:
        1. สร้างรายชื่อปี 2568 (วิทยากร 2, คณะกรรมการ 2)
        2. ตรวจว่าหน้าปี 2569 มีปุ่มโอนและตัวเลือกปี 2568 และ get_academic_year_choices รวมปี 2568
        3. โอนเฉพาะวิทยากร 2568 -> 2569 (ข้อมูลปี 2568 ยังอยู่ครบ)
        4. โอนซ้ำโดยเปิด skip_duplicates -> ไม่เพิ่มรายการซ้ำ
        5. โอนทั้งคู่ (both) พร้อม skip_duplicates + reset_description -> ข้ามวิทยากรซ้ำ 2 คน และสร้างคณะกรรมการใหม่ 2 คนโดย description ว่าง
        6. โอนจากปีเดียวกัน (2569 -> 2569) -> ปฏิเสธและไม่เพิ่มรายการ
        7. ดาวน์โหลด PDF รายชื่อที่โอนมาในปี 2569 ได้สำเร็จ (status 200, b'%PDF-')
        """
        from composn.context_processors import get_academic_year_choices
        from food.models import CertificateRecipient

        admin = User.objects.create_superuser('adm_b10', password='pass', is_first_login=False)
        ac = Client()
        ac.force_login(admin)

        # 1. สร้างวิทยากร 2 และคณะกรรมการ 2 ในปี 2568
        CertificateRecipient.objects.create(
            academic_year=2568, recipient_type='lecturer',
            full_name='อ.สมชาย ปี68', school_province='โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
            description='ได้เข้าร่วมเป็นวิทยากรพิเศษ ปี 2568', sort_order=1,
        )
        CertificateRecipient.objects.create(
            academic_year=2568, recipient_type='lecturer',
            full_name='อ.สมหญิง ปี68', school_province='จุฬาลงกรณ์มหาวิทยาลัย กรุงเทพมหานคร',
            description='', sort_order=2,
        )
        CertificateRecipient.objects.create(
            academic_year=2568, recipient_type='committee',
            full_name='นายวิชัย กรรมการ', school_province='โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
            description='คณะกรรมการฝ่ายสถานที่', sort_order=1,
        )
        CertificateRecipient.objects.create(
            academic_year=2568, recipient_type='committee',
            full_name='นางสาวสุดา กรรมการ', school_province='โรงเรียนเตรียมอุดมศึกษา กรุงเทพมหานคร',
            description='', sort_order=2,
        )

        # 2. ตรวจปีการศึกษาใน get_academic_year_choices และหน้า UI ปี 2569
        self.assertIn(2568, get_academic_year_choices(2569))
        r_page = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=lecturer')
        self.assertEqual(r_page.status_code, 200)
        html_page = r_page.content.decode('utf-8')
        self.assertIn('โอนรายชื่อจากปีที่ผ่านมา', html_page)
        self.assertIn('ปีการศึกษา 2568', html_page)

        # 3. โอนเฉพาะวิทยากร 2568 -> 2569 (คงข้อความ description เดิม)
        r_tr_lec = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
            'action': 'transfer_recipients',
            'year': 2569,
            'tab': 'lecturer',
            'source_year': 2568,
            'transfer_scope': 'lecturer',
            'skip_duplicates': '1',
        })
        self.assertEqual(r_tr_lec.status_code, 302)
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569, recipient_type='lecturer').count(), 2)
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569, recipient_type='committee').count(), 0)
        # ข้อมูลปี 2568 ต้องยังอยู่ครบไม่ถูกลบ
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2568).count(), 4)
        lec_69 = CertificateRecipient.objects.get(academic_year=2569, full_name='อ.สมชาย ปี68')
        self.assertEqual(lec_69.description, 'ได้เข้าร่วมเป็นวิทยากรพิเศษ ปี 2568')

        # 4. โอนซ้ำอีกครั้งโดยเปิด skip_duplicates=1 -> ไม่สร้างรายการซ้ำ
        r_dup = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
            'action': 'transfer_recipients',
            'year': 2569,
            'tab': 'lecturer',
            'source_year': 2568,
            'transfer_scope': 'lecturer',
            'skip_duplicates': '1',
        })
        self.assertEqual(r_dup.status_code, 302)
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569, recipient_type='lecturer').count(), 2)

        # 5. โอนทั้งคู่ (both) พร้อม skip_duplicates=1 และ reset_description=1
        r_both = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=committee', {
            'action': 'transfer_recipients',
            'year': 2569,
            'tab': 'committee',
            'source_year': 2568,
            'transfer_scope': 'both',
            'skip_duplicates': '1',
            'reset_description': '1',
        })
        self.assertEqual(r_both.status_code, 302)
        # วิทยากรยังคงมี 2 คนเท่าเดิม และคณะกรรมการถูกสร้างเพิ่ม 2 คนโดย description ว่าง
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569, recipient_type='lecturer').count(), 2)
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569, recipient_type='committee').count(), 2)
        com_69 = CertificateRecipient.objects.get(academic_year=2569, full_name='นายวิชัย กรรมการ')
        self.assertEqual(com_69.description, '')

        # 6. ปฏิเสธการโอนจากปีเดียวกัน (2569 -> 2569)
        r_same = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=lecturer', {
            'action': 'transfer_recipients',
            'year': 2569,
            'tab': 'lecturer',
            'source_year': 2569,
            'transfer_scope': 'both',
        })
        self.assertEqual(r_same.status_code, 302)
        self.assertEqual(CertificateRecipient.objects.filter(academic_year=2569).count(), 4)

        # 7. ดาวน์โหลด PDF จากรายชื่อที่โอนมาในปี 2569
        for url in (
            reverse('food:certificate_recipient_download', args=[lec_69.pk]),
            reverse('food:certificate_recipient_download', args=[com_69.pk]),
            reverse('food:certificate_recipients_export_pdf') + '?year=2569&type=lecturer',
            reverse('food:certificate_recipients_export_pdf') + '?year=2569&type=committee',
        ):
            r_pdf = ac.get(url)
            self.assertEqual(r_pdf.status_code, 200, url)
            self.assertEqual(r_pdf['Content-Type'], 'application/pdf', url)
            self.assertTrue(r_pdf.content.startswith(b'%PDF-'), url)

    def test_announcement_date_inline_edit_sorting_and_hidden_scores(self):
        import io
        import openpyxl
        from food.models import ScoreCategory, approve_year_scores
        from food.services import save_score

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        staff = User.objects.create_user(
            'stf_ann', password='pass', role='staff', is_first_login=False, staff_can_scores=True
        )

        setting = SystemSetting.get()
        setting.current_academic_year = 2569
        setting.max_qualified_count = 4
        setting.passing_score = 0
        setting.announcement_date_text = ''
        setting.announcement_list_title = ''
        setting.save()

        # นักเรียนปัจจุบัน (69DS02, 69DS10) และนักเรียนใช้สิทธิ์ (67DS15, 68DS05)
        cur_high = make_member('69DS10', year=2569, level='ม.4')
        cur_low = make_member('69DS02', year=2569, level='ม.4')
        rig_old = make_member('67DS15', year=2567, level='ม.5')
        rig_new = make_member('68DS05', year=2568, level='ม.5')

        UseRecord.objects.create(
            user=rig_old, academic_year=2569, camp_level=1, rights_code='67DS15', is_used=True
        )
        UseRecord.objects.create(
            user=rig_new, academic_year=2569, camp_level=1, rights_code='68DS05', is_used=True
        )

        cat = ScoreCategory.objects.create(name='สอบรวม', academic_year=2569, weight=100, default_max_score=100)
        save_score(rig_new, cat, 1, 95, 100, admin)
        save_score(cur_high, cat, 1, 90, 100, admin)
        save_score(rig_old, cat, 1, 80, 100, admin)
        save_score(cur_low, cat, 1, 70, 100, admin)

        ac = Client()
        ac.force_login(admin)
        sc = Client()
        sc.force_login(staff)

        # 1. วันที่ประกาศค่าเริ่มต้นต้องเป็นรูปแบบไทยที่ถูกต้อง ไม่มีอักขระเพี้ยน "| หรือ p.m.
        html = ac.get(reverse('food:score_announcement') + '?year=2569').content.decode('utf-8')
        self.assertIn('ประกาศ ณ วันที่ ', html)
        self.assertNotIn('"|', html)
        self.assertNotIn('p.m.', html)

        # 2. ตรวจว่าไม่แสดงคอลัมน์ คะแนนรวม และ ผลการคัดเลือก ในตารางประกาศผล
        self.assertNotIn('คะแนนรวม</th>', html)
        self.assertNotIn('ผลการคัดเลือก</th>', html)

        # 3. ตรวจลำดับการแสดงผลบนหน้าเว็บ: นักเรียนปัจจุบัน (69DS02 -> 69DS10) แล้วตามด้วยนักเรียนใช้สิทธิ์ (67DS15 -> 68DS05)
        expected_order = ['69DS02', '69DS10', '67DS15', '68DS05']
        positions = [html.index(code) for code in expected_order]
        self.assertEqual(positions, sorted(positions))

        # 4. ตรวจ Export Excel ไม่มีคะแนนรวม/สถานะ และเรียงลำดับถูกต้อง
        resp_xlsx = ac.get(reverse('food:score_announcement') + '?year=2569&export=xlsx')
        self.assertEqual(resp_xlsx.status_code, 200)
        wb = openpyxl.load_workbook(io.BytesIO(resp_xlsx.content))
        ws = wb.active
        rows_xlsx = list(ws.iter_rows(values_only=True))
        self.assertEqual(
            list(rows_xlsx[0]),
            ['ลำดับ', 'เลขประจำตัวค่าย 2', 'เลขประจำตัวค่าย 1', 'ชื่อ นามสกุล', 'โรงเรียน', 'ระดับ'],
        )
        self.assertEqual([r[2] for r in rows_xlsx[1:]], expected_order)

        # 5. อนุมัติประกาศผล -> camp2_code ต้องมอบตามลำดับประกาศ (69DS02=69KM01, 69DS10=69KM02, 67DS15=69KM03, 68DS05=69KM04)
        approve_year_scores(2569, admin)
        for u, exp_km in zip([cur_low, cur_high, rig_old, rig_new], ['69KM01', '69KM02', '69KM03', '69KM04']):
            u.refresh_from_db()
            self.assertEqual(u.camp2_code, exp_km)

        # 6. Admin แก้ไขวันที่และหัวข้อรายชื่อแนบท้ายประกาศจากหน้าประกาศผลได้, Staff ที่ไม่ใช่ Admin ได้ 403
        r_forbid = sc.post(reverse('food:score_announcement'), {
            'action': 'save_announcement_text',
            'year': 2569,
            'announcement_date_text': 'วันที่ ๑๙ พฤศจิกายน พ.ศ. ๒๕๖๙',
            'announcement_list_title': 'หัวข้อทดสอบ',
        })
        self.assertEqual(r_forbid.status_code, 403)

        r_save = ac.post(reverse('food:score_announcement'), {
            'action': 'save_announcement_text',
            'year': 2569,
            'only': 'passed',
            'announcement_date_text': 'วันที่ ๑๙ พฤศจิกายน พ.ศ. ๒๕๖๙',
            'announcement_list_title': 'ประกาศรายชื่อผู้ผ่านคัดเลือกเข้าค่าย ๒ รอบพิเศษ',
        })
        self.assertEqual(r_save.status_code, 302)
        html_updated = ac.get(reverse('food:score_announcement') + '?year=2569').content.decode('utf-8')
        self.assertIn('ประกาศ ณ วันที่ 19 พฤศจิกายน พ.ศ. 2569', html_updated)
        self.assertIn('ประกาศรายชื่อผู้ผ่านคัดเลือกเข้าค่าย 2 รอบพิเศษ', html_updated)
        pdf_resp = ac.get(reverse('food:score_announcement') + '?year=2569&export=pdf')
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertTrue(pdf_resp.content.startswith(b'%PDF-'))

    def test_student_import_template_all_fields_except_camp2_and_shuffled_columns(self):
        """
        1. Template นำเข้านักเรียนมีทุกฟิลด์ที่จัดเก็บ ยกเว้น รหัสค่าย 2
        2. ไฟล์ที่นำเข้าสามารถสลับลำดับคอลัมน์ได้ (มีแถวหัวเรื่อง/คอลัมน์ส่วนเกิน/รหัสค่าย 2 ปนมาด้วย) โดยลง DB ถูกต้องตามหัวคอลัมน์
        3. ทดสอบทั้ง direct import_students, preview + confirm_import ผ่านหน้าเว็บ, การอัปเดตเฉพาะบางคอลัมน์, และ import_camp2 แบบสลับคอลัมน์
        """
        import openpyxl
        from django.core.files.uploadedfile import SimpleUploadedFile
        from .excel import STUDENT_COLUMNS, STUDENT_COLUMNS_TH, students_template_workbook

        # 1. ตรวจสอบ Template ภาษาไทยและภาษาอังกฤษ
        wb_th = students_template_workbook('th')
        th_headers = list(wb_th.worksheets[0].iter_rows(values_only=True))[0]
        self.assertEqual(list(th_headers), STUDENT_COLUMNS_TH)
        self.assertNotIn('รหัสค่าย 2', th_headers)
        self.assertNotIn('camp2_code', th_headers)
        for expected_th in (
            'รหัสค่าย 1', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'โรงเรียน', 'วันเกิด',
            'ปีการศึกษา', 'ระดับชั้นปัจจุบัน', 'เบอร์โทรนักเรียน', 'เบอร์โทรผู้ปกครอง',
            'อีเมล', 'โรคประจำตัว', 'แพ้อาหาร', 'ศาสนาอิสลาม', 'หมายเหตุ',
        ):
            self.assertIn(expected_th, th_headers)

        wb_en = students_template_workbook('en')
        en_headers = list(wb_en.worksheets[0].iter_rows(values_only=True))[0]
        self.assertEqual(list(en_headers), STUDENT_COLUMNS)
        self.assertNotIn('camp2_code', en_headers)

        # Round-trip import จาก Template ภาษาไทย
        buf_tpl = io.BytesIO()
        wb_th.save(buf_tpl)
        buf_tpl.seek(0)
        buf_tpl.name = 'template_all_fields.xlsx'
        res_tpl = import_students(buf_tpl)
        self.assertEqual(res_tpl['created'], 2)
        self.assertEqual(len(res_tpl['errors']), 0)
        u1 = User.objects.get(camp1_code='69DS01')
        self.assertEqual(u1.nickname, 'ต้น')
        self.assertEqual(u1.student_phone, '081-234-5678')
        self.assertEqual(u1.parent_phone, '089-876-5432')
        self.assertEqual(u1.email, 'somchai@example.com')
        self.assertEqual(u1.food_allergy, 'กุ้ง')
        self.assertFalse(u1.is_muslim)
        self.assertIsNone(u1.camp2_code)

        u2 = User.objects.get(camp1_code='69DS02')
        self.assertEqual(u2.nickname, 'หญิง')
        self.assertEqual(u2.chronic_disease, 'หอบหืด')
        self.assertTrue(u2.is_muslim)
        self.assertEqual(u2.note, 'ทานอาหารฮาลาล')

        # 2. ไฟล์ที่สลับลำดับคอลัมน์ทั้งหมด + แถวหัวเรื่อง + คอลัมน์ส่วนเกิน + รหัสค่าย 2 ที่ต้องถูกข้าม
        wb_shuf = Workbook()
        ws = wb_shuf.active
        ws.append(['รายชื่อนักเรียนค่าย 1 (เรียงคอลัมน์แบบไม่ตรง Template)'])
        ws.append([
            'ลำดับ', 'หมายเหตุ', 'อีเมล', 'รหัสค่าย 2', 'แพ้อาหาร', 'ศาสนาอิสลาม',
            'เบอร์โทรผู้ปกครอง', 'ชื่อเล่น', 'รหัสค่าย 1', 'ระดับชั้นปัจจุบัน',
            'โรคประจำตัว', 'วันเกิด', 'โรงเรียน', 'เบอร์โทรนักเรียน', 'ปีการศึกษา', 'ชื่อ-นามสกุล',
        ])
        ws.append([
            1, 'ต้องเตรียมยา', 'khao@example.com', '69KM99', 'ถั่วลิสง', 'อิสลาม',
            '0891112222', 'ข้าว', '69DS51', 'ม.4',
            'ภูมิแพ้', '10/05/2553', 'โรงเรียนเทพศิรินทร์', 812223333, 2569, 'นายข้าวหอม ใจดี',
        ])
        buf_shuf = io.BytesIO()
        wb_shuf.save(buf_shuf)
        buf_shuf.seek(0)
        buf_shuf.name = 'shuffled.xlsx'

        # ทดสอบ Preview -> Confirm Import ผ่าน Admin Web UI
        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        ac = Client()
        ac.force_login(admin)

        r_prev = ac.post(reverse('composn:admin_import_students'), {
            'action': 'preview',
            'file': SimpleUploadedFile(
                'shuffled.xlsx', buf_shuf.getvalue(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ),
        })
        self.assertEqual(r_prev.status_code, 200)
        prev_html = r_prev.content.decode('utf-8')
        self.assertIn('69DS51', prev_html)
        self.assertIn('khao@example.com', prev_html)
        self.assertIn('0812223333', prev_html)
        self.assertIn('ถั่วลิสง', prev_html)
        self.assertIn('ข้าม: ใช้เมนู Import รหัสค่าย 2', prev_html)

        r_conf = ac.post(reverse('composn:admin_import_students'), {'action': 'confirm_import'})
        self.assertEqual(r_conf.status_code, 200)

        u51 = User.objects.get(camp1_code='69DS51')
        self.assertEqual(u51.first_name, 'นายข้าวหอม')
        self.assertEqual(u51.last_name, 'ใจดี')
        self.assertEqual(u51.nickname, 'ข้าว')
        self.assertEqual(u51.school_name, 'โรงเรียนเทพศิรินทร์')
        self.assertEqual(u51.birth_date, date(2010, 5, 10))
        self.assertEqual(u51.academic_year, 2569)
        self.assertEqual(u51.current_level, 'ม.4')
        self.assertEqual(u51.year_plus_1, 'ม.5')
        self.assertEqual(u51.student_phone, '0812223333')
        self.assertEqual(u51.parent_phone, '0891112222')
        self.assertEqual(u51.email, 'khao@example.com')
        self.assertEqual(u51.chronic_disease, 'ภูมิแพ้')
        self.assertEqual(u51.food_allergy, 'ถั่วลิสง')
        self.assertTrue(u51.is_muslim)
        self.assertEqual(u51.note, 'ต้องเตรียมยา')
        self.assertIsNone(u51.camp2_code)  # ห้ามนำเข้ารหัสค่าย 2 ในหน้านำเข้านักเรียน

        # 3. อัปเดตเฉพาะบางคอลัมน์ (ไม่เรียงตาม Template) ต้องไม่ลบฟิลด์เดิมที่มีอยู่
        f_partial = make_xlsx(
            ['เบอร์โทรนักเรียน', 'รหัสค่าย 1', 'ชื่อ-นามสกุล'],
            [['0855556666', '69DS51', 'นายข้าวหอม ใจดี']],
        )
        res_part = import_students(f_partial)
        self.assertEqual(res_part['updated'], 1)
        u51.refresh_from_db()
        self.assertEqual(u51.student_phone, '0855556666')
        self.assertEqual(u51.nickname, 'ข้าว')
        self.assertEqual(u51.email, 'khao@example.com')
        self.assertEqual(u51.food_allergy, 'ถั่วลิสง')
        self.assertTrue(u51.is_muslim)

        # 4. import_camp2 แบบสลับคอลัมน์และมีคอลัมน์ส่วนเกิน
        f_c2 = make_xlsx(
            ['ชื่อ-นามสกุล', 'รหัสค่าย 2', 'รหัสค่าย 1'],
            [['นายข้าวหอม ใจดี', '69KM51', '69DS51']],
        )
        res_c2 = import_camp2(f_c2)
        self.assertEqual(res_c2['updated'], 1)
        u51.refresh_from_db()
        self.assertEqual(u51.camp2_code, '69KM51')

    def test_certificates_rename_grouping_and_info_tooltips(self):
        """
        1. เมนูและหน้าสร้างเกียรติบัตรเปลี่ยนชื่อเป็น 'เกียรติบัตร'
        2. หน้าเกียรติบัตรแสดงผลแยกเป็นกลุ่มและนำคอลัมน์สิทธิ์สอบท้ายค่ายออก
        3. หน้าต่าง ๆ แสดงผลไอคอน info-tip-icon (i) และเรนเดอร์ผ่านโดยไม่มีข้อผิดพลาด
        """
        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        ac = Client()
        ac.force_login(admin)

        m_curr = make_member('69DS71', year=2569, level='ม.4')
        m_curr.first_name, m_curr.last_name = 'ด.ช.ผ่าน', 'เกณฑ์'
        m_curr.save()

        m_rig = make_member('68DS72', year=2568, level='ม.5')
        m_rig.first_name, m_rig.last_name = 'นายใช้', 'สิทธิ์'
        m_rig.save()
        UseRecord.objects.create(
            user=m_rig, academic_year=2569, camp_level=1, rights_code='68DS72', is_used=True,
        )

        # กรองดูแต่ละกลุ่มบนหน้าเกียรติบัตร
        r_elig = ac.get(reverse('food:admin_certificates') + '?year=2569&status=eligible')
        self.assertEqual(r_elig.status_code, 200)
        h_elig = r_elig.content.decode('utf-8')
        self.assertIn('69DS71', h_elig)
        self.assertNotIn('68DS72', h_elig)
        self.assertNotIn('สิทธิ์สอบท้ายค่าย', h_elig)

        r_rig = ac.get(reverse('food:admin_certificates') + '?year=2569&status=rights')
        self.assertEqual(r_rig.status_code, 200)
        h_rig = r_rig.content.decode('utf-8')
        self.assertIn('68DS72', h_rig)
        self.assertNotIn('69DS71', h_rig)
        self.assertNotIn('สิทธิ์สอบท้ายค่าย', h_rig)

        # ตรวจเมนูใน Sidebar / Topbar และไอคอน (i) ในหน้าหลัก ๆ
        pages_with_info_tips = [
            reverse('composn:admin_dashboard'),
            reverse('composn:admin_users'),
            reverse('composn:admin_admins'),
            reverse('composn:admin_import_students'),
            reverse('composn:admin_import_camp2'),
            reverse('composn:admin_import_rights'),
            reverse('composn:admin_settings'),
            reverse('composn:admin_users_settings'),
            reverse('composn:admin_rounds'),
            reverse('composn:admin_calculate_level'),
            reverse('composn:admin_delete_camp2'),
            reverse('food:admin_certificates'),
            reverse('food:staff_scores'),
            reverse('food:score_summary'),
            reverse('food:score_categories'),
            reverse('food:score_announcement'),
            reverse('food:staff_attendance'),
            reverse('food:attendance_summary'),
            reverse('food:signin_sheet'),
            reverse('food:food_settings'),
            reverse('food:orders'),
            reverse('food:surveys_manage'),
        ]
        for url in pages_with_info_tips:
            resp = ac.get(url, follow=True)
            self.assertEqual(resp.status_code, 200, url)
            html = resp.content.decode('utf-8')
            self.assertNotIn('สร้างเกียรติบัตรออนไลน์', html, url)
            self.assertIn('info-tip-tooltip', html, url)
            if url != reverse('composn:admin_dashboard'):
                self.assertIn('info-tip-icon', html, url)
            if url == reverse('composn:admin_import_students'):
                self.assertNotIn(r'\n• ', html, url)
                self.assertIn('\n• ', html, url)
            if url == reverse('composn:admin_calculate_level'):
                self.assertNotIn(r'\nตัวอย่าง', html, url)

        mc = Client()
        mc.force_login(m_curr)
        for m_url in (
            reverse('composn:profile'),
            reverse('composn:training_scores_year', args=[2569]),
            reverse('food:my_food'),
            reverse('food:member_surveys'),
        ):
            r_m = mc.get(m_url)
            self.assertEqual(r_m.status_code, 200, m_url)
            self.assertIn('info-tip-icon', r_m.content.decode('utf-8'), m_url)
        self.assertRedirects(mc.get(reverse('food:my_selections')), reverse('food:my_food'))

    def test_member_sidebar_training_submenus_exam_status_certificate_and_food_merge(self):
        """
        ทดสอบการปรับปรุงเมนูและหน้าฝั่งสมาชิก:
        1. Sidebar แสดงรูปผู้ใช้ด้านบน 'รหัสค่าย 1', แสดง 'รหัสค่าย 2' ก่อนชื่อ-นามสกุล (เมื่อมี), นำ 'นักเรียน/ผู้ใช้สิทธิ์' มุมขวาบนออก
        2. เปลี่ยนเมนู 'ข้อมูลของฉัน' -> 'ข้อมูลส่วนตัว', เพิ่ม 'ข้อมูลการอบรม และคะแนน' ก่อน 'รายการอาหาร', เมนูย่อยตามปีการศึกษา
        3. สถานะ 'มีสิทธิ์สอบท้ายค่าย 1/2' แสดงต่อเมื่อประกาศผลการผ่านคัดเลือกแล้ว และแยกตามปีที่ได้รับรหัสค่าย 2
        4. ปุ่มดาวน์โหลดเกียรติบัตรย้ายมาอยู่หน้า 'ข้อมูลส่วนตัว' (ไม่อยู่ในหน้าข้อมูลการอบรมฯ)
        5. รวม 'รายการที่ฉันเลือก' ในหน้า 'รายการอาหาร', ตัด 'ในค่าย' ออกจาก 'แบบสอบถาม', แสดง badge ตัวเลขเมื่อยังไม่เลือก/ตอบ
        """
        import base64
        import tempfile
        from datetime import date, timedelta
        from decimal import Decimal
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test import override_settings
        from django.utils import timezone
        from food.models import (
            CampSurvey, CampSurveyQuestion, FoodMenu, FoodOrder, ScoreCategory,
            TrainingDate, approve_year_scores,
        )
        from food.services import save_score

        with tempfile.TemporaryDirectory() as tmpdir, override_settings(MEDIA_ROOT=tmpdir):
            self._run_member_sidebar_training_and_food_assertions(
                base64, timedelta, Decimal, SimpleUploadedFile, timezone,
                CampSurvey, CampSurveyQuestion, FoodMenu, FoodOrder, ScoreCategory,
                approve_year_scores, save_score,
            )

    def _run_member_sidebar_training_and_food_assertions(
        self, base64, timedelta, Decimal, SimpleUploadedFile, timezone,
        CampSurvey, CampSurveyQuestion, FoodMenu, FoodOrder, ScoreCategory,
        approve_year_scores, save_score,
    ):

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        setting = SystemSetting.get()
        setting.current_academic_year = 2569
        setting.passing_score = Decimal('50')
        setting.max_qualified_count = 0
        setting.food_system_open = True
        setting.save()

        tiny_png = base64.b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
        )
        # 1. นักเรียนปัจจุบัน 2 คน (ผ่านเกณฑ์ได้ค่าย 2 และไม่ผ่านได้สอบค่าย 1)
        m_pass = make_member('69DS81', year=2569, level='ม.4')
        m_pass.first_name, m_pass.last_name = 'นายผ่าน', 'คัดเลือก'
        m_pass.profile_photo.save('m_pass.png', SimpleUploadedFile('m_pass.png', tiny_png, content_type='image/png'), save=True)

        m_fail = make_member('69DS82', year=2569, level='ม.4')
        m_fail.first_name, m_fail.last_name = 'นายไม่ผ่าน', 'เกณฑ์'
        m_fail.save()

        # 2. นักเรียนรุ่น 2568: คนหนึ่งใช้สิทธิ์ปี 2569 อีกคนไม่ได้ใช้สิทธิ์ปี 2569
        m_rights_used = make_member('68DS83', year=2568, level='ม.3')
        m_rights_used.first_name, m_rights_used.last_name = 'นางสาวใช้', 'สิทธิ์'
        m_rights_used.save()
        UseRecord.objects.create(
            user=m_rights_used, academic_year=2569, camp_level=1,
            rights_code=m_rights_used.camp1_code, is_used=True,
        )

        m_rights_idle = make_member('68DS84', year=2568, level='ม.3')
        m_rights_idle.first_name, m_rights_idle.last_name = 'นายไม่ใช้', 'สิทธิ์'
        m_rights_idle.save()

        # คะแนนและผลปี 2568 สำหรับ m_rights_used (เคยไม่ผ่านค่าย 2 ในปี 2568)
        cat68 = ScoreCategory.objects.create(name='สอบรวม68', weight=Decimal('100'), academic_year=2568, is_visible=True)
        save_score(m_rights_used, cat68, 1, Decimal('30'), Decimal('100'), recorded_by=admin)
        approve_year_scores(2568, admin)

        # คะแนนปี 2569 (ยังไม่ประกาศผล)
        cat69 = ScoreCategory.objects.create(name='สอบรวม69', weight=Decimal('100'), academic_year=2569, is_visible=True)
        save_score(m_pass, cat69, 1, Decimal('80'), Decimal('100'), recorded_by=admin)
        save_score(m_fail, cat69, 1, Decimal('30'), Decimal('100'), recorded_by=admin)
        save_score(m_rights_used, cat69, 1, Decimal('75'), Decimal('100'), recorded_by=admin)

        # ตั้งรายการอาหารและแบบสอบถามที่ยังไม่ได้เลือก/ตอบสำหรับปี 2569
        today = timezone.localdate()
        food_ord = FoodOrder.objects.create(
            academic_year=2569, food_date=today + timedelta(days=1),
            restaurant_name='ร้านทดสอบเมนูรวม', max_selection_count=1, edit_toggle_open=True,
            allow_current_students=True, allow_rights_students=True,
        )
        food_menu = FoodMenu.objects.create(food_order=food_ord, menu_name='ข้าวหมูกระเทียม')

        survey = CampSurvey.objects.create(academic_year=2569, title='แบบสำรวจความพึงพอใจ', is_open=True, created_by=admin)
        CampSurveyQuestion.objects.create(
            survey=survey, question_text='ระดับความพึงพอใจ',
            question_type=CampSurveyQuestion.TYPE_CHOICE, choices_text='มาก, ปานกลาง', sort_order=1,
        )

        # ตรวจก่อนประกาศผล 2569
        mc_pass = Client()
        mc_pass.force_login(m_pass)
        prof_before = mc_pass.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn(m_pass.profile_photo.url, prof_before)
        self.assertLess(prof_before.index(m_pass.profile_photo.url), prof_before.index('รหัสค่าย 1 : 69DS81'))
        self.assertNotIn('นักเรียน/ผู้ใช้สิทธิ์', prof_before)
        self.assertIn('ข้อมูลส่วนตัว', prof_before)
        self.assertNotIn('ข้อมูลของฉัน', prof_before)
        self.assertIn('ข้อมูลการอบรม', prof_before)
        self.assertNotIn('ข้อมูลการอบรม และคะแนน', prof_before)
        self.assertLess(prof_before.index('ข้อมูลการอบรม'), prof_before.index('รายการอาหาร'))
        self.assertIn('ปีการศึกษา 2569', prof_before)
        self.assertNotIn('แบบสอบถามในค่าย', prof_before)
        self.assertNotIn(f'href="{reverse("food:my_selections")}"', prof_before)
        # ค่าเริ่มต้นของเกียรติบัตรคือปิด: ก่อน Admin เปิดแสดง ต้องยังไม่เห็นปุ่มดาวน์โหลดเกียรติบัตร
        self.assertNotIn('ดาวน์โหลดเกียรติบัตร (PDF)', prof_before)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 1', prof_before)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 2', prof_before)

        # เมื่อ Admin เปิดแสดงเกียรติบัตรปี 2569 -> ปุ่มดาวน์โหลดเกียรติบัตรแสดงในหน้าข้อมูลส่วนตัว
        ac_pub = Client()
        ac_pub.force_login(admin)
        ac_pub.post(reverse('food:admin_certificates'), {
            'action': 'toggle_cert_publish',
            'year': '2569',
            'tab': 'students',
        })
        prof_pub = mc_pass.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('ดาวน์โหลดเกียรติบัตร (PDF)', prof_pub)

        # เมนูรายการอาหาร และแบบสอบถาม ต้องขึ้นตัวเลข badge แจ้งเตือนทั้งสองเมนู
        sidebar_before = prof_before[prof_before.index('id="sidebar"'):prof_before.index('class="main-content"')]
        food_nav_before = sidebar_before[sidebar_before.index('รายการอาหาร'):sidebar_before.index('แบบสอบถาม')]
        survey_nav_before = sidebar_before[sidebar_before.index('แบบสอบถาม'):sidebar_before.index('แสดงรายชื่อผู้ใช้สิทธิ์')]
        self.assertIn('>1<', food_nav_before)
        self.assertIn('>1<', survey_nav_before)

        train_before = mc_pass.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('ข้อมูลการอบรม', train_before)
        self.assertNotIn('ข้อมูลการอบรม และคะแนน', train_before)
        self.assertNotIn('ผลการอบรมและคะแนน', train_before)
        self.assertNotIn('ดาวน์โหลดเกียรติบัตร', train_before)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 1', train_before)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 2', train_before)

        # นักเรียนรุ่น 2568 ที่ไม่ได้ใช้สิทธิ์ในปี 2569 ต้องไม่มีเมนูย่อยปี 2569 หรือ 9999
        mc_idle = Client()
        mc_idle.force_login(m_rights_idle)
        prof_idle = mc_idle.get(reverse('composn:profile')).content.decode('utf-8')
        sidebar_idle = prof_idle[prof_idle.index('id="grp-member-training"'):prof_idle.index('รายการอาหาร')]
        self.assertIn('ปีการศึกษา 2568', sidebar_idle)
        self.assertNotIn('ปีการศึกษา 2569', sidebar_idle)
        self.assertNotIn('ปีการศึกษา 9999', sidebar_idle)
        resp_9999 = mc_idle.get(reverse('composn:training_scores_year', args=[9999]))
        self.assertEqual(resp_9999.status_code, 302)

        # ประกาศผลปี 2569 -> m_pass ได้รับรหัสค่าย 2, m_rights_used ได้รับรหัสค่าย 2, m_fail ไม่ได้รับ
        approve_year_scores(2569, admin)
        m_pass.refresh_from_db()
        m_fail.refresh_from_db()
        m_rights_used.refresh_from_db()
        self.assertIsNotNone(m_pass.camp2_code)
        self.assertIsNone(m_fail.camp2_code)
        self.assertIsNotNone(m_rights_used.camp2_code)

        # Sidebar ของผู้มีรหัสค่าย 2 ต้องแสดง รหัสค่าย 1 และบรรทัดถัดไป รหัสค่าย 2 ก่อนชื่อ-นามสกุล
        prof_pass_after = mc_pass.get(reverse('composn:profile')).content.decode('utf-8')
        brand_codes_chunk = prof_pass_after[prof_pass_after.index('รหัสค่าย 1 : 69DS81'):prof_pass_after.index('เมนูสมาชิก')]
        self.assertIn(f'รหัสค่าย 2 : {m_pass.camp2_code}', brand_codes_chunk)
        self.assertLess(brand_codes_chunk.index('รหัสค่าย 1 : 69DS81'), brand_codes_chunk.index(f'รหัสค่าย 2 : {m_pass.camp2_code}'))
        self.assertLess(brand_codes_chunk.index(f'รหัสค่าย 2 : {m_pass.camp2_code}'), brand_codes_chunk.index('นายผ่าน คัดเลือก'))
        self.assertIn('มีสิทธิ์สอบท้ายค่าย 2', prof_pass_after)

        train_pass_after = mc_pass.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('มีสิทธิ์สอบท้ายค่าย 2', train_pass_after)

        mc_fail = Client()
        mc_fail.force_login(m_fail)
        train_fail_after = mc_fail.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('มีสิทธิ์สอบท้ายค่าย 1', train_fail_after)

        # นักเรียนใช้สิทธิ์ m_rights_used (เคยได้ค่าย 1 ในปี 2568 และได้รหัสค่าย 2 ในปี 2569):
        # ปี 2568 ต้องแสดง "มีสิทธิ์สอบท้ายค่าย 1" ส่วนปี 2569 แสดง "มีสิทธิ์สอบท้ายค่าย 2"
        mc_ru = Client()
        mc_ru.force_login(m_rights_used)
        tr_68 = mc_ru.get(reverse('composn:training_scores_year', args=[2568])).content.decode('utf-8')
        tr_69 = mc_ru.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('มีสิทธิ์สอบท้ายค่าย 1', tr_68)
        self.assertNotIn('มีสิทธิ์สอบท้ายค่าย 2', tr_68)
        self.assertIn('มีสิทธิ์สอบท้ายค่าย 2', tr_69)

        # เลือกอาหาร + ตอบแบบสอบถาม -> รายการที่ฉันเลือกแสดงบนหน้าเดียวกับรายการอาหาร และ badge หายไป
        mc_pass.post(reverse('food:my_food_detail', args=[food_ord.pk]), {'action': 'select', 'menu_id': food_menu.pk})
        q_sv = survey.questions.first()
        mc_pass.post(reverse('food:member_surveys'), {'survey_id': survey.pk, f'q_{q_sv.pk}': 'มาก'})
        my_food_after = mc_pass.get(reverse('food:my_food')).content.decode('utf-8')
        self.assertIn('id="my-selections-section"', my_food_after)
        self.assertIn('ข้าวหมูกระเทียม', my_food_after[my_food_after.index('id="my-selections-section"'):])
        sidebar_after = my_food_after[my_food_after.index('id="sidebar"'):my_food_after.index('class="main-content"')]
        nav_food_done = sidebar_after[sidebar_after.index('รายการอาหาร'):sidebar_after.index('แสดงรายชื่อผู้ใช้สิทธิ์')]
        self.assertNotIn('>1<', nav_food_done)

    def test_admin_menus_merged_score_excel_center_logo_and_customizable_dashboard(self):
        """
        ทดสอบการปรับปรุงฝั่ง Admin:
        1. รูปศูนย์อบรม (logo_center) อยู่เหนือ DEBSIRIN POSN COMPUTER
        2. รวมเมนู สรุปคะแนนรวม + บันทึกคะแนน, เกียรติบัตร + เอกสารประกาศผล, สรุปเวลาเรียน + แบบลงชื่อ + บันทึกเวลาเรียน
        3. แก้ชื่อเมนู อาหาร ('รายการอาหาร', 'แบบสอบถาม') และเรียงชื่อเมนูย่อย 5. ตั้งค่าระบบ ครบทั้ง 6 รายการตามลำดับ
        4. Export Excel คะแนนรวมจากหน้าบันทึกคะแนน/สรุปคะแนนรวม ออกมาทุกช่องทั้งหัวข้อใหญ่และหัวข้อย่อย
        5. แดชบอร์ดสามารถเลือกเปิด/ปิดส่วนแสดงผล จัดลำดับขึ้น-ลง เลือกฟิลด์สรุปนักเรียน และคืนค่าเริ่มต้นได้
        """
        import base64
        import openpyxl
        import tempfile
        from decimal import Decimal
        from django.core.files.uploadedfile import SimpleUploadedFile
        from django.test import override_settings
        from food.models import ScoreCategory
        from food.services import save_score

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        tiny_png = base64.b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
        )
        setting = SystemSetting.get()
        setting.site_title = 'DEBSIRIN POSN COMPUTER'
        setting.current_academic_year = 2569

        with tempfile.TemporaryDirectory() as tmpdir, override_settings(MEDIA_ROOT=tmpdir):
            setting.logo_center.save(
                'center_logo.png', SimpleUploadedFile('center_logo.png', tiny_png, content_type='image/png'), save=True,
            )

            # เตรียมข้อมูลคะแนนทั้งหัวข้อใหญ่ที่มีหัวข้อย่อย และหัวข้อใหญ่ที่ซ่อน (is_visible=False)
            m1 = make_member('69DS95', year=2569, level='ม.4', first_name='ธนากร', last_name='ทดสอบ')
            cat_main = ScoreCategory.objects.create(
                academic_year=2569, name='ทฤษฎีรวม', weight=Decimal('60'), default_max_score=Decimal('100'),
                order=1, is_visible=True,
            )
            sub1 = ScoreCategory.objects.create(
                academic_year=2569, parent=cat_main, name='ข้อ 1', weight=Decimal('0'), default_max_score=Decimal('50'),
                order=1, is_visible=True,
            )
            sub2 = ScoreCategory.objects.create(
                academic_year=2569, parent=cat_main, name='ข้อ 2', weight=Decimal('0'), default_max_score=Decimal('50'),
                order=2, is_visible=True,
            )
            cat_hidden = ScoreCategory.objects.create(
                academic_year=2569, name='ปฏิบัติลับ', weight=Decimal('40'), default_max_score=Decimal('100'),
                order=2, is_visible=False,
            )
            save_score(m1, sub1, 1, Decimal('40'), Decimal('50'), recorded_by=admin)
            save_score(m1, sub2, 1, Decimal('30'), Decimal('50'), recorded_by=admin)
            save_score(m1, cat_hidden, 1, Decimal('80'), Decimal('100'), recorded_by=admin)

            ac = Client()
            ac.force_login(admin)

            # 1. ตรวจรูปศูนย์อบรม และเมนูบน Sidebar & Topbar
            dash_html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
            sidebar_html = dash_html[dash_html.index('id="sidebar"'):dash_html.index('class="main-content"')]
            self.assertIn(setting.logo_center.url, sidebar_html)
            self.assertLess(sidebar_html.index(setting.logo_center.url), sidebar_html.index('DEBSIRIN POSN COMPUTER'))

            scores_grp = sidebar_html[sidebar_html.index('id="grp-scores"'):sidebar_html.index('id="grp-food"')]
            for lbl in ('บันทึกเวลาเรียน', 'บันทึกคะแนน', 'จัดที่นั่งสอบ', 'เกียรติบัตร', 'การประกาศผล'):
                self.assertIn(lbl, scores_grp)
            self.assertLess(scores_grp.index('จัดที่นั่งสอบ'), scores_grp.index('เกียรติบัตร'))
            self.assertLess(scores_grp.index('เกียรติบัตร'), scores_grp.index('การประกาศผล'))
            self.assertNotIn('สรุปคะแนนรวม', scores_grp)
            self.assertNotIn('กรอกคะแนน', scores_grp)
            self.assertNotIn('สรุปเวลาเรียน', scores_grp)
            self.assertNotIn('แบบลงชื่อ', scores_grp)

            food_grp = sidebar_html[sidebar_html.index('id="grp-food"'):sidebar_html.index('id="grp-rights"')]
            self.assertIn('รายการอาหาร', food_grp)
            self.assertIn('แบบสอบถาม', food_grp)
            self.assertNotIn('สร้างรายการอาหาร', food_grp)
            self.assertNotIn('สรุปอาหาร', food_grp)

            settings_grp = sidebar_html[sidebar_html.index('id="grp-settings"'):sidebar_html.index('บัญชี')]
            expected_settings = ['ผู้ใช้', 'ตั้งค่าสี ข้อความ', 'กำหนดปีการศึกษา', 'รายการอาหาร', 'บันทึกสำรอง']
            pos = -1
            for lbl in expected_settings:
                nxt = settings_grp.index(lbl, pos + 1)
                self.assertGreater(nxt, pos, lbl)
                pos = nxt
            self.assertNotIn('เกียรติบัตร', settings_grp)
            self.assertNotIn('วันอบรม และคะแนน', settings_grp)
            self.assertNotIn('ตั้งค่าหน้าผู้ใช้', settings_grp)
            self.assertNotIn('ไฟล์สำรอง / Audit', settings_grp)

            # 2. ตรวจแท็บในหน้าย่อยที่นำมารวมกัน
            sc_html = ac.get(reverse('food:staff_scores') + '?year=2569').content.decode('utf-8')
            self.assertIn('บันทึกคะแนน', sc_html)
            self.assertIn('สรุปคะแนนรวม', sc_html)
            self.assertIn('export=all_xlsx', sc_html)

            ann_html = ac.get(reverse('food:score_announcement') + '?year=2569').content.decode('utf-8')
            self.assertIn('เอกสารประกาศผล', ann_html)
            self.assertIn(reverse('food:admin_certificates'), ann_html)

            att_html = ac.get(reverse('food:staff_attendance') + '?year=2569').content.decode('utf-8')
            self.assertIn(reverse('food:attendance_summary'), att_html)
            self.assertIn(reverse('food:signin_sheet'), att_html)

            bak_html = ac.get(reverse('composn:admin_backups')).content.decode('utf-8')
            self.assertIn(reverse('composn:admin_exports'), bak_html)
            self.assertIn(reverse('composn:admin_audit_logs'), bak_html)
            self.assertIn(reverse('composn:admin_trash'), bak_html)

            # 3. ตรวจ Export Excel คะแนนทุกหัวข้อใหญ่และหัวข้อย่อย (ทั้งจากหน้าบันทึกคะแนนและสรุปคะแนนรวม)
            for exp_url in (
                reverse('food:staff_scores') + '?year=2569&export=all_xlsx',
                reverse('food:score_summary') + '?year=2569&export=xlsx',
            ):
                r_xlsx = ac.get(exp_url)
                self.assertEqual(r_xlsx.status_code, 200, exp_url)
                wb = openpyxl.load_workbook(io.BytesIO(r_xlsx.content))
                ws = wb.active
                headers = [c.value for c in ws[1]]
                header_str = ' | '.join(str(h or '') for h in headers)
                self.assertIn('1.1 ข้อ 1', header_str)
                self.assertIn('1.2 ข้อ 2', header_str)
                self.assertIn('1. รวม ทฤษฎีรวม', header_str)
                self.assertIn('2. ปฏิบัติลับ', header_str)
                self.assertIn('คะแนนรวม (100.00)', header_str)
                row1 = [c.value for c in ws[2]]
                self.assertEqual(row1[1], '69DS95')
                # (40+30)/100 * 60 = 42, 80/100 * 40 = 32 -> รวม = 74
                self.assertEqual(float(row1[headers.index('คะแนนรวม (100.00)')]), 74.0)

            # 4. ทดสอบปรับแต่งแดชบอร์ด: เลือกเฉพาะส่วนที่ต้องการ + จัดลำดับ + เพิ่ม summary field ใหม่ + คืนค่าเริ่มต้น
            r_save_dash = ac.post(reverse('composn:admin_dashboard'), {
                'action': 'save_dashboard_config',
                'year': '2569',
                'dashboard_widgets': ['recent_logs', 'student_summary', 'scores_certificates'],
                'dashboard_summary_fields': ['camp2_code', 'profile_photo'],
            }, follow=True)
            self.assertEqual(r_save_dash.status_code, 200)
            custom_html = r_save_dash.content.decode('utf-8')
            self.assertIn('id="widget-recent_logs"', custom_html)
            self.assertIn('id="widget-student_summary"', custom_html)
            self.assertIn('id="widget-scores_certificates"', custom_html)
            self.assertNotIn('id="widget-kpi_overview"', custom_html)
            self.assertNotIn('id="widget-food_orders"', custom_html)
            self.assertLess(custom_html.index('id="widget-recent_logs"'), custom_html.index('id="widget-student_summary"'))
            self.assertLess(custom_html.index('id="widget-student_summary"'), custom_html.index('id="widget-scores_certificates"'))
            self.assertIn('ได้รับรหัสค่าย 2 แล้ว', custom_html)
            self.assertIn('อัปโหลดรูปประจำตัวแล้ว', custom_html)

            # คืนค่าเริ่มต้น
            r_reset_dash = ac.post(reverse('composn:admin_dashboard'), {
                'action': 'reset_dashboard_config',
                'year': '2569',
            }, follow=True)
            self.assertEqual(r_reset_dash.status_code, 200)
            reset_html = r_reset_dash.content.decode('utf-8')
            self.assertIn('id="widget-kpi_overview"', reset_html)
            self.assertIn('id="widget-attendance_summary"', reset_html)
            self.assertIn('id="widget-food_orders"', reset_html)

    def test_profile_status_row_attendance_labels_and_per_year_cert_publish(self):
        """
        ทดสอบ:
        1. หน้าข้อมูลส่วนตัว (#profile-rights-box) แสดงแถว สถานะ (อบรม / ใช้สิทธิ์ / -) และลิงก์ไปยังข้อมูลการอบรม
        2. ตัดคำว่า 'และคะแนน' ออกจากเมนู/หน้า 'ข้อมูลการอบรม'
        3. สรุปเวลาเรียนในหน้าข้อมูลการอบรมแสดง 'ขาดอบรม X ภาคเวลา จาก Y ภาคเวลา' และ 'ขาดสอบ Z ครั้ง' (ไม่แสดง 'มา xx')
        4. การเปิด/ปิดแสดงเกียรติบัตรแยกตามปีการศึกษาโดย Admin (ค่าเริ่มต้น = ปิด, Staff toggle ไม่ได้ 403)
        """
        from datetime import date
        from food.models import Attendance, CertificateYearStatus, TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        staff = User(
            camp1_code='staff_cert_test', first_name='สต๊าฟ', last_name='ทดสอบ',
            role=User.ROLE_STAFF, staff_can_scores=True, is_first_login=False,
        )
        staff.set_password('staff_cert_test')
        staff.save()

        setting = SystemSetting.get()
        setting.current_academic_year = 2569
        setting.max_absences = 5
        setting.save()

        # นักเรียนรุ่น 2568 ใช้สิทธิ์ในปี 2569 + นักเรียนรุ่น 2569 เข้าอบรม
        m68 = make_member('68DS91', year=2568, level='ม.4', first_name='อดีต', last_name='ใช้สิทธิ์')
        UseRecord.objects.create(
            user=m68, academic_year=2569, camp_level=1, rights_code='68DS91', is_used=True,
        )
        m69 = make_member('69DS92', year=2569, level='ม.4', first_name='ปัจจุบัน', last_name='เข้าอบรม')

        # สร้างวันอบรม 2 วัน (4 ภาคเวลา: มา 3 ขาด 1) + วันสอบ 1 วัน (เช้าเข้าสอบ, บ่ายขาดสอบ -> ขาดสอบ 1 ครั้ง)
        TrainingDate.objects.create(academic_year=2569, date=date(2026, 10, 1), event_type=TrainingDate.TYPE_TRAINING)
        TrainingDate.objects.create(academic_year=2569, date=date(2026, 10, 2), event_type=TrainingDate.TYPE_TRAINING)
        TrainingDate.objects.create(
            academic_year=2569, date=date(2026, 10, 3), event_type=TrainingDate.TYPE_EXAM, exam_period=TrainingDate.EXAM_BOTH,
        )
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 1), session_period='morning', status=Attendance.STATUS_PRESENT)
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 1), session_period='afternoon', status=Attendance.STATUS_PRESENT)
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 2), session_period='morning', status=Attendance.STATUS_PRESENT)
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 2), session_period='afternoon', status=Attendance.STATUS_ABSENT)
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 3), session_period='morning', status=Attendance.STATUS_PRESENT)
        Attendance.objects.create(user=m69, academic_year=2569, date=date(2026, 10, 3), session_period='afternoon', status=Attendance.STATUS_ABSENT)

        # m68 รุ่น 2568: ไม่มีการขาดอบรมและไม่มีวันสอบ -> ขาดสอบ 0 ครั้ง
        TrainingDate.objects.create(academic_year=2568, date=date(2025, 10, 1), event_type=TrainingDate.TYPE_TRAINING)
        Attendance.objects.create(user=m68, academic_year=2568, date=date(2025, 10, 1), session_period='morning', status=Attendance.STATUS_PRESENT)
        Attendance.objects.create(user=m68, academic_year=2568, date=date(2025, 10, 1), session_period='afternoon', status=Attendance.STATUS_PRESENT)

        # 1. ตรวจแถว สถานะ บน #profile-rights-box และชื่อเมนู 'ข้อมูลการอบรม'
        mc68 = Client()
        mc68.force_login(m68)
        prof68 = mc68.get(reverse('composn:profile')).content.decode('utf-8')
        rights_box = prof68[prof68.index('id="profile-rights-box"'):prof68.index('class="col-lg-8"')]
        self.assertIn('สถานะ', rights_box)
        self.assertIn('อบรม', rights_box)
        self.assertIn('ใช้สิทธิ์', rights_box)
        self.assertIn(reverse('composn:training_scores_year', args=[2568]), rights_box)
        self.assertIn(reverse('composn:training_scores_year', args=[2569]), rights_box)
        self.assertIn('ข้อมูลการอบรม', prof68)
        self.assertNotIn('ข้อมูลการอบรม และคะแนน', prof68)

        # 2. ตรวจเวลาเรียนฝั่งสมาชิก: m69 ขาดอบรม 1 จาก 4 ภาคเวลา + ขาดสอบ 1 ครั้ง และ m68 ขาดสอบ 0 ครั้ง (ไม่แสดง 'มา xx')
        mc69 = Client()
        mc69.force_login(m69)
        tr69 = mc69.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('เข้าอบรม 75%', tr69)
        self.assertIn('ขาดอบรม 1 ภาคเวลา จาก 4 ภาคเวลา', tr69)
        self.assertIn('ขาดสอบ 1 ครั้ง', tr69)
        self.assertNotIn('>มา 3<', tr69)
        self.assertNotIn('ข้อมูลการอบรม และคะแนน', tr69)

        tr68 = mc68.get(reverse('composn:training_scores_year', args=[2568])).content.decode('utf-8')
        self.assertIn('เข้าอบรม 100%', tr68)
        self.assertIn('ขาดอบรม 0 ภาคเวลา จาก 2 ภาคเวลา', tr68)
        self.assertIn('ขาดสอบ 0 ครั้ง', tr68)
        self.assertNotIn('>มา 2<', tr68)

        # 3. ตรวจสถานะเปิด/ปิดเกียรติบัตรแยกตามปี (ค่าเริ่มต้น ปิด ทั้ง 2568 และ 2569)
        self.assertFalse(CertificateYearStatus.is_year_published(2568))
        self.assertFalse(CertificateYearStatus.is_year_published(2569))

        # บ่ายสอบให้ m69 เข้าสอบครบเพื่อให้ m69 ผ่านเกณฑ์เกียรติบัตร 2569
        att_exam_pm = Attendance.objects.get(user=m69, academic_year=2569, date=date(2026, 10, 3), session_period='afternoon')
        att_exam_pm.status = Attendance.STATUS_PRESENT
        att_exam_pm.save()
        tr69_zero = mc69.get(reverse('composn:training_scores_year', args=[2569])).content.decode('utf-8')
        self.assertIn('ขาดสอบ 0 ครั้ง', tr69_zero)

        # ก่อนเปิดแสดงเกียรติบัตร ทั้ง 2568 และ 2569 ต้องไม่เห็นกล่องเกียรติบัตรและโหลดไม่ได้
        prof69_closed = mc69.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('id="profile-certificate-box"', prof69_closed)
        self.assertEqual(mc69.get(reverse('food:my_certificate') + '?year=2569').status_code, 302)
        self.assertNotIn('id="profile-certificate-box"', prof68)
        self.assertEqual(mc68.get(reverse('food:my_certificate') + '?year=2568').status_code, 302)

        # Staff ไม่มีสิทธิ์เปิด/ปิดการแสดงเกียรติบัตร (403)
        sc = Client()
        sc.force_login(staff)
        self.assertEqual(sc.post(reverse('food:admin_certificates'), {
            'action': 'toggle_cert_publish', 'year': '2569', 'tab': 'students',
        }).status_code, 403)

        # Admin เปิดแสดงเกียรติบัตรเฉพาะปี 2569 -> m69 เห็นและโหลด PDF ได้ ส่วนปี 2568 ยังปิดอยู่
        ac = Client()
        ac.force_login(admin)
        cert_page_before = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=students').content.decode('utf-8')
        self.assertIn('id="cert-publish-bar"', cert_page_before)
        self.assertIn('ปิดการแสดง (ค่าเริ่มต้น)', cert_page_before)

        ac.post(reverse('food:admin_certificates'), {
            'action': 'toggle_cert_publish', 'year': '2569', 'tab': 'students',
        })
        self.assertTrue(CertificateYearStatus.is_year_published(2569))
        self.assertFalse(CertificateYearStatus.is_year_published(2568))

        prof69_open = mc69.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('id="profile-certificate-box"', prof69_open)
        self.assertIn('ดาวน์โหลดเกียรติบัตร (PDF)', prof69_open)
        resp69_pdf = mc69.get(reverse('food:my_certificate') + '?year=2569')
        self.assertEqual(resp69_pdf.status_code, 200)
        self.assertEqual(resp69_pdf['Content-Type'], 'application/pdf')

        # ปี 2568 ยังปิดอยู่ -> m68 ยังไม่เห็นและโหลดไม่ได้
        prof68_still_closed = mc68.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('id="profile-certificate-box"', prof68_still_closed)
        self.assertEqual(mc68.get(reverse('food:my_certificate') + '?year=2568').status_code, 302)

        # Admin ปิดแสดงเกียรติบัตรปี 2569 กลับ -> m69 ต้องกลับมาถูกซ่อนและบล็อก
        ac.post(reverse('food:admin_certificates'), {
            'action': 'toggle_cert_publish', 'year': '2569', 'tab': 'settings',
        })
        self.assertFalse(CertificateYearStatus.is_year_published(2569))
        prof69_reclosed = mc69.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('id="profile-certificate-box"', prof69_reclosed)
        self.assertEqual(mc69.get(reverse('food:my_certificate') + '?year=2569').status_code, 302)

    def test_all_logos_transparent_on_upload_templates_and_pdf(self):
        """
        ทดสอบการทำให้รูปตราต่างๆ มีพื้นหลังโปร่งใสทั้งหมด:
        1. อัปโหลด JPEG พื้นขาว (ภาพขาว-ดำตราครุฑ และภาพสีมีกรอบปิดพร้อมสีขาวด้านใน) -> แปลงเป็น PNG พื้นโปร่งใสอัตโนมัติโดยคงสีขาวด้านในไว้
        2. อัปโหลด PNG ที่โปร่งใสอยู่แล้ว -> คงความโปร่งใสและรายละเอียดเดิมไว้
        3. ปุ่ม action='make_logos_transparent' -> แปลงรูปทึบที่เคยบันทึกไว้ให้โปร่งใสได้
        4. หน้าเว็บ (Sidebar, ตั้งค่าระบบ, Login, ใบลงชื่อ, ประกาศผล) และ PDF (ใบลงชื่อ, ประกาศผล, เกียรติบัตร) ใช้พื้นหลังโปร่งใสทั้งหมด
        """
        import tempfile
        from PIL import Image, ImageDraw
        from django.core.files.base import ContentFile
        from django.core.files.uploadedfile import SimpleUploadedFile
        from food.announcement_pdf import build_announcement_pdf
        from food.certificate_pdf import build_certificate_pdf
        from food.signin_pdf import _load_logo_reader

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        ac = Client()
        ac.force_login(admin)

        m_pdf = make_member('69DS99', year=2569, level='ม.4', first_name='โลโก้', last_name='โปร่งใส')

        # 1) ภาพขาว-ดำพื้นขาว (เลียนแบบตราครุฑ JPEG)
        garuda_img = Image.new('L', (60, 60), 255)
        ImageDraw.Draw(garuda_img).rectangle([18, 18, 42, 42], fill=0)
        garuda_buf = io.BytesIO()
        garuda_img.save(garuda_buf, format='JPEG')

        # 2) ภาพสีพื้นขาว มีกรอบปิดทึบสีน้ำเงินและพื้นที่สีขาวอยู่ด้านใน (ต้องไม่ลบสีขาวด้านใน)
        center_img = Image.new('RGB', (80, 60), (255, 255, 255))
        d_c = ImageDraw.Draw(center_img)
        d_c.rectangle([12, 10, 68, 50], fill=(20, 60, 180))
        d_c.rectangle([26, 20, 54, 40], fill=(255, 255, 255))
        center_buf = io.BytesIO()
        center_img.save(center_buf, format='JPEG')

        # 3) PNG โปร่งใสอยู่แล้ว พร้อมสีขาวด้านใน
        school_img = Image.new('RGBA', (60, 60), (0, 0, 0, 0))
        d_s = ImageDraw.Draw(school_img)
        d_s.ellipse([10, 10, 50, 50], fill=(0, 80, 160, 255))
        d_s.rectangle([24, 24, 36, 36], fill=(255, 255, 255, 255))
        school_buf = io.BytesIO()
        school_img.save(school_buf, format='PNG')

        with tempfile.TemporaryDirectory() as tmpdir, override_settings(MEDIA_ROOT=tmpdir):
            resp = ac.post(reverse('composn:admin_settings'), {
                'action': 'save_all_settings',
                'current_academic_year': '2569',
                'max_absences': '5',
                'auto_logout_minutes': '60',
                'logo_garuda': SimpleUploadedFile('garuda_in.jpg', garuda_buf.getvalue(), content_type='image/jpeg'),
                'logo_center': SimpleUploadedFile('center_in.jpg', center_buf.getvalue(), content_type='image/jpeg'),
                'logo_school': SimpleUploadedFile('school_in.png', school_buf.getvalue(), content_type='image/png'),
            })
            self.assertEqual(resp.status_code, 302)

            s = SystemSetting.get()
            self.assertTrue(s.logo_garuda.name.lower().endswith('.png'))
            self.assertTrue(s.logo_center.name.lower().endswith('.png'))
            self.assertTrue(s.logo_school.name.lower().endswith('.png'))

            with Image.open(s.logo_garuda.path) as ig:
                self.assertEqual(ig.mode, 'RGBA')
                self.assertEqual(ig.getpixel((0, 0))[3], 0)
                self.assertGreater(ig.getpixel((30, 30))[3], 240)

            with Image.open(s.logo_center.path) as ic:
                self.assertEqual(ic.mode, 'RGBA')
                self.assertEqual(ic.getpixel((0, 0))[3], 0)
                self.assertGreater(ic.getpixel((15, 15))[3], 240)
                interior_px = ic.getpixel((40, 30))
                self.assertEqual(interior_px[3], 255)
                self.assertGreaterEqual(min(interior_px[:3]), 250)

            with Image.open(s.logo_school.path) as isc:
                self.assertEqual(isc.mode, 'RGBA')
                self.assertEqual(isc.getpixel((0, 0))[3], 0)
                self.assertEqual(isc.getpixel((30, 30)), (255, 255, 255, 255))

            # จำลองรูปทึบเก่าใน DB แล้วกดปุ่ม 'make_logos_transparent'
            obec_img = Image.new('RGB', (50, 50), (255, 255, 255))
            ImageDraw.Draw(obec_img).ellipse([12, 12, 38, 38], fill=(200, 90, 0))
            obec_buf = io.BytesIO()
            obec_img.save(obec_buf, format='JPEG')
            s.logo_obec.save('obec_legacy.jpg', ContentFile(obec_buf.getvalue()), save=True)

            r_conv = ac.post(reverse('composn:admin_settings'), {'action': 'make_logos_transparent'})
            self.assertEqual(r_conv.status_code, 302)
            s.refresh_from_db()
            self.assertTrue(s.logo_obec.name.lower().endswith('.png'))
            with Image.open(s.logo_obec.path) as io_img:
                self.assertEqual(io_img.mode, 'RGBA')
                self.assertEqual(io_img.getpixel((0, 0))[3], 0)
                self.assertGreater(io_img.getpixel((25, 25))[3], 240)

            # ตรวจ HTML หน้าตั้งค่า / Sidebar / Login / ใบลงชื่อ / ประกาศผล
            set_html = ac.get(reverse('composn:admin_settings')).content.decode('utf-8')
            self.assertIn('ปรับพื้นหลังโปร่งใสทั้งหมด', set_html)
            self.assertIn('bg-transparent', set_html)
            self.assertNotIn('border rounded p-2 bg-light', set_html)

            sidebar_html = set_html[set_html.index('id="sidebar"'):set_html.index('class="main-content"')]
            self.assertIn('alt="รูปศูนย์อบรม"', sidebar_html)
            self.assertIn('background:transparent', sidebar_html)
            self.assertNotIn('bg-white', sidebar_html[:sidebar_html.index('<h5')])

            login_html = Client().get(reverse('composn:login')).content.decode('utf-8')
            self.assertIn('background:transparent;', login_html)

            ann_html = ac.get(reverse('food:score_announcement') + '?year=2569').content.decode('utf-8')
            self.assertIn('background:transparent;', ann_html)

            sheet_html = ac.get(reverse('food:signin_sheet') + '?year=2569').content.decode('utf-8')
            self.assertIn('background: transparent;', sheet_html)

            # ตรวจ _load_logo_reader รักษา RGBA alpha + PDF generators
            rdr = _load_logo_reader(s.logo_center)
            self.assertIsNotNone(rdr)
            with Image.open(rdr.fp) as rdr_img:
                self.assertEqual(rdr_img.mode, 'RGBA')
                self.assertEqual(rdr_img.getpixel((0, 0))[3], 0)

            r_sheet_pdf = ac.get(reverse('food:signin_sheet') + '?year=2569&export=pdf')
            self.assertEqual(r_sheet_pdf.status_code, 200)
            self.assertEqual(r_sheet_pdf['Content-Type'], 'application/pdf')
            self.assertTrue(r_sheet_pdf.content.startswith(b'%PDF'))

            ann_pdf = build_announcement_pdf([], 2569, s)
            self.assertTrue(ann_pdf.startswith(b'%PDF'))

            cert_pdf = build_certificate_pdf([m_pdf], 2569, setting=s)
            self.assertTrue(cert_pdf.startswith(b'%PDF'))

    def test_all_menus_and_headings_match_admin_menu_names(self):
        """
        ตรวจสอบว่าชื่อเมนูทั้งหมด (Admin Topbar/Sidebar, Staff Sidebar/สิทธิ์,
        ชื่อหน้า/หัวข้อหน้า, แท็บย่อย และปุ่มลัด) ใช้คำเดียวกับเมนูหลักของ Admin
        """
        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        staff_all = User.objects.create_user(
            'staff_menu_all', 'password123', role='staff', is_first_login=False,
            staff_can_view_users=True, staff_can_attendance=True,
            staff_can_signin_sheet=True, staff_can_manage_food=True, staff_can_scores=True,
        )
        m_u = make_member('69DS98', year=2569, level='ม.4')

        # 1. Admin Topbar และ Sidebar ใช้ชื่อเมนูตรงกัน
        ac = Client()
        ac.force_login(admin)
        dash_html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        sidebar_html = dash_html[dash_html.index('id="sidebar"'):dash_html.index('class="main-content"')]
        topbar_html = dash_html[dash_html.index('class="topbar"'):dash_html.index('class="content-area"')]
        self.assertIn('เวลาเรียน / คะแนน', sidebar_html)
        self.assertIn('เวลาเรียน / คะแนน', topbar_html)
        self.assertNotIn('2. เวลาเรียน / คะแนน', sidebar_html)
        self.assertNotIn('2. เวลาเรียน / คะแนน', topbar_html)
        self.assertNotIn('(เพิ่ม/ลบ/แก้ไข/ค้นหา)', topbar_html)
        self.assertIn('ผู้ดูแลและเจ้าหน้าที่', topbar_html)
        self.assertIn('ผู้ใช้ (นักเรียน/สมาชิก)', topbar_html)

        # 2. Staff Sidebar และป้ายสิทธิ์ staff_permission_labels() ตรงกับเมนู Admin
        self.assertEqual(
            staff_all.staff_permission_labels(),
            ['ผู้ใช้ (นักเรียน/สมาชิก)', 'บันทึกเวลาเรียน', 'แบบลงชื่อ', 'รายการอาหาร / แบบสอบถาม', 'บันทึกคะแนน'],
        )
        sc = Client()
        sc.force_login(staff_all)
        staff_html = sc.get(reverse('food:staff_scores')).content.decode('utf-8')
        staff_sidebar = staff_html[staff_html.index('id="sidebar"'):staff_html.index('class="main-content"')]
        for expected in (
            'ผู้ใช้ (นักเรียน/สมาชิก)', 'เวลาเรียน / คะแนน',
            'บันทึกคะแนน', 'เกียรติบัตร', 'การประกาศผล', 'บันทึกเวลาเรียน',
            'รายการอาหาร', 'แบบสอบถาม',
        ):
            self.assertIn(expected, staff_sidebar)
        self.assertNotIn('รายชื่อนักเรียน/ผู้ใช้', staff_sidebar)
        self.assertNotIn('2. เวลาเรียน / คะแนน', staff_sidebar)
        self.assertNotIn('2. เวลาเรียน/คะแนน', staff_sidebar)

        # 3. ตรวจสอบ Title และ Heading ของหน้าต่างๆ ใช้ชื่อเดียวกับเมนู Admin
        page_checks = [
            ('composn:admin_dashboard', '<title>แดชบอร์ด', 'แดชบอร์ด'),
            ('composn:admin_admins', '<title>ผู้ดูแลและเจ้าหน้าที่', 'ผู้ดูแลและเจ้าหน้าที่'),
            ('composn:admin_users', '<title>ผู้ใช้ (นักเรียน/สมาชิก)', 'ผู้ใช้ (นักเรียน/สมาชิก)'),
            ('composn:admin_settings', '<title>ตั้งค่าสี ข้อความ', 'ตั้งค่าสี ข้อความ'),
            ('composn:admin_users_settings', '<title>ผู้ใช้', 'ผู้ใช้'),
            ('food:score_categories', '<title>กำหนดปีการศึกษา', 'กำหนดปีการศึกษา'),
            ('food:food_settings', '<title>รายการอาหาร', 'รายการอาหาร'),
            ('composn:admin_backups', '<title>บันทึกสำรอง', 'บันทึกสำรอง'),
            ('food:score_announcement', '<title>การประกาศผล', 'การประกาศผล'),
            ('composn:admin_use_records', '<title>แสดงรายชื่อผู้ใช้สิทธิ์', 'แสดงรายชื่อผู้ใช้สิทธิ์'),
        ]
        for url_name, exp_title, exp_h4 in page_checks:
            html = ac.get(reverse(url_name)).content.decode('utf-8')
            self.assertIn(exp_title, html, url_name)
            self.assertIn(exp_h4, html, url_name)

        sc_cat_html = ac.get(reverse('food:score_categories')).content.decode('utf-8')
        self.assertIn('บันทึกคะแนน', sc_cat_html)
        self.assertNotIn('ไปหน้ากรอกคะแนน', sc_cat_html)

        bak_html = ac.get(reverse('composn:admin_backups')).content.decode('utf-8')
        bak_head = bak_html[bak_html.index('class="content-area"'):bak_html.index('<div class="card mb-3">')]
        self.assertIn('บันทึกสำรอง', bak_head)
        self.assertNotIn('ไฟล์สำรอง', bak_head)

        mc = Client()
        mc.force_login(m_u)
        food_member_html = mc.get(reverse('food:my_food')).content.decode('utf-8')
        self.assertIn('<title>รายการอาหาร', food_member_html)
        self.assertNotIn('<title>เลือกอาหาร', food_member_html)

    def test_two_line_site_title_survey_target_groups_and_single_btn_attendance(self):
        """
        ทดสอบ:
        1. ชื่อระบบ 2 บรรทัด (site_title + site_title_line2) แสดงใน Sidebar, Login และรวมบน <title>
        2. แบบสอบถามเลือกกลุ่มเป้าหมาย 'นักเรียนปัจจุบัน' / 'นักเรียนใช้สิทธิ์' ได้เหมือนอาหาร (ทั้งแสดงผล, แจ้งเตือน, และป้องกัน POST ข้ามกลุ่ม)
        3. หน้าบันทึกเวลาเรียนแบบปุ่มเดียวสะอาดตา (#att-status-popover, ปุ่มเดียวต่อช่องทั้งวันอบรมและวันสอบ, ยุบตัวเลือกวันที่, AJAX save)
        """
        from datetime import timedelta
        from django.utils import timezone
        from food.models import Attendance, CampSurvey, CampSurveyAnswer, TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        setting = SystemSetting.get()
        setting.current_academic_year = 2569
        setting.food_system_open = True
        setting.save()

        m_curr = make_member('69DS71', year=2569, level='ม.4', first_name='ปุ่มเดียว', last_name='ปัจจุบัน')
        m_rights = make_member('68DS72', year=2568, level='ม.3', first_name='สิทธิ์', last_name='แบบสอบถาม')
        UseRecord.objects.create(
            user=m_rights, academic_year=2569, camp_level=1, rights_code='68DS72', is_used=True,
        )

        ac = Client()
        ac.force_login(admin)

        # 1. ทดสอบชื่อระบบ 2 บรรทัด
        r_set = ac.post(reverse('composn:admin_settings'), {
            'action': 'save_all_settings',
            'site_title': 'DEBSIRIN POSN',
            'site_title_line2': 'ค่าย 1 คอมพิวเตอร์',
            'current_academic_year': '2569',
            'max_rights': '5',
            'max_absences': '5',
            'auto_logout_minutes': '60',
        })
        self.assertEqual(r_set.status_code, 302)
        setting.refresh_from_db()
        self.assertEqual(setting.site_title, 'DEBSIRIN POSN')
        self.assertEqual(setting.site_title_line2, 'ค่าย 1 คอมพิวเตอร์')
        self.assertEqual(setting.site_title_inline, 'DEBSIRIN POSN ค่าย 1 คอมพิวเตอร์')

        dash_html = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        dash_sb = dash_html[dash_html.index('id="sidebar"'):dash_html.index('class="main-content"')]
        self.assertIn('DEBSIRIN POSN', dash_sb)
        self.assertIn('ค่าย 1 คอมพิวเตอร์', dash_sb)
        self.assertIn('<title>แดชบอร์ด - DEBSIRIN POSN ค่าย 1 คอมพิวเตอร์</title>', dash_html)

        login_html = Client().get(reverse('composn:login')).content.decode('utf-8')
        self.assertIn('<title>เข้าสู่ระบบ - DEBSIRIN POSN ค่าย 1 คอมพิวเตอร์</title>', login_html)
        self.assertIn('DEBSIRIN POSN', login_html)
        self.assertIn('ค่าย 1 คอมพิวเตอร์', login_html)

        mc_curr = Client()
        mc_curr.force_login(m_curr)
        mc_rights = Client()
        mc_rights.force_login(m_rights)
        prof_curr = mc_curr.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('ค่าย 1 คอมพิวเตอร์', prof_curr)

        # 2. ทดสอบแบบสอบถามเลือกกลุ่มเป้าหมาย (นักเรียนปัจจุบัน / นักเรียนใช้สิทธิ์)
        # ห้ามสร้างเมื่อ submit target_groups_submitted=1 โดยไม่เลือกกลุ่มใดเลย
        ac.post(reverse('food:surveys_manage'), {
            'year': '2569',
            'action': 'create_survey',
            'target_groups_submitted': '1',
            'title': 'ไม่ควรสร้างได้',
        })
        self.assertFalse(CampSurvey.objects.filter(title='ไม่ควรสร้างได้').exists())

        # สร้าง A (ปัจจุบันเท่านั้น), B (ใช้สิทธิ์เท่านั้น), C (ทั้ง 2 กลุ่ม)
        ac.post(reverse('food:surveys_manage'), {
            'year': '2569',
            'action': 'create_survey',
            'target_groups_submitted': '1',
            'allow_current_students': '1',
            'title': 'แบบสอบถามเฉพาะปัจจุบัน A',
            'first_question': 'ข้อ A',
            'first_question_type': 'text',
        })
        ac.post(reverse('food:surveys_manage'), {
            'year': '2569',
            'action': 'create_survey',
            'target_groups_submitted': '1',
            'allow_rights_students': '1',
            'title': 'แบบสอบถามเฉพาะใช้สิทธิ์ B',
            'first_question': 'ข้อ B',
            'first_question_type': 'text',
        })
        ac.post(reverse('food:surveys_manage'), {
            'year': '2569',
            'action': 'create_survey',
            'target_groups_submitted': '1',
            'allow_current_students': '1',
            'allow_rights_students': '1',
            'title': 'แบบสอบถามทั้งสองกลุ่ม C',
            'first_question': 'ข้อ C',
            'first_question_type': 'text',
        })
        sv_a = CampSurvey.objects.get(title='แบบสอบถามเฉพาะปัจจุบัน A')
        sv_b = CampSurvey.objects.get(title='แบบสอบถามเฉพาะใช้สิทธิ์ B')
        sv_c = CampSurvey.objects.get(title='แบบสอบถามทั้งสองกลุ่ม C')

        # ตรวจสิทธิ์การมองเห็น, badge, และบล็อก POST ข้ามกลุ่ม
        curr_sv_html = mc_curr.get(reverse('food:member_surveys')).content.decode('utf-8')
        self.assertIn('แบบสอบถามเฉพาะปัจจุบัน A', curr_sv_html)
        self.assertIn('แบบสอบถามทั้งสองกลุ่ม C', curr_sv_html)
        self.assertNotIn('แบบสอบถามเฉพาะใช้สิทธิ์ B', curr_sv_html)
        curr_sidebar = curr_sv_html[curr_sv_html.index('id="sidebar"'):curr_sv_html.index('class="main-content"')]
        curr_sv_nav = curr_sidebar[curr_sidebar.index('แบบสอบถาม'):curr_sidebar.index('แสดงรายชื่อผู้ใช้สิทธิ์')]
        self.assertIn('>2<', curr_sv_nav)
        curr_food_html = mc_curr.get(reverse('food:my_food')).content.decode('utf-8')
        self.assertIn('(2 รายการ)', curr_food_html)

        # ป้องกันนักเรียนปัจจุบันตอบแบบสอบถาม B (ใช้สิทธิ์อย่างเดียว)
        q_b = sv_b.questions.first()
        mc_curr.post(reverse('food:member_surveys'), {'survey_id': str(sv_b.pk), f'q_{q_b.pk}': 'แอบส่ง'})
        self.assertFalse(CampSurveyAnswer.objects.filter(survey=sv_b, user=m_curr).exists())

        rights_sv_html = mc_rights.get(reverse('food:member_surveys')).content.decode('utf-8')
        self.assertIn('แบบสอบถามเฉพาะใช้สิทธิ์ B', rights_sv_html)
        self.assertIn('แบบสอบถามทั้งสองกลุ่ม C', rights_sv_html)
        self.assertNotIn('แบบสอบถามเฉพาะปัจจุบัน A', rights_sv_html)
        q_a = sv_a.questions.first()
        mc_rights.post(reverse('food:member_surveys'), {'survey_id': str(sv_a.pk), f'q_{q_a.pk}': 'แอบส่ง'})
        self.assertFalse(CampSurveyAnswer.objects.filter(survey=sv_a, user=m_rights).exists())

        # Admin อัปเดตกลุ่มของ A ให้เปิดให้นักเรียนใช้สิทธิ์ตอบได้ด้วย
        ac.post(reverse('food:surveys_manage'), {
            'year': '2569',
            'action': 'update_survey_groups',
            'survey_id': str(sv_a.pk),
            'allow_current_students': '1',
            'allow_rights_students': '1',
        })
        sv_a.refresh_from_db()
        self.assertTrue(sv_a.allow_rights_students)
        self.assertIn('แบบสอบถามเฉพาะปัจจุบัน A', mc_rights.get(reverse('food:member_surveys')).content.decode('utf-8'))

        # 3. ทดสอบหน้าบันทึกเวลาเรียนแบบปุ่มเดียวสะอาดตา
        today = timezone.localdate()
        train_day = today - timedelta(days=1)
        exam_day = today
        TrainingDate.objects.create(academic_year=2569, date=train_day, event_type='training', title='อบรม')
        TrainingDate.objects.create(
            academic_year=2569, date=exam_day, event_type='exam', exam_period=TrainingDate.EXAM_MORNING, title='สอบ',
        )

        att_html = ac.get(
            reverse('food:staff_attendance') + f'?year=2569&date={train_day.isoformat()}&period=morning'
        ).content.decode('utf-8')
        self.assertNotIn('att-mini-grid', att_html)
        self.assertIn('id="att-status-popover"', att_html)
        self.assertIn('id="att-mode-group"', att_html)
        self.assertIn('id="colDateCheckboxes"', att_html)
        row_start = att_html.index(f'data-user="{m_curr.pk}"')
        row_end = att_html.index('</tr>', row_start)
        row_html = att_html[row_start:row_end]
        # มีปุ่ม matrix 2 ช่องวันอบรม (เช้า+บ่าย) + 1 ช่องวันสอบ (เช้า) = 3 ปุ่ม และนำคอลัมน์ สถานะ ที่ซ้ำออกแล้ว
        self.assertEqual(row_html.count('matrix-st-btn'), 3)
        self.assertEqual(row_html.count('att-cell-btn st-btn'), 0)
        self.assertIn('pinned-day-top', att_html)
        self.assertIn('btn-col-selected', att_html)
        self.assertIn('data-event-type="training"', row_html)
        self.assertIn('data-event-type="exam"', row_html)

        # AJAX บันทึกสถานะยังคงทำงานได้ครบ (absent/sick/personal/present)
        for st_val in ('absent', 'sick', 'personal', 'present'):
            r_ajax = ac.post(reverse('food:attendance_save'), {
                'user_id': m_curr.pk, 'date': train_day.isoformat(), 'period': 'morning', 'status': st_val, 'year': 2569,
            })
            self.assertEqual(r_ajax.status_code, 200)
            self.assertEqual(
                Attendance.objects.get(user=m_curr, date=train_day, session_period='morning').status,
                st_val,
            )

    def test_certificate_group_defaults_settings_and_prefilled_individual_edit(self):
        """
        ทดสอบการตั้งค่าข้อความเริ่มต้นของทุกประเภทเกียรติบัตรที่ ตั้งค่าระบบ > เกียรติบัตร
        และการใส่ข้อความเริ่มต้นให้อัตโนมัติในช่องแก้ไขรายบุคคล (คงลิงก์ค่าเริ่มต้นเมื่อไม่แก้ / แยกเฉพาะบุคคลเมื่อแก้)
        """
        from food.models import CertificateRecipient, CertificateSetting

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        ac = Client()
        ac.force_login(admin)

        # 1. หน้า ตั้งค่าระบบ > เกียรติบัตร (?tab=settings) มีการ์ดข้อความเริ่มต้นแยกครบทั้ง 3 กลุ่ม และไม่ซ้ำซ้อน
        settings_html = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=settings').content.decode('utf-8')
        self.assertIn('id="cert-default-texts-card"', settings_html)
        self.assertEqual(settings_html.count('name="student_line3_text"'), 1)
        self.assertEqual(settings_html.count('name="lecturer_line3_text"'), 1)
        self.assertEqual(settings_html.count('name="committee_line3_text"'), 1)

        # บันทึกข้อความเริ่มต้นใหม่ทั้ง 3 ประเภท
        stu_v1 = 'ผ่านการอบรมนักเรียน ค่าย ๑ ปีการศึกษา {year_thai}\nณ ศูนย์โรงเรียนเทพศิรินทร์ (รุ่น 1)'
        lec_v1 = 'ได้เข้าร่วมเป็นวิทยากรอบรมนักเรียน ค่าย ๑ สอวน. วิชาคอมพิวเตอร์\nณ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย ๑ โรงเรียนเทพศิรินทร์ จังหวัดกรุงเทพมหานคร (กลุ่มวิทยากร)'
        com_v1 = 'ได้เข้าร่วมเป็นคณะกรรมการดำเนินงาน ค่าย ๑ ปีการศึกษา {year_thai}\nณ ศูนย์โรงเรียนเทพศิรินทร์'
        r_save1 = ac.post(reverse('food:admin_certificates'), {
            'action': 'save_settings',
            'tab': 'settings',
            'year': 2569,
            'student_line3_text': stu_v1,
            'lecturer_line3_text': lec_v1,
            'committee_line3_text': com_v1,
        })
        self.assertEqual(r_save1.status_code, 302)
        cs = CertificateSetting.get()
        self.assertEqual(cs.student_line3_text, stu_v1)
        self.assertEqual(cs.lecturer_line3_text, lec_v1)
        self.assertEqual(cs.committee_line3_text, com_v1)

        # 2. แท็บ นักเรียน / วิทยากร / คณะกรรมการ แสดงข้อความเริ่มต้นของกลุ่ม + ลิงก์ตั้งค่า + pre-fill script
        tab1_html = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=students').content.decode('utf-8')
        self.assertIn('ผ่านการอบรมนักเรียน ค่าย ๑ ปีการศึกษา ๒๕๖๙', tab1_html)
        self.assertIn('tab=settings#cert-default-texts-card', tab1_html)

        tab2_html = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=lecturer').content.decode('utf-8')
        self.assertIn('(กลุ่มวิทยากร)', tab2_html)
        self.assertIn('CERT_DEFAULT_DESC', tab2_html)
        self.assertIn('resetRecipientDescToDefault()', tab2_html)
        self.assertIn('tab=settings#cert-default-texts-card', tab2_html)

        tab3_html = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=committee').content.decode('utf-8')
        self.assertIn('ได้เข้าร่วมเป็นคณะกรรมการดำเนินงาน ค่าย ๑ ปีการศึกษา ๒๕๖๙', tab3_html)

        # 3. เพิ่มวิทยากรโดยส่งค่าเริ่มต้นจาก Modal (ไม่แก้ไข) -> เก็บ rec.description == '' เพื่อตามค่าเริ่มต้นกลุ่ม
        ac.post(reverse('food:admin_certificates'), {
            'action': 'add_recipient',
            'year': 2569,
            'recipient_type': 'lecturer',
            'full_name': 'อ.ตรงค่าเริ่มต้น วิทยากร',
            'school_province': 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
            'description': lec_v1,
        })
        rec = CertificateRecipient.objects.get(full_name='อ.ตรงค่าเริ่มต้น วิทยากร')
        self.assertEqual(rec.description, '')
        self.assertEqual(rec.resolved_description(cs), lec_v1)

        # เมื่อ Admin แก้ไขข้อความเริ่มต้นกลุ่มวิทยากรเป็น v2 -> รายชื่อที่ใช้ค่าเริ่มต้นกลุ่มต้องเปลี่ยนตามอัตโนมัติ
        lec_v2 = 'ได้เข้าร่วมเป็นวิทยากรพิเศษ ค่าย ๑ ปีการศึกษา {year_thai}\nณ ศูนย์โรงเรียนเทพศิรินทร์ (อัปเดตกลุ่ม)'
        ac.post(reverse('food:admin_certificates'), {
            'action': 'save_settings',
            'tab': 'settings',
            'year': 2569,
            'student_line3_text': stu_v1,
            'lecturer_line3_text': lec_v2,
            'committee_line3_text': com_v1,
        })
        cs.refresh_from_db()
        rec.refresh_from_db()
        self.assertEqual(rec.resolved_description(cs), lec_v2)

        # 4. แก้ไขรายบุคคลแบบกำหนดเอง -> ต้องเก็บแยกเฉพาะบุคคลและไม่ถูกทับเมื่อเปลี่ยนค่าเริ่มต้นกลุ่ม
        custom_txt = 'วิทยากรบรรยายพิเศษหัวข้ออัลกอริทึมเฉพาะบุคคล'
        ac.post(reverse('food:admin_certificates'), {
            'action': 'edit_recipient',
            'year': 2569,
            'recipient_id': rec.pk,
            'recipient_type': 'lecturer',
            'full_name': rec.full_name,
            'school_province': rec.school_province,
            'description': custom_txt,
        })
        rec.refresh_from_db()
        self.assertEqual(rec.description, custom_txt)
        self.assertEqual(rec.resolved_description(cs), custom_txt)

        lec_v3 = 'ข้อความกลุ่ม v3'
        ac.post(reverse('food:admin_certificates'), {
            'action': 'save_settings',
            'tab': 'settings',
            'year': 2569,
            'student_line3_text': stu_v1,
            'lecturer_line3_text': lec_v3,
            'committee_line3_text': com_v1,
        })
        cs.refresh_from_db()
        self.assertEqual(rec.resolved_description(cs), custom_txt)

        # 5. รีเซ็ตกลับเป็นค่าเริ่มต้นของกลุ่ม -> description กลับมาเป็น '' และตามกลุ่ม v3
        ac.post(reverse('food:admin_certificates'), {
            'action': 'edit_recipient',
            'year': 2569,
            'recipient_id': rec.pk,
            'recipient_type': 'lecturer',
            'full_name': rec.full_name,
            'school_province': rec.school_province,
            'description': lec_v3,
        })
        rec.refresh_from_db()
        self.assertEqual(rec.description, '')
        self.assertEqual(rec.resolved_description(cs), lec_v3)

    def test_certificate_headers_ref_code_qr_and_public_verify(self):
        """
        ทดสอบ:
        1. ข้อความส่วนหัว 1 และ 2 อยู่บนสุดของตั้งค่าเกียรติบัตร ปรับแต่ง/รีเซ็ตตำแหน่งได้ และออก PDF ตัวอย่างได้
        2. การออกเลขที่อ้างอิงเกียรติบัตร (CertificateRecord) + QR Code บน PDF (คง ref_code เดิมเมื่อโหลดซ้ำ)
        3. หน้าตรวจสอบสาธารณะ /certificate/verify/<ref_code>/ สำหรับนักเรียน วิทยากร คณะกรรมการ รหัสผิด และยกเลิก
        """
        from django.utils import timezone
        from food.models import Attendance, CertificateRecipient, CertificateRecord, CertificateSetting, TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        m_stu = make_member('69DS98', year=2569, level='ม.4')
        m_stu.first_name, m_stu.last_name, m_stu.school_name = 'นายธนกร', 'ทดสอบระบบ', 'โรงเรียนเทพศิรินทร์'
        m_stu.save()

        today = timezone.localdate()
        td = TrainingDate.objects.create(academic_year=2569, date=today, event_type='training', title='อบรมวันทดสอบ')
        Attendance.objects.create(
            user=m_stu, academic_year=2569, date=td.date, session_period='morning',
            status=Attendance.STATUS_PRESENT, recorded_by=admin,
        )

        ac = Client()
        ac.force_login(admin)

        # 1. หน้าตั้งค่ามีส่วนหัว 1 & 2 อยู่ด้านบนสุด พร้อมข้อความเริ่มต้นตามที่กำหนด
        settings_html = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=settings').content.decode('utf-8')
        for acc_id in (
            'cert-settings-accordion', 'collapse-cert-header1', 'collapse-cert-header2',
            'collapse-cert-bg', 'collapse-cert-defaults', 'collapse-cert-line1',
            'collapse-cert-line4', 'collapse-cert-qr', 'btn-cert-expand-all', 'btn-cert-collapse-all',
        ):
            self.assertIn(f'id="{acc_id}"', settings_html)
        self.assertIn('id="cert-header1-card"', settings_html)
        self.assertIn('id="cert-header2-card"', settings_html)
        self.assertIn('id="cert-qr-card"', settings_html)
        self.assertLess(settings_html.index('id="cert-header1-card"'), settings_html.index('id="inp-bg-image"'))
        self.assertLess(settings_html.index('id="cert-header2-card"'), settings_html.index('id="inp-bg-image"'))
        self.assertIn('มูลนิธิส่งเสริมโอลิมปิกวิชาการและพัฒนามาตรฐานวิทยาศาสตร์ศึกษาฯ สอวน.', settings_html)
        self.assertIn('ร่วมกับ ศูนย์โอลิมปิกวิชาการ สอวน. ค่าย 1 โรงเรียนเทพศิรินทร์', settings_html)
        self.assertIn('ขอมอบเกียรติบัตรฉบับนี้ให้ไว้เพื่อแสดงว่า', settings_html)

        # บันทึกค่าปรับแต่งส่วนหัว 1, 2 และ QR Code
        r_save = ac.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
            'action': 'save_settings',
            'tab': 'settings',
            'year': 2569,
            'has_header_fields': '1',
            'has_qr_fields': '1',
            'header1_text': 'มูลนิธิ สอวน.\nร่วมกับ ศูนย์ สอวน. ค่าย 1 เทพศิรินทร์',
            'header1_y': '42.0',
            'header1_font_size': '20.0',
            'header1_line_spacing': '9.5',
            'header1_bold': '1',
            'header1_color': '#0f4c5c',
            'header2_text': 'ขอมอบเกียรติบัตรเพื่อแสดงว่า',
            'header2_y': '68.0',
            'header2_font_size': '23.0',
            'header2_bold': '1',
            'header2_color': '#b8860b',
            'line1_y': '94.0', 'line1_font_size': '34.0', 'line1_bold': '1', 'line1_color': '#000000',
            'line2_y': '111.0', 'line2_font_size': '22.0', 'line2_bold': '1', 'line2_color': '#000000',
            'line3_y': '127.0', 'line3_font_size': '21.0', 'line3_line_spacing': '11.0', 'line3_bold': '1', 'line3_color': '#000000',
            'line4_y': '162.0', 'line4_font_size': '21.0', 'line4_bold': '1', 'line4_color': '#000000',
            'show_qr_verify': '1',
            'qr_x': '248.0',
            'qr_y': '160.0',
            'qr_size': '26.0',
            'ref_code_font_size': '10.0',
            'ref_code_color': '#111111',
        })
        self.assertEqual(r_save.status_code, 302)
        cs = CertificateSetting.get()
        self.assertIn('มูลนิธิ สอวน.', cs.header1_text)
        self.assertEqual(cs.header2_text, 'ขอมอบเกียรติบัตรเพื่อแสดงว่า')
        self.assertAlmostEqual(cs.qr_x, 248.0)
        self.assertAlmostEqual(cs.qr_size, 26.0)
        self.assertTrue(cs.show_qr_verify)

        # รีเซ็ตพิกัด: แกน Y/X กลับค่าเริ่มต้น แต่ข้อความส่วนหัวที่แก้ไว้ยังคงอยู่
        ac.post(reverse('food:admin_certificates') + '?year=2569&tab=settings', {
            'action': 'reset_settings', 'tab': 'settings', 'year': 2569,
        })
        cs.refresh_from_db()
        self.assertAlmostEqual(cs.header1_y, 45.0)
        self.assertAlmostEqual(cs.qr_x, 252.0)
        self.assertEqual(cs.header2_text, 'ขอมอบเกียรติบัตรเพื่อแสดงว่า')

        # Preview PDF ใช้ได้ปกติและไม่สร้าง CertificateRecord ปลอมลงฐานข้อมูล
        cnt_before_pv = CertificateRecord.objects.count()
        r_pv = ac.get(reverse('food:certificate_preview_pdf') + '?year=2569&group=student&header1_y=44&qr_size=24')
        self.assertEqual(r_pv.status_code, 200)
        self.assertTrue(r_pv.content.startswith(b'%PDF'))
        self.assertEqual(CertificateRecord.objects.count(), cnt_before_pv)

        # 2. ดาวน์โหลดเกียรติบัตรนักเรียน -> ออก ref_code และโหลดซ้ำได้ ref_code เดิม
        r_stu_pdf1 = ac.get(reverse('food:certificate_download', args=[m_stu.pk]) + '?year=2569')
        self.assertEqual(r_stu_pdf1.status_code, 200)
        self.assertTrue(r_stu_pdf1.content.startswith(b'%PDF'))
        stu_rec = CertificateRecord.objects.get(academic_year=2569, user=m_stu, recipient_type=CertificateRecord.TYPE_STUDENT)
        self.assertTrue(stu_rec.ref_code.startswith('DSPOSN-2569-S-'))
        r_stu_pdf2 = ac.get(reverse('food:certificate_download', args=[m_stu.pk]) + '?year=2569')
        self.assertEqual(r_stu_pdf2.status_code, 200)
        self.assertEqual(CertificateRecord.objects.filter(academic_year=2569, user=m_stu).count(), 1)

        # ดาวน์โหลดเกียรติบัตรวิทยากร -> ออก ref_code
        lec = CertificateRecipient.objects.create(
            academic_year=2569, recipient_type=CertificateRecipient.TYPE_LECTURER,
            full_name='อาจารย์วิทยากร ทดสอบ', school_province='โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร',
        )
        r_lec_pdf = ac.get(reverse('food:certificate_recipient_download', args=[lec.pk]))
        self.assertEqual(r_lec_pdf.status_code, 200)
        self.assertTrue(r_lec_pdf.content.startswith(b'%PDF'))
        lec_rec = CertificateRecord.objects.get(academic_year=2569, recipient=lec)
        self.assertTrue(lec_rec.ref_code.startswith('DSPOSN-2569-L-'))

        # 3. Public Verification Page (ไม่ต้อง Login)
        anon = Client()
        v_stu = anon.get(reverse('food:certificate_verify', args=[stu_rec.ref_code]))
        self.assertEqual(v_stu.status_code, 200)
        v_stu_html = v_stu.content.decode('utf-8')
        self.assertIn('ยืนยันความถูกต้อง: เกียรติบัตรฉบับนี้ออกโดยศูนย์อบรมจริง', v_stu_html)
        self.assertIn(stu_rec.ref_code, v_stu_html)
        self.assertIn('ธนกร ทดสอบระบบ', v_stu_html)

        # ตรวจแบบ GET ?code=... สำหรับวิทยากร
        v_lec = anon.get(reverse('food:certificate_verify_index') + f'?code={lec_rec.ref_code}')
        self.assertEqual(v_lec.status_code, 200)
        self.assertIn('อาจารย์วิทยากร ทดสอบ', v_lec.content.decode('utf-8'))

        # ผู้ใช้ที่ยังติด is_first_login=True ก็สามารถสแกนตรวจหน้าสาธารณะได้โดยไม่โดน redirect ไปเปลี่ยนรหัส
        m_first = make_member('69DS99', year=2569, level='ม.4')
        fc = Client()
        fc.force_login(m_first)
        self.assertEqual(fc.get(reverse('food:certificate_verify', args=[stu_rec.ref_code])).status_code, 200)

        # รหัสปลอม / ไม่มีอยู่
        v_bad = anon.get(reverse('food:certificate_verify', args=['DSPOSN-2569-FAKE-999999']))
        self.assertEqual(v_bad.status_code, 200)
        self.assertIn('ไม่พบเลขที่อ้างอิงเกียรติบัตรนี้ในระบบ', v_bad.content.decode('utf-8'))

        # ยกเลิกเกียรติบัตร -> แจ้งเตือนว่าถูกยกเลิก
        lec_rec.is_revoked = True
        lec_rec.save(update_fields=['is_revoked'])
        v_rev = anon.get(reverse('food:certificate_verify', args=[lec_rec.ref_code])).content.decode('utf-8')
        self.assertIn('ถูกยกเลิกการใช้งานแล้ว', v_rev)

    def test_score_categories_colors_layout_toggle_pills_and_password_delete(self):
        """
        ทดสอบ:
        1. สีกล่องหัวข้อใหญ่สุ่มไม่ซ้ำกันอัตโนมัติ และเลือก/บันทึกสีใหม่ได้
        2. ย้ายประเภทหัวข้อใหญ่และเวลาเปิด-ปิดเข้าไปในช่องชื่อหัวข้อคะแนน, นำข้อความ 'ใช้เวลาเปิด-ปิดตามหัวข้อใหญ่' ออก
        3. ปุ่ม แสดง / บันทึก (เปิด <-> ปิด) ส่งค่า '1'/'0' ทำงานถูกต้อง
        4. การลบหัวข้อที่มีคะแนนบันทึกอยู่แล้วต้องยืนยันด้วยรหัสผ่าน (ไม่ใส่/รหัสผิด=ห้ามลบ, รหัสถูก=ลบหัวข้อ+คะแนนสำเร็จ)
        """
        from decimal import Decimal
        from food.models import Score, ScoreCategory
        from food.services import save_score

        admin = User.objects.get(camp1_code='admin')
        admin.set_password('AdmPass!99')
        admin.is_first_login = False
        admin.save()
        m = make_member('69DS96', year=2569, level='ม.4')

        ac = Client()
        ac.force_login(admin)

        # 1. สร้างหัวข้อใหญ่ 3 หัวข้อโดยไม่ระบุสี -> ระบบต้องสุ่มสีไม่ซ้ำกันทุกกล่อง
        for idx, (gt, nm) in enumerate([
            (ScoreCategory.GROUP_COLLECT, 'คะแนนเก็บ'),
            (ScoreCategory.GROUP_THEORY, 'สอบทฤษฎี 1'),
            (ScoreCategory.GROUP_PRACTICE, 'สอบปฏิบัติ 1'),
        ], 1):
            ac.post(reverse('food:score_categories'), {
                'action': 'create', 'year': '2569', 'tab': 'scores',
                'group_type': gt, 'name': nm, 'weight': '20', 'default_max_score': '100', 'order': str(idx),
            })
        mains = list(ScoreCategory.objects.filter(academic_year=2569, parent__isnull=True).order_by('order', 'id'))
        self.assertEqual(len(mains), 3)
        colors_set = {mc.box_color.lower() for mc in mains if mc.box_color}
        self.assertEqual(len(colors_set), 3)

        # สร้างหัวข้อย่อยและบันทึกคะแนนในหัวข้อที่ 2
        c2 = mains[1]
        sub_empty = ScoreCategory.objects.create(
            parent=c2, group_type=c2.group_type, name='ข้อย่อยไม่มีคะแนน',
            weight=Decimal('10'), default_max_score=Decimal('10'), academic_year=2569, order=1,
        )
        sub_scored = ScoreCategory.objects.create(
            parent=c2, group_type=c2.group_type, name='ข้อย่อยมีคะแนน',
            weight=Decimal('10'), default_max_score=Decimal('10'), academic_year=2569, order=2,
        )
        save_score(m, sub_scored, 1, Decimal('8'), Decimal('10'), recorded_by=admin)

        # 2. ตรวจโครงสร้าง HTML ของหน้ากำหนดหัวข้อคะแนน
        cat_html = ac.get(reverse('food:score_categories') + '?year=2569&tab=scores').content.decode('utf-8')
        thead_html = cat_html[cat_html.index('<thead'):cat_html.index('</thead>')]
        self.assertIn('id="btn-all-visible"', thead_html)
        self.assertIn('id="btn-all-open"', thead_html)
        self.assertIn('ชื่อหัวข้อคะแนน / หัวข้อย่อย (Sub)', thead_html)
        self.assertNotIn('ประเภทหัวข้อใหญ่', thead_html)
        self.assertNotIn('เวลาเริ่มเปิด', thead_html)
        self.assertNotIn('เวลาปิด', thead_html)
        self.assertNotIn('>สถานะ<', thead_html)
        self.assertNotIn('ใช้เวลาเปิด-ปิดตามหัวข้อใหญ่', cat_html)
        self.assertNotIn('(เปิดทั้งหมด)', cat_html)
        self.assertNotIn('(ทั้งหมด)', cat_html)
        self.assertIn(f'name="group_type_{c2.pk}"', cat_html)
        self.assertIn(f'name="open_at_{c2.pk}"', cat_html)
        self.assertIn(f'name="close_at_{c2.pk}"', cat_html)
        self.assertIn(f'name="box_color_{c2.pk}"', cat_html)

        # 3. ทดสอบ save_all_categories ส่ง '0' เพื่อปิด แสดง และ บันทึก พร้อมเปลี่ยนสีกล่อง
        c1 = mains[0]
        r_bulk = ac.post(reverse('food:score_categories'), {
            'action': 'save_all_categories',
            'year': '2569',
            'tab': 'scores',
            f'group_type_{c1.pk}': ScoreCategory.GROUP_COLLECT,
            f'name_{c1.pk}': 'คะแนนเก็บปรับสี',
            f'box_color_{c1.pk}': '#e11d48',
            f'weight_{c1.pk}': '30',
            f'default_max_score_{c1.pk}': '100',
            f'order_{c1.pk}': '1',
            f'is_visible_{c1.pk}': '0',
            f'is_open_{c1.pk}': '0',
        })
        self.assertEqual(r_bulk.status_code, 302)
        c1.refresh_from_db()
        self.assertEqual(c1.box_color, '#e11d48')
        self.assertFalse(c1.is_visible)
        self.assertFalse(c1.is_open)

        # 4. ทดสอบการลบ:
        # 4.1 หัวข้อไม่มีคะแนน -> ลบได้เลยโดยไม่ต้องใช้รหัสผ่าน
        ac.post(reverse('food:score_categories'), {
            'action': 'delete', 'year': '2569', 'tab': 'scores', 'id': str(sub_empty.pk),
        })
        self.assertFalse(ScoreCategory.objects.filter(pk=sub_empty.pk).exists())

        # 4.2 หัวข้อที่มีคะแนน -> ไม่ส่งรหัสผ่าน ห้ามลบ
        ac.post(reverse('food:score_categories'), {
            'action': 'delete', 'year': '2569', 'tab': 'scores', 'id': str(c2.pk),
        })
        self.assertTrue(ScoreCategory.objects.filter(pk=c2.pk).exists())
        self.assertTrue(Score.objects.filter(category=sub_scored).exists())

        # 4.3 ส่งรหัสผ่านผิด -> ห้ามลบ
        ac.post(reverse('food:score_categories'), {
            'action': 'delete', 'year': '2569', 'tab': 'scores', 'id': str(c2.pk),
            'confirm_password': 'wrong-password',
        })
        self.assertTrue(ScoreCategory.objects.filter(pk=c2.pk).exists())
        self.assertTrue(Score.objects.filter(category=sub_scored).exists())

        # 4.4 ส่งรหัสผ่าน Admin ถูกต้อง -> ลบหัวข้อใหญ่ + หัวข้อย่อย + คะแนนที่บันทึกไว้เรียบร้อย
        ac.post(reverse('food:score_categories'), {
            'action': 'delete', 'year': '2569', 'tab': 'scores', 'id': str(c2.pk),
            'confirm_password': 'AdmPass!99',
        })
        self.assertFalse(ScoreCategory.objects.filter(pk=c2.pk).exists())
        self.assertFalse(ScoreCategory.objects.filter(pk=sub_scored.pk).exists())
        self.assertFalse(Score.all_objects.filter(category_id=sub_scored.pk).exists())

    def test_attendance_date_filter_no_redundant_status_and_pinned_selected_date(self):
        """
        ทดสอบหน้าบันทึกเวลาเรียน:
        1. ปุ่มเลือกแสดงวันที่ (เฉพาะวันที่เลือก + กรอบเลือกวันที่เปิดให้เลือกง่าย)
        2. นำคอลัมน์ 'สถานะ (...)' ก่อน 'หมายเหตุ' ออกเมื่อมีตารางวันอบรม/สอบ
        3. คลิก/เลือกวันไหนแล้ว วันนั้นจะถูกย้ายมาชิดกับชื่อนักเรียนค้างไว้ (pinned-day-top / pinned-day-sub-1)
        """
        from datetime import timedelta
        from django.utils import timezone
        from food.models import TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        m = make_member('69DS97', year=2569, level='ม.4')

        today = timezone.localdate()
        d1 = today - timedelta(days=2)
        d2 = today - timedelta(days=1)
        d3 = today
        TrainingDate.objects.create(academic_year=2569, date=d1, event_type='training', title='วันแรก')
        TrainingDate.objects.create(academic_year=2569, date=d2, event_type='training', title='วันที่สอง')
        TrainingDate.objects.create(academic_year=2569, date=d3, event_type='exam', exam_period=TrainingDate.EXAM_BOTH, title='วันสอบ')

        ac = Client()
        ac.force_login(admin)

        # เลือกวันที่สอง (d2) -> คอลัมน์แรกของวันที่ต้องเป็น d2 ชิดกับชื่อ-นามสกุล
        h2 = ac.get(reverse('food:staff_attendance') + f'?year=2569&date={d2.isoformat()}&period=morning').content.decode('utf-8')
        self.assertIn('id="btn-col-selected"', h2)
        self.assertIn('class="collapse show mt-2 pt-2 border-top" id="colDateCheckboxes"', h2)
        thead2 = h2[h2.index('<thead'):h2.index('</thead>')]
        self.assertNotIn('<th rowspan="2" class="day-group-start" style="min-width:105px;">สถานะ', thead2)
        idx_d1 = thead2.index(f'data-col-date="{d1.isoformat()}"')
        idx_d2 = thead2.index(f'data-col-date="{d2.isoformat()}"')
        idx_d3 = thead2.index(f'data-col-date="{d3.isoformat()}"')
        self.assertLess(idx_d2, idx_d1)
        self.assertLess(idx_d2, idx_d3)
        self.assertIn(f'pinned-day-top pinned-day-last border-warning border-2" data-col-date="{d2.isoformat()}"', thead2)

        # เลือกวันที่สาม (d3) -> คอลัมน์แรกต้องเปลี่ยนเป็น d3 ชิดกับชื่อ-นามสกุลแทน
        h3 = ac.get(reverse('food:staff_attendance') + f'?year=2569&date={d3.isoformat()}&period=morning').content.decode('utf-8')
        thead3 = h3[h3.index('<thead'):h3.index('</thead>')]
        self.assertLess(thead3.index(f'data-col-date="{d3.isoformat()}"'), thead3.index(f'data-col-date="{d1.isoformat()}"'))
        self.assertLess(thead3.index(f'data-col-date="{d3.isoformat()}"'), thead3.index(f'data-col-date="{d2.isoformat()}"'))
        self.assertIn(f'pinned-day-top pinned-day-last border-warning border-2" data-col-date="{d3.isoformat()}"', thead3)
        # ตรวจเครื่องมือคลิกบันทึกเวลาเรียนใหม่ (เมนู + จิ้มด่วน 4 สถานะ, ไม่มี lpTimer)
        self.assertIn('data-att-tool="menu"', h2)
        self.assertIn('data-att-tool="stamp:present"', h2)
        self.assertIn('data-att-tool="stamp:absent"', h2)
        self.assertIn('data-att-tool="stamp:sick"', h2)
        self.assertIn('data-att-tool="stamp:personal"', h2)
        self.assertIn('id="att-tool-hint"', h2)
        self.assertIn('id="att-save-toast"', h2)
        self.assertIn('id="att-pinned-day-summary"', h2)
        self.assertNotIn('lpTimer', h2)
        self.assertNotIn('จิ้ม', h2)

    def test_site_favicon_upload_transparent_and_head_links(self):
        """
        ทดสอบไอคอนหน้า Titlebar (Favicon):
        1. หน้าตั้งค่า กำหนดปีการศึกษา แสดงช่องอัปโหลด site_favicon พร้อมแจ้งขนาดแนะนำ (64x64 / 128x128 / 256x256 px)
        2. อัปโหลดไฟล์ภาพ -> แปลงเป็น PNG พื้นโปร่งใสและแสดง <link rel="icon"> บน base.html, login.html, certificate_verify.html
        3. ลบ site_favicon -> <link rel="icon" หายไปอย่างสะอาด
        """
        import tempfile
        from PIL import Image, ImageDraw
        from django.core.files.uploadedfile import SimpleUploadedFile

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        ac = Client()
        ac.force_login(admin)

        set_html = ac.get(reverse('composn:admin_settings')).content.decode('utf-8')
        self.assertIn('name="site_favicon"', set_html)
        self.assertIn('64×64 / 128×128 / 256×256 px', set_html)

        fav_img = Image.new('RGB', (128, 128), (255, 255, 255))
        ImageDraw.Draw(fav_img).ellipse([24, 24, 104, 104], fill=(26, 115, 232))
        fav_buf = io.BytesIO()
        fav_img.save(fav_buf, format='JPEG')

        with tempfile.TemporaryDirectory() as tmpdir, override_settings(MEDIA_ROOT=tmpdir):
            r_up = ac.post(reverse('composn:admin_settings'), {
                'action': 'save_all_settings',
                'current_academic_year': '2569',
                'max_absences': '5',
                'auto_logout_minutes': '60',
                'site_favicon': SimpleUploadedFile('tab_icon.jpg', fav_buf.getvalue(), content_type='image/jpeg'),
            })
            self.assertEqual(r_up.status_code, 302)
            s = SystemSetting.get()
            self.assertTrue(s.site_favicon.name.lower().endswith('.png'))
            with Image.open(s.site_favicon.path) as fi:
                self.assertEqual(fi.mode, 'RGBA')
                self.assertEqual(fi.getpixel((0, 0))[3], 0)

            for url in (
                reverse('composn:admin_dashboard'),
                reverse('composn:login'),
                reverse('food:certificate_verify_index'),
            ):
                page_html = Client().get(url).content.decode('utf-8') if url != reverse('composn:admin_dashboard') else ac.get(url).content.decode('utf-8')
                self.assertIn(f'<link rel="icon" href="{s.site_favicon.url}">', page_html)

            # ลบ favicon
            ac.post(reverse('composn:admin_settings'), {
                'action': 'save_all_settings',
                'current_academic_year': '2569',
                'max_absences': '5',
                'auto_logout_minutes': '60',
                'clear_site_favicon': '1',
            })
            dash_after = ac.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
            self.assertNotIn('<link rel="icon"', dash_after)

    def test_certificates_export_zip_selectable_and_all_groups(self):
        """
        ทดสอบการดาวน์โหลดเกียรติบัตรแยกไฟล์รายบุคคล (.zip):
        1. แสดง checkbox เลือกรายบุคคล ปุ่มดาวน์โหลด ZIP ที่เลือก และ Modal เลือกกลุ่มทุกประเภท
        2. เลือกเฉพาะนักเรียนที่ต้องการ -> ได้ ZIP ที่มีเฉพาะ PDF ของคนที่เลือก (และไม่รวมคนที่ไม่มีสิทธิ์)
        3. ส่งโดยไม่เลือกใครเลย -> ป้องกันและ redirect กลับ
        4. เลือกเฉพาะวิทยากร -> ได้ ZIP เฉพาะวิทยากรที่เลือก
        5. ดาวน์โหลดทุกกลุ่ม (scope=all) -> ได้ ZIP แยกโฟลเดอร์ 01_นักเรียน / 02_วิทยากร / 03_คณะกรรมการ ครบถ้วน และเปิดปิดกลุ่มได้
        6. ตรวจสิทธิ์เข้าถึงของผู้ไม่มีสิทธิ์
        """
        import zipfile
        from django.utils import timezone
        from food.models import Attendance, CertificateRecipient, CertificateRecord, TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()
        staff_no_perm = User.objects.create_user(
            'stf_nozip', password='pass', role='staff', is_first_login=False, staff_can_scores=False,
        )

        # กำหนด max_absences = 1 และสร้างวันอบรม 1 วัน (2 ช่วง) เพื่อให้ m_bad ขาดทั้ง 2 ช่วงจนไม่มีสิทธิ์
        st = SystemSetting.get()
        st.max_absences = 1
        st.save(update_fields=['max_absences'])

        m1 = make_member('69DS91', year=2569, level='ม.4', first_name='เอกชัย', last_name='ผ่าน')
        m2 = make_member('69DS92', year=2569, level='ม.4', first_name='ทวีชัย', last_name='ผ่านด้วย')
        m_bad = make_member('69DS93', year=2569, level='ม.4', first_name='ไม่ผ่าน', last_name='ขาดสอบ')

        today = timezone.localdate()
        TrainingDate.objects.create(academic_year=2569, date=today, event_type='training', title='อบรมทดสอบ')
        for u in (m1, m2):
            Attendance.objects.create(user=u, academic_year=2569, date=today, session_period='morning', status='present', recorded_by=admin)
            Attendance.objects.create(user=u, academic_year=2569, date=today, session_period='afternoon', status='present', recorded_by=admin)
        Attendance.objects.create(user=m_bad, academic_year=2569, date=today, session_period='morning', status='absent', recorded_by=admin)
        Attendance.objects.create(user=m_bad, academic_year=2569, date=today, session_period='afternoon', status='absent', recorded_by=admin)

        lec1 = CertificateRecipient.objects.create(academic_year=2569, recipient_type='lecturer', full_name='วิทยากร หนึ่ง', school_province='กทม.')
        lec2 = CertificateRecipient.objects.create(academic_year=2569, recipient_type='lecturer', full_name='วิทยากร สอง', school_province='กทม.')
        com1 = CertificateRecipient.objects.create(academic_year=2569, recipient_type='committee', full_name='กรรมการ หนึ่ง', school_province='กทม.')

        ac = Client()
        ac.force_login(admin)

        # 1. ตรวจ UI บนหน้า Tab นักเรียน & Tab วิทยากร
        h_stu = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=students').content.decode('utf-8')
        self.assertIn('id="btn-zip-selected-students"', h_stu)
        self.assertIn('id="chk-all-cert-students"', h_stu)
        self.assertIn(f'class="form-check-input chk-cert-student" value="{m1.pk}"', h_stu)
        self.assertIn('id="exportAllZipModal"', h_stu)

        h_lec = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=lecturer').content.decode('utf-8')
        self.assertIn('id="btn-zip-selected-recipients"', h_lec)
        self.assertIn('id="chk-all-cert-recipients"', h_lec)
        self.assertIn(f'class="form-check-input chk-cert-recipient" value="{lec1.pk}"', h_lec)

        # 2. เลือกเฉพาะ m1 (+ แอบส่ง m_bad ที่ไม่มีสิทธิ์ -> ต้องถูกตัดออก)
        zip_url = reverse('food:certificates_export_zip')
        r_stu_zip = ac.post(zip_url, {
            'year': '2569',
            'scope': 'students',
            'select_mode': 'selected',
            'student_ids': [str(m1.pk), str(m_bad.pk)],
        })
        self.assertEqual(r_stu_zip.status_code, 200)
        self.assertEqual(r_stu_zip['Content-Type'], 'application/zip')
        with zipfile.ZipFile(io.BytesIO(r_stu_zip.content)) as zf:
            names = zf.namelist()
            self.assertEqual(len(names), 1)
            self.assertTrue(names[0].startswith('69DS91_'))
            self.assertTrue(names[0].endswith('.pdf'))
            self.assertTrue(zf.read(names[0]).startswith(b'%PDF'))
        self.assertTrue(CertificateRecord.objects.filter(academic_year=2569, user=m1).exists())
        self.assertFalse(CertificateRecord.objects.filter(academic_year=2569, user=m_bad).exists())

        # 3. ไม่เลือกใครเลย -> redirect 302
        r_empty = ac.post(zip_url, {
            'year': '2569', 'scope': 'students', 'select_mode': 'selected',
        })
        self.assertEqual(r_empty.status_code, 302)

        # 4. เลือกเฉพาะวิทยากร lec1
        r_lec_zip = ac.post(zip_url, {
            'year': '2569', 'scope': 'lecturer', 'select_mode': 'selected', 'recipient_ids': [str(lec1.pk)],
        })
        self.assertEqual(r_lec_zip.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(r_lec_zip.content)) as zf:
            names = zf.namelist()
            self.assertEqual(len(names), 1)
            self.assertIn('วิทยากร_หนึ่ง', names[0])
            self.assertTrue(zf.read(names[0]).startswith(b'%PDF'))

        # 5. ดาวน์โหลดทุกกลุ่มรวมเป็น ZIP เดียว แยกโฟลเดอร์ (m1, m2, lec1, lec2, com1 = 5 ไฟล์)
        r_all_zip = ac.post(zip_url, {
            'year': '2569',
            'scope': 'all',
            'select_mode': 'all',
            'has_group_toggles': '1',
            'include_students': '1',
            'include_lecturer': '1',
            'include_committee': '1',
        })
        self.assertEqual(r_all_zip.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(r_all_zip.content)) as zf:
            names = zf.namelist()
            self.assertEqual(len(names), 5)
            self.assertEqual(sum(1 for n in names if n.startswith('01_นักเรียน/')), 2)
            self.assertEqual(sum(1 for n in names if n.startswith('02_วิทยากร/')), 2)
            self.assertEqual(sum(1 for n in names if n.startswith('03_คณะกรรมการ/')), 1)
            for n in names:
                self.assertTrue(zf.read(n).startswith(b'%PDF'))

        # เลือกเฉพาะกลุ่มคณะกรรมการผ่าน Modal
        r_com_only = ac.post(zip_url, {
            'year': '2569', 'scope': 'all', 'select_mode': 'all', 'has_group_toggles': '1',
            'include_committee': '1',
        })
        self.assertEqual(r_com_only.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(r_com_only.content)) as zf:
            names = zf.namelist()
            self.assertEqual(len(names), 1)
            self.assertTrue(names[0].startswith('03_คณะกรรมการ/'))

        # 6. ตรวจสิทธิ์เข้าถึง
        self.assertEqual(Client().get(zip_url).status_code, 302)
        sc_no = Client()
        sc_no.force_login(staff_no_perm)
        self.assertEqual(sc_no.get(zip_url).status_code, 403)

    def test_announcement_cert_download_and_preview_fidelity(self):
        """
        ทดสอบ:
        1. ดาวน์โหลดเกียรติบัตรได้โดยตรงจากหน้า เอกสารประกาศผล (PDF รวม + ZIP เลือกรายบุคคล/กลุ่ม + return_to=announcement)
        2. ตัวอย่างหน้าตั้งค่าเกียรติบัตรตรงกับ PDF 1:1 (THSarabunNewCert override, font.ready, QR/caption, fallback chrome, Live PDF iframe)
        """
        import io
        import os
        import zipfile
        from django.conf import settings
        from django.utils import timezone
        from food.models import Attendance, CertificateRecipient, TrainingDate

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        st = SystemSetting.get()
        st.max_absences = 2
        st.save(update_fields=['max_absences'])

        m1 = make_member('69DS95', year=2569, level='ม.4', first_name='ธนากร', last_name='ประกาศผล')
        today = timezone.localdate()
        TrainingDate.objects.create(academic_year=2569, date=today, event_type='training', title='อบรมวันเดียว')
        Attendance.objects.create(user=m1, academic_year=2569, date=today, session_period='morning', status='present', recorded_by=admin)
        Attendance.objects.create(user=m1, academic_year=2569, date=today, session_period='afternoon', status='present', recorded_by=admin)

        lec = CertificateRecipient.objects.create(academic_year=2569, recipient_type='lecturer', full_name='อ.วิทยากร ทดสอบ', school_province='กทม.')
        com = CertificateRecipient.objects.create(academic_year=2569, recipient_type='committee', full_name='อ.กรรมการ ทดสอบ', school_province='กทม.')

        ac = Client()
        ac.force_login(admin)

        # 1. หน้า เอกสารประกาศผล ต้องมีส่วนดาวน์โหลดเกียรติบัตรครบ (PDF แยกกลุ่ม + Modal ZIP)
        ann_url = reverse('food:score_announcement') + '?year=2569'
        resp_ann = ac.get(ann_url)
        self.assertEqual(resp_ann.status_code, 200)
        html_ann = resp_ann.content.decode('utf-8')
        self.assertIn('id="announcement-cert-download-card"', html_ann)
        self.assertIn('id="announcementCertZipModal"', html_ann)
        self.assertIn('id="btn-announcement-cert-zip-modal"', html_ann)
        self.assertIn(reverse('food:certificates_export_all'), html_ann)
        self.assertIn(reverse('food:certificate_recipients_export_pdf'), html_ann)
        self.assertIn(reverse('food:certificates_export_zip'), html_ann)
        self.assertIn('return_to=announcement', html_ann)

        # 2. ส่งคำขอดาวน์โหลด ZIP จาก Modal บนหน้าเอกสารประกาศผล
        zip_url = reverse('food:certificates_export_zip')
        r_zip = ac.post(zip_url, {
            'year': '2569',
            'scope': 'all',
            'select_mode': 'selected',
            'has_group_toggles': '1',
            'return_to': 'announcement',
            'include_students': '1',
            'include_lecturer': '1',
            'include_committee': '1',
            'student_ids': [str(m1.pk)],
            'recipient_ids': [str(lec.pk), str(com.pk)],
        })
        self.assertEqual(r_zip.status_code, 200)
        self.assertEqual(r_zip['Content-Type'], 'application/zip')
        with zipfile.ZipFile(io.BytesIO(r_zip.content)) as zf:
            names = zf.namelist()
            self.assertEqual(len(names), 3)
            self.assertEqual(sum(1 for n in names if n.startswith('01_นักเรียน/')), 1)
            self.assertEqual(sum(1 for n in names if n.startswith('02_วิทยากร/')), 1)
            self.assertEqual(sum(1 for n in names if n.startswith('03_คณะกรรมการ/')), 1)
            for n in names:
                self.assertTrue(zf.read(n).startswith(b'%PDF'))

        # ไม่เลือกรายชื่อเลยพร้อม return_to=announcement -> redirect กลับหน้าเอกสารประกาศผล
        r_empty = ac.post(zip_url, {
            'year': '2569',
            'scope': 'all',
            'select_mode': 'selected',
            'has_group_toggles': '1',
            'return_to': 'announcement',
        })
        self.assertEqual(r_empty.status_code, 302)
        self.assertIn(reverse('food:score_announcement'), r_empty.url)

        # 3. ตรวจ certificate_preview_pdf (SAMEORIGIN สำหรับ iframe + lecturer_line3_text + clear_background_image)
        pv_pdf_url = (
            reverse('food:certificate_preview_pdf')
            + '?year=2569&group=lecturer&lecturer_line3_text=ข้อความทดสอบ&clear_background_image=1'
        )
        r_pdf = ac.get(pv_pdf_url)
        self.assertEqual(r_pdf.status_code, 200)
        self.assertEqual(r_pdf['Content-Type'], 'application/pdf')
        self.assertEqual(r_pdf.get('X-Frame-Options'), 'SAMEORIGIN')
        self.assertTrue(r_pdf.content.startswith(b'%PDF'))

        # 4. ตรวจหน้าตั้งค่าเกียรติบัตร (1:1 WYSIWYG + Live PDF Iframe Mode + ฟอนต์ THSarabunNew)
        resp_cfg = ac.get(reverse('food:admin_certificates') + '?year=2569&tab=settings')
        self.assertEqual(resp_cfg.status_code, 200)
        html_cfg = resp_cfg.content.decode('utf-8')
        self.assertIn('#cert-preview-stage *', html_cfg)
        self.assertIn('document.fonts.ready', html_cfg)
        self.assertIn('id="btn-pv-mode-wysiwyg"', html_cfg)
        self.assertIn('id="btn-pv-mode-pdf"', html_cfg)
        self.assertIn('id="cert-pdf-iframe"', html_cfg)
        self.assertIn('id="pv-qr-cap1"', html_cfg)
        self.assertIn('id="pv-qr-cap2"', html_cfg)
        self.assertIn('สแกนตรวจสอบเกียรติบัตร', html_cfg)
        self.assertIn('id="cert-fallback-chrome"', html_cfg)

        # 5. ยืนยันไฟล์ฟอนต์ THSarabunNew และ Bold อยู่ครบใน static/fonts/
        self.assertTrue(os.path.exists(os.path.join(settings.BASE_DIR, 'static', 'fonts', 'THSarabunNew.ttf')))
        self.assertTrue(os.path.exists(os.path.join(settings.BASE_DIR, 'static', 'fonts', 'THSarabunNew-Bold.ttf')))

    def test_announcement_settings_sync_1to1_and_two_page_pdf(self):
        """
        ทดสอบความตรงกัน 1:1 ระหว่างหน้าตั้งค่า/หน้าประกาศผล และการสร้าง PDF (2 หน้าพอดี):
        1. บันทึกค่าเอกสารประกาศผลจากหน้าตั้งค่า (score_categories tab=qualification) ครบทุกฟิลด์รวมย่อหน้า 1-2, กั้นหน้า-กั้นหลัง, ขนาดฟอนต์, ความกว้างคอลัมน์
        2. การตัดบรรทัดของย่อหน้า 1 และ 2 บนหน้าเว็บใช้ get_announcement_html_lines ตรงกับ ReportLab PDF 100%
        3. เอกสาร PDF มี 2 หน้าเสมอแม้มีรายชื่อจำนวนมาก (เช่น 45 คน) และมี X-Frame-Options: SAMEORIGIN สำหรับแสดงตัวอย่างในหน้าตั้งค่า
        """
        from food.announcement_pdf import build_announcement_pdf, get_announcement_html_lines
        from food.models import ScoreCategory
        from food.services import save_score
        import re

        admin = User.objects.get(camp1_code='admin')
        admin.is_first_login = False
        admin.save()

        ac = Client()
        ac.force_login(admin)

        # 1. บันทึกการตั้งค่าเอกสารประกาศผลจากหน้าตั้งค่า (score_categories?tab=qualification)
        r_save = ac.post(reverse('food:score_categories') + '?year=2569', {
            'action': 'save_qualification',
            'tab': 'qualification',
            'passing_score': '0',
            'max_qualified_count': '42',
            'announcement_school_name': 'โรงเรียนเทพศิรินทร์',
            'announcement_title': 'ประกาศผลการคัดเลือกเข้าค่ายโอลิมปิกวิชาการ (สอวน.)',
            'announcement_center_name': 'ศูนย์โรงเรียนเทพศิรินทร์-สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหาร ลาดกระบัง',
            'announcement_body_p1': 'ตามที่ สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน ได้ดำเนินการจัดตั้งศูนย์อบรม สอวน. สาขาวิชาคอมพิวเตอร์ ค่าย ๑ เป็นที่เรียบร้อยแล้วนั้น',
            'announcement_body_p2': 'บัดนี้ คณะกรรมการได้ตัดสินผลการสอบคัดเลือกนักเรียนเข้าค่าย ๒ จำนวน {count} คน ตามรายชื่อแนบท้ายประกาศนี้',
            'announcement_director_name': 'นายวิธาน พรหมสินธุศักดิ์',
            'announcement_director_position': 'ผู้อำนวยการโรงเรียนเทพศิรินทร์\nประธานศูนย์โรงเรียนเทพศิรินทร์ สอวน.วิชาคอมพิวเตอร์ ค่าย ๑',
            'announcement_date_text': 'ประกาศ ณ วันที่ ๒๑ พฤศจิกายน พ.ศ. ๒๕๖๙',
            'announcement_list_title': 'ประกาศรายชื่อผู้ที่ผ่านการคัดเลือกเข้า สอวน.คอมพิวเตอร์ ค่ายที่ 2 ปีการศึกษา 2569',
            'announcement_note': 'ให้นักเรียนรายงานตัวตามกำหนดการ',
            'announcement_margin_left_mm': '22.0',
            'announcement_margin_right_mm': '20.0',
            'announcement_margin_top_mm': '22.0',
            'announcement_margin_bottom_mm': '20.0',
            'announcement_font_size_title1': '18.0',
            'announcement_font_size_title2': '16.0',
            'announcement_font_size_title3': '14.5',
            'announcement_font_size_body': '15.0',
            'announcement_font_size_director_name': '15.0',
            'announcement_font_size_director_pos': '13.5',
            'announcement_font_size_list_title1': '15.0',
            'announcement_font_size_list_title2': '13.0',
            'announcement_font_size_table': '10.0',
            'announcement_line_spacing': '1.80',
            'announcement_paragraph_spacing_mm': '4.5',
            'announcement_indent_mm': '24.0',
            'announcement_garuda_height_mm': '32.0',
            'announcement_col_w_no': '14.0',
            'announcement_col_w_camp2': '24.0',
            'announcement_col_w_rights': '24.0',
            'announcement_col_w_name': '52.0',
            'announcement_col_w_school': '52.0',
            'announcement_col_w_grade': '14.0',
        })
        self.assertEqual(r_save.status_code, 302)

        st = SystemSetting.get()
        self.assertEqual(st.max_qualified_count, 42)
        self.assertEqual(st.announcement_margin_left_mm, 22.0)
        self.assertEqual(st.announcement_margin_right_mm, 20.0)
        self.assertEqual(st.announcement_font_size_title1, 18.0)
        self.assertEqual(st.announcement_font_size_title2, 16.0)
        self.assertEqual(st.announcement_font_size_title3, 14.5)
        self.assertEqual(st.announcement_font_size_body, 15.0)
        self.assertEqual(st.announcement_font_size_director_name, 15.0)
        self.assertEqual(st.announcement_font_size_director_pos, 13.5)
        self.assertEqual(st.announcement_font_size_list_title1, 15.0)
        self.assertEqual(st.announcement_font_size_list_title2, 13.0)
        self.assertEqual(st.announcement_paragraph_spacing_mm, 4.5)
        self.assertEqual(st.announcement_indent_mm, 24.0)
        self.assertEqual(st.announcement_garuda_height_mm, 32.0)
        self.assertEqual(st.announcement_col_w_name, 52.0)
        self.assertIn('{count}', st.announcement_body_p2)

        # 2. สร้างนักเรียน 42 คน เพื่อทดสอบว่าเอกสาร PDF ถูกจัดให้อยู่ใน 2 หน้าพอดี (ไม่ล้นไปหน้า 3)
        cat = ScoreCategory.objects.create(name='คะแนนรวมทดสอบ', academic_year=2569, weight=100, default_max_score=100)
        for idx in range(1, 43):
            u = make_member(f'69DS{idx:02d}', year=2569, level='ม.4', first_name=f'นักเรียน{idx}', last_name='ทดสอบสองหน้า')
            save_score(u, cat, 1, 100 - idx, 100, admin)

        # 3. ตรวจสอบหน้าเอกสารประกาศผล (HTML 1:1) และ PDF
        resp_ann = ac.get(reverse('food:score_announcement') + '?year=2569')
        self.assertEqual(resp_ann.status_code, 200)
        html_ann = resp_ann.content.decode('utf-8')
        self.assertIn('id="announcement-page-1"', html_ann)
        self.assertIn('id="announcement-page-2"', html_ann)
        self.assertIn('word-body-line', html_ann)
        self.assertIn('word-justified-line', html_ann)
        self.assertIn('--ann-para-gap-mm: 4.5mm;', html_ann)
        self.assertNotIn('class="word-j-tok"', html_ann)
        self.assertIn('จำนวน 42 คน', html_ann)
        self.assertIn('ค่าย 2 ปีการศึกษา 2569', html_ann)
        p1_start = html_ann.index('id="announcement-page-1"')
        p2_start = html_ann.index('id="announcement-page-2"')
        self.assertIsNone(re.search(r'[๐-๙]', html_ann[p1_start:p2_start]))

        p1_lines, p2_lines = get_announcement_html_lines(st, 42)
        self.assertGreaterEqual(len(p1_lines), 1)
        self.assertGreaterEqual(len(p2_lines), 1)
        for ln in p1_lines + p2_lines:
            self.assertIn(ln['text'], html_ann)
            self.assertEqual(len(ln['positioned_words']), 1)

        # 3.1 ตรวจสอบ TraceCanvas ของ draw_announcement_page ตามแบบสารบรรณ:
        #     - เส้นคั่นทึบตรงกลางยาว 50 มม. (ไม่ใช้เส้นจุด '.....')
        #     - ใช้ขนาดตัวอักษร body = 15.0pt จริง และแต่ละบรรทัดวาดเพียง 1 ครั้ง
        #     - ขอบซ้ายของ 'ประกาศ ณ วันที่...' ตรงกับขอบซ้ายของคำว่า 'ประกาศ...' ในหัวข้อบรรทัดที่ 1
        #     - ระยะห่างระหว่างบรรทัดสุดท้ายเนื้อหา -> วันที่ = 1.5 * line_h และ วันที่ -> ชื่อผู้ลงนาม = 3.0 * line_h
        import io
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as rl_canvas
        from food.announcement_pdf import draw_announcement_page, get_announcement_page1_metrics, _ensure_thai_fonts

        metrics = get_announcement_page1_metrics(st)
        self.assertIn('word-doc-divider-line', html_ann)
        self.assertIn(f"padding-left: {metrics['date_indent_mm']}mm;", html_ann)

        font_reg, font_bold = _ensure_thai_fonts()

        class _TraceCanvas(rl_canvas.Canvas):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.ops = []
                self.lines = []
                self._cur_font = ('', 0)

            def setFont(self, psfontname, size, leading=None):
                self._cur_font = (psfontname, float(size))
                return super().setFont(psfontname, size, leading=leading)

            def line(self, x1, y1, x2, y2):
                self.lines.append((float(x1), float(y1), float(x2), float(y2)))
                return super().line(x1, y1, x2, y2)

            def drawString(self, x, y, text, mode=None, charSpace=0, direction=None, wordSpace=None):
                self.ops.append(('L', float(x), float(y), self._cur_font[0], self._cur_font[1], text))
                return super().drawString(x, y, text, mode=mode, charSpace=charSpace, direction=direction, wordSpace=wordSpace)

            def drawCentredString(self, x, y, text, mode=None, charSpace=0, direction=None, wordSpace=None):
                self.ops.append(('C', float(x), float(y), self._cur_font[0], self._cur_font[1], text))
                return super().drawCentredString(x, y, text, mode=mode, charSpace=charSpace, direction=direction, wordSpace=wordSpace)

        tc = _TraceCanvas(io.BytesIO(), pagesize=A4)
        draw_announcement_page(tc, 2569, st, 42, A4[0], A4[1], font_reg, font_bold)
        expected_line_h = (15.0 * 0.44 * 1.80) * mm
        expected_para_gap = 4.5 * mm

        # ตรวจเส้นคั่นแนวนอนตรงกลางยาว 50 มม.
        self.assertEqual(len(tc.lines), 1)
        lx1, ly1, lx2, ly2 = tc.lines[0]
        self.assertAlmostEqual(ly1, ly2, delta=0.01)
        self.assertAlmostEqual(abs(lx2 - lx1), 50.0 * mm, delta=0.01)
        self.assertAlmostEqual((lx1 + lx2) / 2.0, metrics['center_x'], delta=0.01)

        # โครงสร้างลำดับ ops: [0..2]=3 หัวข้อ, [3..3+n1-1]=P1, [3+n1..3+n1+n2-1]=P2, [date], [sig_name], [pos1, pos2]
        n1 = len(p1_lines)
        n2 = len(p2_lines)
        self.assertEqual(len(tc.ops), 3 + n1 + n2 + 1 + 1 + 2)
        body_ops = tc.ops[3:3 + n1 + n2]
        for op in body_ops:
            self.assertAlmostEqual(op[4], 15.0, places=2)

        # ตรวจระยะห่างหัวข้อ 1->2, 2->3 (= 1 line_h), หัวข้อ 3 -> บรรทัดแรก P1 (= 2 line_h เว้นช่องเส้นคั่น) และภายใน P1
        self.assertAlmostEqual(tc.ops[0][2] - tc.ops[1][2], expected_line_h, delta=0.01)
        self.assertAlmostEqual(tc.ops[1][2] - tc.ops[2][2], expected_line_h, delta=0.01)
        self.assertAlmostEqual(tc.ops[2][2] - tc.ops[3][2], 2.0 * expected_line_h, delta=0.01)
        p1_ops = tc.ops[3:3 + n1]
        for prev_op, next_op in zip(p1_ops, p1_ops[1:]):
            self.assertAlmostEqual(prev_op[2] - next_op[2], expected_line_h, delta=0.01)

        # ตรวจระยะห่างระหว่างบรรทัดสุดท้าย P1 -> บรรทัดแรก P2 เท่ากับ expected_line_h + expected_para_gap
        p1_last_op = tc.ops[3 + n1 - 1]
        p2_first_op = tc.ops[3 + n1]
        self.assertAlmostEqual(p1_last_op[2] - p2_first_op[2], expected_line_h + expected_para_gap, delta=0.01)

        # ตรวจระยะห่างภายในย่อหน้า P2 เท่ากับ expected_line_h
        p2_ops = tc.ops[3 + n1:3 + n1 + n2]
        for prev_op, next_op in zip(p2_ops, p2_ops[1:]):
            self.assertAlmostEqual(prev_op[2] - next_op[2], expected_line_h, delta=0.01)

        # ตรวจระยะห่างระหว่างบรรทัดสุดท้าย P2 -> วันที่ = 1.5 * expected_line_h และขอบซ้ายตรงกับคำว่า 'ประกาศ...' ด้านบน
        date_op = tc.ops[3 + n1 + n2]
        self.assertEqual(date_op[0], 'L')
        self.assertAlmostEqual(date_op[1], metrics['center_x'] - (metrics['title1_w'] / 2.0), delta=0.01)
        self.assertAlmostEqual(date_op[1], metrics['date_x'], delta=0.01)
        self.assertAlmostEqual(p2_ops[-1][2] - date_op[2], 1.5 * expected_line_h, delta=0.01)

        # ตรวจว่าไม่มีตัวเลขไทย (๐-๙) หลงเหลืออยู่ในหน้าประกาศ PDF เลย
        for op in tc.ops:
            self.assertIsNone(re.search(r'[๐-๙]', op[5]))

        sig_name_op = tc.ops[3 + n1 + n2 + 1]
        pos_ops = tc.ops[3 + n1 + n2 + 2:]
        self.assertAlmostEqual(date_op[2] - sig_name_op[2], 3.0 * expected_line_h, delta=0.01)
        self.assertAlmostEqual(metrics['sig_x'], metrics['date_x'] + (metrics['date_w'] / 2.0), delta=0.01)
        self.assertEqual(sig_name_op[0], 'C')
        self.assertAlmostEqual(sig_name_op[1], metrics['sig_x'], delta=0.01)
        self.assertIn(f"margin-left: {metrics['sig_left_from_margin_mm']}mm;", html_ann)
        for pos_op in pos_ops:
            self.assertEqual(pos_op[0], 'C')
            self.assertAlmostEqual(pos_op[1], metrics['sig_x'], delta=0.01)
        self.assertAlmostEqual(sig_name_op[2] - pos_ops[0][2], expected_line_h, delta=0.01)
        self.assertAlmostEqual(pos_ops[0][2] - pos_ops[1][2], expected_line_h, delta=0.01)

        # 4. ตรวจสอบ PDF จริงต้องมี 2 หน้าพอดี (/Type /Page จำนวน 2 หน้า), แสดงใน iframe ได้ (SAMEORIGIN), และไม่ถูก cache ค้าง
        resp_pdf = ac.get(reverse('food:score_announcement') + '?year=2569&export=pdf')
        self.assertEqual(resp_pdf.status_code, 200)
        self.assertEqual(resp_pdf['Content-Type'], 'application/pdf')
        self.assertEqual(resp_pdf.get('X-Frame-Options'), 'SAMEORIGIN')
        self.assertIn('no-store', resp_pdf.get('Cache-Control', ''))
        pdf_bytes = resp_pdf.content
        self.assertTrue(pdf_bytes.startswith(b'%PDF'))
        page_count = len(re.findall(rb'/Type\s*/Page\b', pdf_bytes))
        self.assertEqual(page_count, 2)

        # 5. ทดสอบปรับระยะห่างระหว่างย่อหน้าเป็น 0 มม. ผ่านหน้าประกาศผล (ต้องไม่ถูกปัดกลับเป็นค่าเริ่มต้น 4.0 มม.)
        r_zero = ac.post(reverse('food:score_announcement'), {
            'action': 'save_announcement_text',
            'year': '2569',
            'only': 'passed',
            'announcement_paragraph_spacing_mm': '0',
        })
        self.assertEqual(r_zero.status_code, 302)
        st = SystemSetting.get()
        self.assertEqual(st.announcement_paragraph_spacing_mm, 0.0)
        html_zero = ac.get(reverse('food:score_announcement') + '?year=2569').content.decode('utf-8')
        self.assertIn('--ann-para-gap-mm: 0.0mm;', html_zero)

    def test_academic_year_range_2566_to_2569_and_add_new_year(self):
        """
        ทดสอบหน้ากำหนดปีการศึกษา (รวมอยู่ใน food:score_categories และตั้งค่าสี ข้อความใน composn:admin_settings):
        1. เริ่มต้นเลือกปีการศึกษาได้เฉพาะ 2566–2569 (ไม่มี 2563–2565 หรือ 2570–2574 ที่ยังไม่ได้สร้าง)
        2. มีปุ่ม 'เพิ่มปีการศึกษา' และ Modal แนะนำปีถัดไป (2570)
        3. ปฏิเสธการบันทึกปีที่ยังไม่ได้สร้าง และปฏิเสธการเพิ่มปีซ้ำหรือปี < 2566
        4. เพิ่มปี 2570 พร้อมตั้งเป็นปีปัจจุบันทันที -> บันทึกถาวรและปรากฏในทุกหน้าของระบบ
        5. สลับกลับเป็นปี 2569 แล้ว ปี 2570 ยังคงอยู่ในตัวเลือกของระบบ
        """
        import re as _re
        from composn.context_processors import get_academic_year_choices
        from composn.forms import UserForm

        admin = User.objects.create_superuser('adm_year', password='pass', is_first_login=False)
        ac = Client()
        ac.force_login(admin)

        # หน้า settings เดิมโหลดได้ปกติในฐานะ "ตั้งค่าสี ข้อความ"
        r_set_page = ac.get(reverse('composn:admin_settings'))
        self.assertEqual(r_set_page.status_code, 200)
        self.assertIn('ตั้งค่าสี ข้อความ', r_set_page.content.decode('utf-8'))

        # 1. ตรวจตัวเลือกเริ่มต้นในหน้ากำหนดปีการศึกษา (food:score_categories) และ UserForm
        r_get = ac.get(reverse('food:score_categories'))
        self.assertEqual(r_get.status_code, 200)
        html = r_get.content.decode('utf-8')
        m_select = _re.search(r'<select[^>]*name="current_academic_year"[^>]*>(.*?)</select>', html, _re.S)
        self.assertIsNotNone(m_select)
        select_html = m_select.group(1)
        for y in (2566, 2567, 2568, 2569):
            self.assertIn(f'value="{y}"', select_html)
        for y in (2563, 2564, 2565, 2570, 2571, 2574):
            self.assertNotIn(f'value="{y}"', select_html)

        self.assertIn('เพิ่มปีการศึกษา', html)
        self.assertIn('id="addAcademicYearModal"', html)
        self.assertIn('value="2570"', html)

        uf_years = {y for y, _ in UserForm().fields['academic_year'].widget.choices if y}
        self.assertEqual(uf_years, {2566, 2567, 2568, 2569})

        # 2. ป้องกันการตั้งค่าเป็นปีที่ยังไม่ได้สร้าง (ทั้งหน้ากำหนดปีการศึกษาและ fallback API)
        for bad_year in ('2565', '2570'):
            ac.post(reverse('food:score_categories'), {
                'action': 'save_academic_year_settings',
                'current_academic_year': bad_year,
                'max_absences': '5',
            })
            self.assertEqual(SystemSetting.get().current_academic_year, 2569)
            ac.post(reverse('composn:admin_settings'), {
                'action': 'save_all_settings',
                'site_title': 'ระบบ DS-COMPOSN',
                'current_academic_year': bad_year,
                'max_absences': '5',
                'auto_logout_minutes': '5',
            })
            self.assertEqual(SystemSetting.get().current_academic_year, 2569)

        # 3. ปฏิเสธการเพิ่มปีซ้ำ, ปีต่ำกว่า 2566 และค่าที่ไม่ใช่ตัวเลข
        for invalid_add in ('2569', '2565', 'abc', ''):
            r_bad = ac.post(reverse('food:score_categories'), {
                'action': 'add_academic_year',
                'new_academic_year': invalid_add,
                'set_as_current': '1',
            })
            self.assertEqual(r_bad.status_code, 302)
            st = SystemSetting.get()
            self.assertEqual(st.current_academic_year, 2569)
            self.assertEqual(st.get_academic_years(), [2569, 2568, 2567, 2566])

        # 4. เพิ่มปี 2570 และสลับเป็นปีการศึกษาปัจจุบันทันที
        r_add = ac.post(reverse('food:score_categories'), {
            'action': 'add_academic_year',
            'new_academic_year': '2570',
            'set_as_current': '1',
        })
        self.assertEqual(r_add.status_code, 302)
        st = SystemSetting.get()
        self.assertEqual(st.current_academic_year, 2570)
        self.assertEqual(st.get_academic_years(), [2570, 2569, 2568, 2567, 2566])
        self.assertEqual(get_academic_year_choices(st.current_academic_year, setting=st), [2570, 2569, 2568, 2567, 2566])

        # 5. ตรวจทุกหน้าที่ใช้ตัวเลือกปีการศึกษาว่าแสดง 2570
        for url_name in (
            'composn:admin_dashboard',
            'composn:admin_users',
            'composn:admin_user_create',
            'food:score_categories',
            'food:score_announcement',
        ):
            resp = ac.get(reverse(url_name))
            self.assertEqual(resp.status_code, 200, url_name)
            page_html = resp.content.decode('utf-8')
            self.assertIn('ปีการศึกษา 2570', page_html, url_name)

        # หน้ากำหนดปีการศึกษาต้องแนะนำปีถัดไปเป็น 2571 แล้ว
        html_after_add = ac.get(reverse('food:score_categories')).content.decode('utf-8')
        self.assertIn('value="2571"', html_after_add)

        # 6. สลับกลับเป็นปี 2569 -> ปี 2570 ยังคงอยู่ในตัวเลือก ไม่หายไป
        r_switch_back = ac.post(reverse('food:score_categories'), {
            'action': 'save_academic_year_settings',
            'current_academic_year': '2569',
            'max_absences': '5',
        })
        self.assertEqual(r_switch_back.status_code, 302)
        st.refresh_from_db()
        self.assertEqual(st.current_academic_year, 2569)
        self.assertIn(2570, st.get_academic_years())
        self.assertIn(2570, get_academic_year_choices(st.current_academic_year, setting=st))
        html_back = ac.get(reverse('food:score_categories')).content.decode('utf-8')
        m_sel_back = _re.search(r'<select[^>]*name="current_academic_year"[^>]*>(.*?)</select>', html_back, _re.S)
        self.assertIsNotNone(m_sel_back)
        self.assertIn('value="2570"', m_sel_back.group(1))


class SchoolNameNormalizationTests(TestCase):
    """ตรวจสอบการเติมคำว่า 'โรงเรียน' นำหน้าชื่อโรงเรียนอัตโนมัติ"""

    def test_normalize_school_name_helper(self):
        from .models import normalize_school_name

        cases = [
            ('เทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('สวนกุหลาบวิทยาลัย', 'โรงเรียนสวนกุหลาบวิทยาลัย'),
            ('กรุงเทพคริสเตียนวิทยาลัย', 'โรงเรียนกรุงเทพคริสเตียนวิทยาลัย'),
            ('สาธิตจุฬาลงกรณ์มหาวิทยาลัย (ฝ่ายมัธยม)', 'โรงเรียนสาธิตจุฬาลงกรณ์มหาวิทยาลัย (ฝ่ายมัธยม)'),
            ('โรงเรียนเทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('โรงเรียน  เทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('รร.เทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('ร.ร. สวนกุหลาบวิทยาลัย', 'โรงเรียนสวนกุหลาบวิทยาลัย'),
            ('รร เทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('รร. โรงเรียนเทพศิรินทร์', 'โรงเรียนเทพศิรินทร์'),
            ('สถาบันโคเซ็น สจล.', 'สถาบันโคเซ็น สจล.'),
            ('ศูนย์การเรียนนิวตัน', 'ศูนย์การเรียนนิวตัน'),
            ('วิทยาลัยเทคนิคกรุงเทพ', 'วิทยาลัยเทคนิคกรุงเทพ'),
            ('มหาวิทยาลัยธรรมศาสตร์', 'มหาวิทยาลัยธรรมศาสตร์'),
            ('', ''),
            (None, ''),
            ('   ', ''),
            ('-', ''),
            ('—', ''),
        ]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(normalize_school_name(raw), expected)

    def test_excel_preview_and_import_auto_prefix(self):
        from .excel import import_students_from_preview, preview_students

        headers = ['รหัสค่าย 1', 'ชื่อ-นามสกุล', 'โรงเรียน']
        rows = [
            ['69DS11', 'ด.ช.หนึ่ง ทดสอบ', 'เทพศิรินทร์'],
            ['69DS12', 'ด.ช.สอง ทดสอบ', 'รร. สวนกุหลาบวิทยาลัย'],
            ['69DS13', 'ด.ช.สาม ทดสอบ', 'โรงเรียนเตรียมอุดมศึกษา'],
            ['69DS14', 'ด.ช.สี่ ทดสอบ', 'สถาบันโคเซ็น สจล.'],
            ['69DS15', 'ด.ช.ห้า ทดสอบ', '-'],
        ]
        f = make_xlsx(headers, rows)
        pv = preview_students(f)
        schools_pv = {r['camp1_code']: r['school_name'] for r in pv['rows']}
        self.assertEqual(schools_pv['69DS11'], 'โรงเรียนเทพศิรินทร์')
        self.assertEqual(schools_pv['69DS12'], 'โรงเรียนสวนกุหลาบวิทยาลัย')
        self.assertEqual(schools_pv['69DS13'], 'โรงเรียนเตรียมอุดมศึกษา')
        self.assertEqual(schools_pv['69DS14'], 'สถาบันโคเซ็น สจล.')
        self.assertEqual(schools_pv['69DS15'], '')

        res = import_students_from_preview(pv)
        self.assertEqual(res['created'], 5)
        self.assertEqual(User.objects.get(camp1_code='69DS11').school_name, 'โรงเรียนเทพศิรินทร์')
        self.assertEqual(User.objects.get(camp1_code='69DS12').school_name, 'โรงเรียนสวนกุหลาบวิทยาลัย')
        self.assertEqual(User.objects.get(camp1_code='69DS14').school_name, 'สถาบันโคเซ็น สจล.')
        self.assertEqual(User.objects.get(camp1_code='69DS15').school_name, '')

        f2 = make_xlsx(headers, [['69DS16', 'ด.ช.หก ทดสอบ', 'สามเสนวิทยาลัย']])
        res2 = import_students(f2)
        self.assertEqual(res2['created'], 1)
        self.assertEqual(User.objects.get(camp1_code='69DS16').school_name, 'โรงเรียนสามเสนวิทยาลัย')

    def test_forms_and_certificate_school_formatting(self):
        from .forms import SchoolRequestForm, UserForm
        from food.models import CertificateSetting

        uf = UserForm(data={
            'camp1_code': '69DS21',
            'full_name': 'ด.ช.เจ็ด ทดสอบ',
            'school_name': 'บดินทรเดชา (สิงห์ สิงหเสนี)',
            'role': User.ROLE_MEMBER,
            'is_active': True,
        })
        self.assertTrue(uf.is_valid(), uf.errors)
        self.assertEqual(uf.cleaned_data['school_name'], 'โรงเรียนบดินทรเดชา (สิงห์ สิงหเสนี)')

        sf_ok = SchoolRequestForm(data={'new_school_name': 'ร.ร.หอวัง'})
        self.assertTrue(sf_ok.is_valid(), sf_ok.errors)
        self.assertEqual(sf_ok.cleaned_data['new_school_name'], 'โรงเรียนหอวัง')

        sf_bad = SchoolRequestForm(data={'new_school_name': '-'})
        self.assertFalse(sf_bad.is_valid())

        cs = CertificateSetting.get()
        self.assertEqual(cs.format_student_school('เทพศิรินทร์'), 'โรงเรียนเทพศิรินทร์ กรุงเทพมหานคร')































