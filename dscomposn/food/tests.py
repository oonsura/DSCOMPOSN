import datetime as dt
import io
import re
from decimal import Decimal

from django.test import Client, TestCase
from django.urls import reverse
from django.utils import timezone
from openpyxl import Workbook

from composn import levels, rights
from composn.excel import import_camp2
from composn.models import ArchiveRecord, AuditLog, RightsRound, SystemSetting, UseRecord, User
from food.models import (
    Attendance, QualificationResult, Score, ScoreApproval, ScoreCategory,
    TrainingDate, approve_year_scores, unapprove_year_scores,
)
from food.services import FoodError, delete_attendance, save_attendance, save_score


def make_member(code='69DS01', level='ม.4', year=2569, camp2=None, **kw):
    fname = kw.pop('first_name', 'สมชาย')
    lname = kw.pop('last_name', 'ใจดี')
    u = User(
        camp1_code=code,
        camp2_code=camp2,
        first_name=fname,
        last_name=lname,
        school_name=kw.pop('school_name', 'รร.ทดสอบ'),
        camp_center=kw.pop('camp_center', code[2:4]),
        academic_year=year,
        current_level=level,
        role=User.ROLE_MEMBER,
        is_first_login=False,
        **kw,
    )
    levels.apply_levels(u)
    u.set_password(code)
    u.save()
    return u


def make_staff(code='staff01', **kw):
    u = User(
        camp1_code=code,
        first_name='เจ้าหน้าที่',
        last_name='ทดสอบ',
        role=User.ROLE_STAFF,
        is_first_login=False,
        staff_can_view_users=True,
        staff_can_attendance=True,
        staff_can_signin_sheet=True,
        staff_can_manage_food=True,
        staff_can_scores=True,
        **kw,
    )
    u.set_password('StaffPass!1')
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


class FoodBaseTestCase(TestCase):
    def setUp(self):
        self.admin = User.objects.filter(camp1_code='admin').first()
        if not self.admin:
            self.admin = User.objects.create_superuser('admin', '1234')
        else:
            self.admin.set_password('1234')
            self.admin.is_first_login = False
            self.admin.save()

        self.setting = SystemSetting.get()
        self.setting.current_academic_year = 2569
        self.setting.max_qualified_count = 1
        self.setting.max_absences = 5
        self.setting.add_academic_year(2570)
        self.setting.save()

        self.client = Client()
        self.client.force_login(self.admin)

        self.m1 = make_member('69DS01', level='ม.4', year=2569, first_name='คนที่หนึ่ง', last_name='ทดสอบ')
        self.m2 = make_member('69DS02', level='ม.4', year=2569, first_name='คนที่สอง', last_name='ทดสอบ')

        self.main_cat = ScoreCategory.objects.create(
            name='คะแนนสอบรวม', weight=Decimal('100'), academic_year=2569,
            order=1, default_max_score=Decimal('100'), is_visible=True, is_open=True,
        )
        self.sub_cat = ScoreCategory.objects.create(
            parent=self.main_cat, name='สอบปฏิบัติ', weight=Decimal('100'), academic_year=2569,
            order=1, default_max_score=Decimal('100'), is_visible=True, is_open=True,
        )
        Score.objects.create(user=self.m1, category=self.sub_cat, exam_no=1, score=Decimal('90'), max_score=Decimal('100'))
        Score.objects.create(user=self.m2, category=self.sub_cat, exam_no=1, score=Decimal('60'), max_score=Decimal('100'))

        self.past_date = timezone.localdate() - dt.timedelta(days=5)
        self.td = TrainingDate.objects.create(
            academic_year=2569, date=self.past_date, event_type=TrainingDate.TYPE_TRAINING, title='วันอบรมทดสอบ',
        )


def _tag_with_id(html, elem_id):
    m = re.search(rf'<[^>]*\bid="{re.escape(elem_id)}"[^>]*>', html)
    return m.group(0) if m else ''


class ApprovalFlowAndPasswordTests(FoodBaseTestCase):
    def test_approve_button_red_to_green_and_current_password_required_to_unapprove(self):
        old_cat_url = reverse('food:score_categories') + '?year=2569&tab=qualification'
        redir = self.client.get(old_cat_url)
        self.assertEqual(redir.status_code, 302)
        self.assertIn(reverse('food:score_announcement'), redir['Location'])

        ann_qual_url = reverse('food:score_announcement') + '?year=2569&tab=qualification'

        # 1. Pre-approval: red card & red button, no unapprove button
        resp = self.client.get(ann_qual_url)
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode('utf-8')
        self.assertIn('border-danger', _tag_with_id(html, 'approval-status-card'))
        self.assertIn('btn-danger', _tag_with_id(html, 'btn-approve-scores'))
        self.assertNotIn('id="btn-unapprove-scores"', html)

        # 2. Approve scores -> turns green & locked, QualificationResult created, camp2_code assigned
        resp = self.client.post(reverse('food:score_announcement'), {
            'action': 'approve_scores',
            'tab': 'qualification',
            'year': '2569',
        }, follow=True)
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(ScoreApproval.is_year_approved(2569))
        appr = ScoreApproval.objects.get(academic_year=2569)
        orig_approved_at = appr.approved_at

        self.m1.refresh_from_db()
        self.m2.refresh_from_db()
        self.assertIsNotNone(self.m1.camp2_code)
        self.assertTrue(QualificationResult.objects.filter(academic_year=2569, user=self.m1, passed=True).exists())
        self.assertTrue(QualificationResult.objects.filter(academic_year=2569, user=self.m2, passed=False).exists())

        html = resp.content.decode('utf-8')
        self.assertIn('border-success', _tag_with_id(html, 'approval-status-card'))
        approve_btn = _tag_with_id(html, 'btn-approve-scores')
        self.assertIn('btn-success', approve_btn)
        self.assertIn('disabled', approve_btn)
        self.assertNotIn('btn-danger', approve_btn)
        self.assertIn('id="btn-unapprove-scores"', html)
        self.assertIn('id="unapproveScoresPasswordModal"', html)

        # 3. Re-posting approve_scores while approved is blocked/no-op
        self.client.post(reverse('food:score_announcement'), {
            'action': 'approve_scores',
            'tab': 'qualification',
            'year': '2569',
        })
        appr.refresh_from_db()
        self.assertEqual(appr.approved_at, orig_approved_at)

        # 4. Unapprove requires current admin password (empty or wrong rejected)
        self.client.post(reverse('food:score_announcement'), {
            'action': 'unapprove_scores',
            'tab': 'qualification',
            'year': '2569',
            'confirm_password': '',
        })
        self.assertTrue(ScoreApproval.is_year_approved(2569))

        self.client.post(reverse('food:score_announcement'), {
            'action': 'unapprove_scores',
            'tab': 'qualification',
            'year': '2569',
            'confirm_password': 'wrong-password',
        })
        self.assertTrue(ScoreApproval.is_year_approved(2569))
        self.assertTrue(
            AuditLog.objects.filter(action='update_settings', new_value__contains='unapprove_scores_failed').exists()
        )

        # 5. Changing admin password -> old password ('1234') rejected, NEW current password succeeds
        self.admin.set_password('NewAdminPass!2569')
        self.admin.save()
        self.client.force_login(self.admin)

        self.client.post(reverse('food:score_announcement'), {
            'action': 'unapprove_scores',
            'tab': 'qualification',
            'year': '2569',
            'confirm_password': '1234',
        })
        self.assertTrue(ScoreApproval.is_year_approved(2569))

        resp = self.client.post(reverse('food:score_announcement'), {
            'action': 'unapprove_scores',
            'tab': 'qualification',
            'year': '2569',
            'confirm_password': 'NewAdminPass!2569',
        }, follow=True)
        self.assertEqual(resp.status_code, 200)
        self.assertFalse(ScoreApproval.is_year_approved(2569))
        self.m1.refresh_from_db()
        self.assertIsNone(self.m1.camp2_code)

        html = resp.content.decode('utf-8')
        self.assertIn('border-danger', _tag_with_id(html, 'approval-status-card'))
        self.assertIn('btn-danger', _tag_with_id(html, 'btn-approve-scores'))
        self.assertNotIn('id="btn-unapprove-scores"', html)


class ResultLockingTests(FoodBaseTestCase):
    def setUp(self):
        super().setUp()
        # Create pre-approval records needed to test deletion/restore/rights locks
        self.att = Attendance.objects.create(
            user=self.m1, academic_year=2569, date=self.past_date,
            session_period=Attendance.PERIOD_MORNING, status=Attendance.STATUS_ABSENT, recorded_by=self.admin,
        )
        self.soft_att = Attendance.objects.create(
            user=self.m2, academic_year=2569, date=self.past_date,
            session_period=Attendance.PERIOD_MORNING, status=Attendance.STATUS_SICK, recorded_by=self.admin,
        )
        self.soft_att.delete(user=self.admin)

        self.round_2569 = RightsRound.objects.create(
            academic_year=2569, camp_level=1, title='รอบทดสอบ 2569', is_open=True, opened_by=self.admin,
        )
        self.m_rights = make_member('68DS01', level='ม.4', year=2568, first_name='สิทธิ์ปี', last_name='หกแปด')
        self.use_rec = UseRecord.objects.create(
            user=self.m_rights, rights_round=self.round_2569, academic_year=2569, camp_level=1,
            rights_code='69DS99', grade_level='ม.5', confirmed_at=timezone.now(), is_used=True,
        )
        self.arch_rec = ArchiveRecord.objects.create(
            user=self.m_rights, use_record=self.use_rec, academic_year=2569, camp_level=1, rights_code='69DS99',
        )
        self.soft_use = UseRecord.objects.create(
            user=make_member('68DS02', level='ม.4', year=2568), rights_round=self.round_2569,
            academic_year=2569, camp_level=1, rights_code='69DS88', grade_level='ม.5',
            confirmed_at=timezone.now(), is_used=True,
        )
        self.soft_use.delete(user=self.admin)

        # Approve year 2569 now
        approve_year_scores(2569, self.admin)
        self.m1.refresh_from_db()

    def test_service_layer_locks(self):
        # Attendance services locked
        with self.assertRaises(FoodError):
            save_attendance(self.m1, self.past_date, Attendance.PERIOD_AFTERNOON, Attendance.STATUS_ABSENT, recorded_by=self.admin)
        with self.assertRaises(FoodError):
            delete_attendance(self.att, by=self.admin)

        # Score service locked
        with self.assertRaises(FoodError):
            save_score(self.m1, self.sub_cat, 1, Decimal('50'), Decimal('100'), recorded_by=self.admin)

        # Rights services locked
        elig = rights.check_eligibility(make_member('68DS03', level='ม.4', year=2568), self.round_2569)
        self.assertFalse(elig.ok)
        self.assertEqual(elig.code, 'results_approved')

        with self.assertRaises(rights.RightsError):
            rights.cancel_rights(self.m_rights, self.round_2569)
        with self.assertRaises(rights.RightsError):
            rights.open_round(2569, 2, self.admin)

        closed_round = RightsRound.objects.create(academic_year=2569, camp_level=2, title='รอบปิดแล้ว', is_open=False, opened_by=self.admin)
        with self.assertRaises(rights.RightsError):
            rights.reopen_round(closed_round, self.admin)
        with self.assertRaises(rights.RightsError):
            rights.import_camp1_rights(['68DS03'], self.admin, academic_year=2569)

        # Closing an already-open round while approved remains allowed
        rights.close_round(self.round_2569, self.admin, mark_waived=False)
        self.round_2569.refresh_from_db()
        self.assertFalse(self.round_2569.is_open)

        # import_camp2 locked for approved year
        f = make_xlsx(['camp1_code', 'camp2_code'], [['69DS02', '69KM99']])
        res = import_camp2(f)
        self.assertEqual(res['updated'], 0)
        self.assertGreater(len(res['errors']), 0)
        self.m2.refresh_from_db()
        self.assertIsNone(self.m2.camp2_code)

    def test_attendance_and_score_view_locks(self):
        # 1. staff_attendance GET skips auto-bulk_create & shows lock banner
        m3 = make_member('69DS03', level='ม.4', year=2569)
        resp = self.client.get(reverse('food:staff_attendance') + f'?year=2569&date={self.past_date:%Y-%m-%d}')
        self.assertEqual(resp.status_code, 200)
        self.assertIn('id="att-approved-lock-banner"', resp.content.decode('utf-8'))
        self.assertFalse(Attendance.objects.filter(user=m3, date=self.past_date).exists())

        # 2. attendance_save_api returns 403 JSON
        resp = self.client.post(reverse('food:attendance_save'), {
            'user_id': self.m1.pk,
            'date': f'{self.past_date:%Y-%m-%d}',
            'period': Attendance.PERIOD_MORNING,
            'status': Attendance.STATUS_PRESENT,
        })
        self.assertEqual(resp.status_code, 403)
        self.assertFalse(resp.json()['ok'])
        self.att.refresh_from_db()
        self.assertEqual(self.att.status, Attendance.STATUS_ABSENT)

        # 3. attendance_history POST delete refused
        self.client.post(reverse('food:attendance_history', args=[self.m1.pk]), {'id': self.att.pk})
        self.att.refresh_from_db()
        self.assertFalse(self.att.is_deleted)

        # 4. staff_scores GET lock banner & disabled inputs; blocks manual, paste, import
        resp = self.client.get(reverse('food:staff_scores') + f'?year=2569&main_id={self.main_cat.pk}')
        self.assertEqual(resp.status_code, 200)
        self.assertIn('id="scores-approved-lock-banner"', resp.content.decode('utf-8'))

        self.client.post(reverse('food:staff_scores'), {
            'action': 'save_manual',
            'year': '2569',
            'main_id': self.main_cat.pk,
            f'score_{self.m1.pk}_{self.sub_cat.pk}': '10',
        })
        self.assertEqual(Score.objects.get(user=self.m1, category=self.sub_cat, exam_no=1).score, Decimal('90'))

        self.client.post(reverse('food:staff_scores'), {
            'action': 'paste_excel',
            'year': '2569',
            'main_id': self.main_cat.pk,
            'category': self.sub_cat.pk,
            'orientation': 'vertical',
            'paste_data': '15\n25',
        })
        self.assertEqual(Score.objects.get(user=self.m1, category=self.sub_cat, exam_no=1).score, Decimal('90'))

        xls = make_xlsx(['รหัสค่าย 1', 'สอบปฏิบัติ'], [['69DS01', 12]])
        self.client.post(reverse('food:staff_scores'), {
            'action': 'import_excel',
            'year': '2569',
            'main_id': self.main_cat.pk,
            'excel_file': xls,
        })
        self.assertIsNone(self.client.session.get('score_import_preview'))

    def test_score_categories_locks_and_allows_announcement_text(self):
        cat_url = reverse('food:score_categories')

        # Tab 1 locked actions
        self.client.post(cat_url, {'action': 'create', 'year': '2569', 'name': 'ใหม่', 'weight': '10', 'default_max_score': '10'})
        self.assertFalse(ScoreCategory.objects.filter(academic_year=2569, name='ใหม่').exists())

        self.client.post(cat_url, {'action': 'create_sub', 'year': '2569', 'parent_id': self.main_cat.pk, 'name': 'ย่อยใหม่'})
        self.assertFalse(ScoreCategory.objects.filter(parent=self.main_cat, name='ย่อยใหม่').exists())

        self.client.post(cat_url, {
            'action': 'save_all_categories',
            'year': '2569',
            f'name_{self.main_cat.pk}': 'แก้ชื่อ',
            f'weight_{self.main_cat.pk}': '50',
            f'default_max_score_{self.main_cat.pk}': '50',
            f'order_{self.main_cat.pk}': '1',
        })
        self.main_cat.refresh_from_db()
        self.assertEqual(self.main_cat.name, 'คะแนนสอบรวม')

        self.client.post(cat_url, {'action': 'delete', 'year': '2569', 'id': self.sub_cat.pk, 'confirm_password': '1234'})
        self.assertTrue(ScoreCategory.objects.filter(pk=self.sub_cat.pk).exists())

        # Tab 2 locked actions
        new_d = (self.past_date + dt.timedelta(days=1)).strftime('%Y-%m-%d')
        self.client.post(cat_url, {'action': 'add_training_date', 'year': '2569', 'date_start': new_d, 'event_type': 'training'})
        self.assertFalse(TrainingDate.objects.filter(academic_year=2569, date=self.past_date + dt.timedelta(days=1)).exists())

        resp = self.client.post(
            cat_url,
            {'action': 'update_training_date', 'year': '2569', 'td_id': self.td.pk, 'event_type': TrainingDate.TYPE_HOLIDAY},
            HTTP_X_REQUESTED_WITH='XMLHttpRequest',
        )
        self.assertEqual(resp.status_code, 403)
        self.td.refresh_from_db()
        self.assertEqual(self.td.event_type, TrainingDate.TYPE_TRAINING)

        self.client.post(cat_url, {'action': 'delete_training_date', 'year': '2569', 'td_id': self.td.pk})
        self.assertTrue(TrainingDate.objects.filter(pk=self.td.pk).exists())

        # Tab 3 blocks changing max_qualified_count
        self.client.post(cat_url, {
            'action': 'save_qualification',
            'year': '2569',
            'max_qualified_count': '2',
            'passing_score': '0',
            'announcement_title': 'ควรไม่บันทึกเมื่อส่งเกณฑ์ผิด',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.max_qualified_count, 1)

        # Tab 3 announcement wording WITHOUT modifying max_qualified_count succeeds
        self.client.post(cat_url, {
            'action': 'save_qualification',
            'year': '2569',
            'announcement_title': 'ประกาศผลการคัดเลือกฉบับปรับปรุง',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.announcement_title, 'ประกาศผลการคัดเลือกฉบับปรับปรุง')

        # score_announcement POST save_announcement_text still succeeds, and GET shows frozen results + badge
        self.client.post(reverse('food:score_announcement'), {
            'action': 'save_announcement_text',
            'year': '2569',
            'announcement_school_name': 'โรงเรียนเทพศิรินทร์ (ปรับแก้ข้อความ)',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.announcement_school_name, 'โรงเรียนเทพศิรินทร์ (ปรับแก้ข้อความ)')

        resp = self.client.get(reverse('food:score_announcement') + '?year=2569&only=all')
        self.assertEqual(resp.status_code, 200)
        html = resp.content.decode('utf-8')
        self.assertIn('อนุมัติประกาศผลแล้ว', html)
        self.assertIn(self.m1.camp2_code, html)

    def test_admin_settings_users_trash_and_delete_camp2_locks(self):
        # 1. system_settings blocks max_absences change when approved, allows site_title update
        self.client.post(reverse('composn:admin_settings'), {
            'action': 'save_all_settings',
            'site_title': 'ชื่อระบบใหม่',
            'current_academic_year': '2569',
            'max_rights': '5',
            'max_absences': '9',
            'auto_logout_minutes': '5',
            'display_title': 'หัวข้อทดสอบ',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.max_absences, 5)
        self.assertEqual(self.setting.site_title, 'ชื่อระบบใหม่')

        # 2. user_edit blocks structural changes on approved member, allows student_phone/email edit
        edit_url = reverse('composn:admin_user_edit', args=[self.m1.pk])
        self.client.post(edit_url, {
            'camp1_code': '69DS01',
            'camp2_code': '69KM88',
            'first_name': self.m1.first_name,
            'last_name': self.m1.last_name,
            'academic_year': '2569',
            'current_level': 'ม.4',
            'role': User.ROLE_MEMBER,
            'is_active': 'on',
            'student_phone': '0812345678',
        })
        self.m1.refresh_from_db()
        self.assertEqual(self.m1.camp2_code, '69KM01')

        self.client.post(edit_url, {
            'camp1_code': self.m1.camp1_code,
            'camp2_code': self.m1.camp2_code,
            'first_name': self.m1.first_name,
            'last_name': self.m1.last_name,
            'academic_year': '2569',
            'current_level': 'ม.4',
            'role': User.ROLE_MEMBER,
            'is_active': 'on',
            'student_phone': '0899999999',
            'email': 'updated@example.com',
        })
        self.m1.refresh_from_db()
        self.assertEqual(self.m1.student_phone, '0899999999')
        self.assertEqual(self.m1.email, 'updated@example.com')

        # user_toggle and user_delete blocked
        self.client.post(reverse('composn:admin_user_toggle', args=[self.m1.pk]))
        self.m1.refresh_from_db()
        self.assertTrue(self.m1.is_active)

        self.client.post(reverse('composn:admin_user_delete', args=[self.m1.pk]))
        self.assertTrue(User.objects.filter(pk=self.m1.pk).exists())

        # 3. delete_camp2_view blocked for approved cohort
        self.client.post(reverse('composn:admin_delete_camp2'), {
            'action': 'confirm_delete',
            'admin_password': '1234',
            'confirmed_codes': '69KM01',
        })
        self.m1.refresh_from_db()
        self.assertEqual(self.m1.camp2_code, '69KM01')

        # 4. use_record_delete and archive_record_delete blocked
        self.client.post(reverse('composn:admin_use_record_delete', args=[self.use_rec.pk]))
        self.use_rec.refresh_from_db()
        self.assertFalse(self.use_rec.is_deleted)

        self.client.post(reverse('composn:admin_archive_record_delete', args=[self.arch_rec.pk]))
        self.arch_rec.refresh_from_db()
        self.assertFalse(self.arch_rec.is_deleted)

        # 5. trash restore blocked for approved-year records
        self.client.post(reverse('composn:admin_trash'), {'model': 'attendance', 'id': self.soft_att.pk})
        self.assertTrue(Attendance.all_objects.get(pk=self.soft_att.pk).is_deleted)

        self.client.post(reverse('composn:admin_trash'), {'model': 'use_record', 'id': self.soft_use.pk})
        self.assertTrue(UseRecord.all_objects.get(pk=self.soft_use.pk).is_deleted)

    def test_unapproved_other_year_remains_editable(self):
        # Year 2569 is approved, year 2570 is NOT approved
        self.assertTrue(ScoreApproval.is_year_approved(2569))
        self.assertFalse(ScoreApproval.is_year_approved(2570))

        m_2570 = make_member('70DS01', level='ม.4', year=2570)
        TrainingDate.objects.create(academic_year=2570, date=self.past_date, event_type=TrainingDate.TYPE_TRAINING)
        cat_2570 = ScoreCategory.objects.create(
            name='หมวดปี2570', weight=Decimal('100'), academic_year=2570, default_max_score=Decimal('100'),
        )

        # Services for 2570 succeed
        att_2570, _ = save_attendance(
            m_2570, self.past_date, Attendance.PERIOD_MORNING, Attendance.STATUS_PRESENT,
            recorded_by=self.admin, academic_year=2570,
        )
        self.assertEqual(att_2570.academic_year, 2570)
        sc_2570 = save_score(m_2570, cat_2570, 1, Decimal('85'), Decimal('100'), recorded_by=self.admin)
        self.assertEqual(sc_2570.score, Decimal('85'))

        # Views for 2570 succeed
        self.client.post(reverse('food:score_categories'), {
            'action': 'create',
            'year': '2570',
            'name': 'หมวดใหม่2570',
            'weight': '20',
            'default_max_score': '20',
        })
        self.assertTrue(ScoreCategory.objects.filter(academic_year=2570, name='หมวดใหม่2570').exists())


class MobileTemplatesRenderingTests(FoodBaseTestCase):
    def test_mobile_nav_and_responsive_templates_across_roles(self):
        # 1. Admin mobile navigation & admin responsive pages
        admin_urls = [
            reverse('composn:admin_dashboard'),
            reverse('composn:admin_admins'),
            reverse('composn:admin_admin_create'),
            reverse('composn:admin_users'),
            reverse('composn:admin_user_create'),
            reverse('composn:admin_user_edit', args=[self.m1.pk]),
            reverse('composn:admin_users_settings'),
            reverse('composn:admin_import_students'),
            reverse('composn:admin_import_camp2'),
            reverse('composn:admin_audit_logs'),
            reverse('composn:admin_exports'),
            reverse('composn:admin_rights_settings'),
            reverse('composn:admin_use_records'),
            reverse('composn:admin_backups'),
            reverse('composn:admin_calculate_level'),
            reverse('composn:admin_school_requests'),
            reverse('composn:admin_trash'),
            reverse('composn:admin_rounds'),
            reverse('composn:admin_import_rights'),
            reverse('composn:admin_settings'),
            reverse('composn:admin_delete_camp2'),
            reverse('composn:admin_user_detail', args=[self.m1.pk]),
            reverse('composn:display'),
            reverse('composn:change_password'),
            reverse('food:food_settings'),
            reverse('food:surveys_manage'),
            reverse('food:orders'),
            reverse('food:order_create'),
            reverse('food:admin_certificates') + '?tab=settings',
            reverse('food:signin_sheet'),
            reverse('food:staff_attendance'),
            reverse('food:attendance_history', args=[self.m1.pk]),
            reverse('food:staff_scores'),
            reverse('food:score_announcement'),
            reverse('food:score_categories'),
            reverse('food:attendance_summary'),
            reverse('food:score_summary'),
        ]
        for url in admin_urls:
            resp = self.client.get(url)
            self.assertEqual(resp.status_code, 200, f'Failed loading {url}')
            html = resp.content.decode('utf-8')
            self.assertIn('name="viewport"', html, f'Missing viewport meta on {url}')
            self.assertIn('id="mobileBottomNav"', html, f'Missing #mobileBottomNav on {url}')
            self.assertIn('sidebar-close-btn', html, f'Missing .sidebar-close-btn on {url}')

        # Certificate designer supports touch drag & non-overlapping responsive layout
        cert_html = self.client.get(reverse('food:admin_certificates') + '?tab=settings').content.decode('utf-8')
        self.assertIn('touch-action: none;', cert_html)
        self.assertIn('sticky-xl-top', cert_html)

        # staff_attendance renders mobile cards & view switch
        att_html = self.client.get(reverse('food:staff_attendance')).content.decode('utf-8')
        self.assertIn('id="att-mobile-cards"', att_html)
        self.assertIn('id="att-mobile-view-switch"', att_html)

        # Tables wrapped in .table-responsive
        responsive_urls = [
            reverse('composn:admin_users'),
            reverse('composn:admin_admins'),
            reverse('composn:admin_audit_logs'),
            reverse('composn:admin_use_records'),
            reverse('composn:admin_backups'),
            reverse('composn:admin_calculate_level'),
            reverse('composn:admin_school_requests'),
            reverse('composn:admin_trash'),
            reverse('composn:admin_rounds'),
            reverse('composn:admin_import_rights'),
            reverse('composn:admin_settings'),
            reverse('composn:admin_user_detail', args=[self.m1.pk]),
            reverse('composn:display'),
            reverse('food:attendance_history', args=[self.m1.pk]),
            reverse('food:score_announcement'),
        ]
        for url in responsive_urls:
            html = self.client.get(url).content.decode('utf-8')
            self.assertIn('table-responsive', html, f'Missing .table-responsive on {url}')

        # 2. Staff role mobile bottom nav & all staff routes
        staff = make_staff()
        self.client.force_login(staff)
        for url in [
            reverse('food:staff_attendance'),
            reverse('food:staff_scores'),
            reverse('food:score_announcement'),
            reverse('food:orders'),
            reverse('food:surveys_manage'),
            reverse('composn:admin_users'),
            reverse('composn:change_password'),
        ]:
            resp = self.client.get(url)
            self.assertEqual(resp.status_code, 200, f'Staff failed loading {url}')
            staff_html = resp.content.decode('utf-8')
            self.assertIn('id="mobileBottomNav"', staff_html)
            self.assertIn('sidebar-close-btn', staff_html)

        # 3. Member role mobile bottom nav & all member routes
        self.client.force_login(self.m1)
        for url in [
            reverse('composn:profile'),
            reverse('composn:training_scores'),
            reverse('food:my_food'),
            reverse('food:my_selections'),
            reverse('food:member_surveys'),
            reverse('composn:change_password'),
        ]:
            resp = self.client.get(url)
            self.assertEqual(resp.status_code, 200, f'Member failed loading {url}')
            m_html = resp.content.decode('utf-8')
            self.assertIn('id="mobileBottomNav"', m_html)
            self.assertIn('sidebar-close-btn', m_html)
        self.assertIn(self.client.get(reverse('food:my_certificate')).status_code, (200, 302))

        # 4. Public / unauthenticated pages (login & certificate verification)
        self.client.logout()
        for url in [reverse('composn:login'), reverse('food:certificate_verify_index')]:
            resp = self.client.get(url)
            self.assertEqual(resp.status_code, 200, f'Public page failed loading {url}')
            self.assertIn('name="viewport"', resp.content.decode('utf-8'))


class LoginSecurityAndLockoutTests(FoodBaseTestCase):
    def test_username_lockout_and_admin_unlock(self):
        from django.test import override_settings

        anon = Client()
        login_url = reverse('composn:login')
        with override_settings(LOGIN_MAX_FAILURES_PER_USER=5, LOGIN_MAX_FAILURES_PER_IP=50):
            for _ in range(5):
                r = anon.post(login_url, {'username': '69ds01', 'password': 'wrong'})
                self.assertEqual(r.status_code, 200)

            # 6th attempt even with valid password is blocked with 429
            blocked = anon.post(login_url, {'username': '69DS01', 'password': '69DS01'})
            self.assertEqual(blocked.status_code, 429)
            self.assertTrue(AuditLog.objects.filter(action='login_blocked').exists())

            # Admin sees active lockout & unlocks it
            audit_url = reverse('composn:admin_audit_logs')
            a_resp = self.client.get(audit_url)
            self.assertEqual(a_resp.status_code, 200)
            a_html = a_resp.content.decode('utf-8')
            self.assertIn('id="active-lockouts-card"', a_html)
            self.assertIn('69DS01', a_html)

            self.client.post(audit_url, {
                'action': 'unlock_login',
                'lock_type': 'username',
                'lock_value': '69DS01',
            })

            # Now valid login succeeds
            ok_resp = anon.post(login_url, {'username': '69DS01', 'password': '69DS01'})
            self.assertEqual(ok_resp.status_code, 302)

    def test_ip_lockout_across_multiple_usernames(self):
        from django.test import override_settings

        login_url = reverse('composn:login')
        attacker = Client(REMOTE_ADDR='203.0.113.10')
        normal_user = Client(REMOTE_ADDR='203.0.113.99')
        with override_settings(LOGIN_MAX_FAILURES_PER_USER=10, LOGIN_MAX_FAILURES_PER_IP=4):
            for idx in range(4):
                r = attacker.post(login_url, {'username': f'FAKE{idx:02d}', 'password': 'bad'})
                self.assertEqual(r.status_code, 200)

            # Attacker IP blocked even when switching to a valid account
            blocked = attacker.post(login_url, {'username': '69DS01', 'password': '69DS01'})
            self.assertEqual(blocked.status_code, 429)

            # Different IP is unaffected
            ok_other_ip = normal_user.post(login_url, {'username': '69DS01', 'password': '69DS01'})
            self.assertEqual(ok_other_ip.status_code, 302)

    def test_change_password_enforces_min_length_and_disallows_camp_codes(self):
        self.m1.camp2_code = '69KM01'
        self.m1.save(update_fields=['camp2_code'])
        mc = Client()
        mc.force_login(self.m1)
        chg_url = reverse('composn:change_password')

        # Too short (< 8 chars) rejected
        mc.post(chg_url, {'old_password': '69DS01', 'new_password1': 'abc1234', 'new_password2': 'abc1234'})
        self.m1.refresh_from_db()
        self.assertTrue(self.m1.check_password('69DS01'))

        # Using camp2_code rejected
        mc.post(chg_url, {'old_password': '69DS01', 'new_password1': '69KM01', 'new_password2': '69KM01'})
        self.m1.refresh_from_db()
        self.assertFalse(self.m1.check_password('69KM01'))

        # Strong >= 8 chars accepted
        mc.post(chg_url, {'old_password': '69DS01', 'new_password1': 'SafePass#2569', 'new_password2': 'SafePass#2569'})
        self.m1.refresh_from_db()
        self.assertTrue(self.m1.check_password('SafePass#2569'))

    def test_strict_single_session_across_all_roles(self):
        from django.contrib.sessions.models import Session

        staff = make_staff('staff_s1')
        accounts = [
            ('admin', '1234', reverse('composn:admin_dashboard')),
            ('staff_s1', 'StaffPass!1', reverse('food:staff_attendance')),
            ('69DS01', '69DS01', reverse('composn:profile')),
        ]
        login_url = reverse('composn:login')
        for uname, pw, protected_url in accounts:
            c1, c2 = Client(), Client()
            self.assertEqual(c1.post(login_url, {'username': uname, 'password': pw}).status_code, 302)
            self.assertEqual(c1.get(protected_url).status_code, 200)

            # Second login succeeds and kicks the first session
            self.assertEqual(c2.post(login_url, {'username': uname, 'password': pw}).status_code, 302)
            self.assertEqual(c2.get(protected_url).status_code, 200)

            # First client's session is rejected both for normal requests and AJAX ping
            r1 = c1.get(protected_url)
            self.assertEqual(r1.status_code, 302)
            self.assertIn('/login/', r1['Location'])
            ping1 = c1.get(reverse('composn:session_ping'), HTTP_X_REQUESTED_WITH='XMLHttpRequest')
            self.assertEqual(ping1.status_code, 401)

            # Strictly 1 active DB session exists for this user
            u = User.objects.get(camp1_code=uname)
            active_for_user = [
                s for s in Session.objects.all()
                if str(s.get_decoded().get('_auth_user_id', '')) == str(u.pk)
            ]
            self.assertEqual(len(active_for_user), 1)
            self.assertEqual(active_for_user[0].session_key, u.current_session_key)


class ExamSeatingTests(FoodBaseTestCase):
    """ทดสอบระบบจัดที่นั่งสอบ (Exam Seating Management) ครบทุกฟังก์ชัน"""

    def setUp(self):
        super().setUp()
        # เพิ่มนักเรียน 2569 รวม 12 คน + ผู้ใช้สิทธิ์ 2 คน
        self.extra_members = []
        for idx in range(3, 13):
            self.extra_members.append(
                make_member(f'69DS{idx:02d}', year=2569, first_name=f'นักเรียน{idx}', last_name='ทดสอบ')
            )
        self.all_members = [self.m1, self.m2] + self.extra_members
        self.exam_date = TrainingDate.objects.create(
            academic_year=2569,
            date=dt.date(2569, 10, 13),
            event_type=TrainingDate.TYPE_EXAM,
            exam_period=TrainingDate.EXAM_BOTH,
            camp_level=1,
            title='สอบทฤษฎี',
        )
        self.exam_date2 = TrainingDate.objects.create(
            academic_year=2569,
            date=dt.date(2569, 10, 14),
            event_type=TrainingDate.TYPE_EXAM,
            exam_period=TrainingDate.EXAM_BOTH,
            camp_level=1,
            title='สอบปฏิบัติ',
        )

    def test_default_rooms_seeded_and_menu_placement(self):
        from food.models import ExamRoom, ExamSeat

        # 1. ห้องสอบเริ่มต้น 4 ห้องครบตามสเปก
        r134 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        r135 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 135')
        r136 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 136')
        r_hall = ExamRoom.objects.get(name='ห้องประชุมเทิดพระเกียรติ ชั้น 2')

        self.assertTrue(r134.has_computer)
        self.assertTrue(r135.has_computer)
        self.assertTrue(r136.has_computer)
        self.assertFalse(r_hall.has_computer)

        self.assertEqual(r134.seats.count(), 50)
        self.assertEqual(r135.seats.count(), 55)
        self.assertEqual(r136.seats.count(), 55)
        self.assertEqual(r_hall.seats.count(), 70)

        # ตรวจสอบ 136 ไม่มีเครื่อง 9 และ 10 ตามผังจริง
        labels_136 = set(r136.seats.values_list('seat_label', flat=True))
        self.assertNotIn('9', labels_136)
        self.assertNotIn('10', labels_136)
        self.assertIn('57', labels_136)

        # ตรวจสอบเครื่องสำรองเริ่มต้นของ 134
        reserves_134 = set(r134.seats.filter(default_status=ExamSeat.STATUS_RESERVE).values_list('seat_label', flat=True))
        self.assertEqual(reserves_134, {'1', '16', '41', '49'})

        # 2. ตรวจสอบเมนู "จัดที่นั่งสอบ", "เกียรติบัตร", "การประกาศผล" อยู่ในข้อ 2 (grp-scores) ตามลำดับใหม่
        resp = self.client.get(reverse('composn:admin_dashboard'))
        html = resp.content.decode('utf-8')
        sidebar = html[html.index('id="sidebar"'):html.index('class="main-content"')]
        grp_scores = sidebar[sidebar.index('id="grp-scores"'):sidebar.index('id="grp-food"')]
        self.assertIn('จัดที่นั่งสอบ', grp_scores)
        self.assertLess(grp_scores.index('จัดที่นั่งสอบ'), grp_scores.index('เกียรติบัตร'))
        self.assertLess(grp_scores.index('เกียรติบัตร'), grp_scores.index('การประกาศผล'))

        topbar = html[html.index('class="topbar"'):html.index('class="content-area"')]
        self.assertIn('จัดที่นั่งสอบ', topbar)

    def test_staff_permissions_for_exam_seating(self):
        st_no = make_staff('st_no_seat', staff_can_exam_seating=False)
        st_yes = make_staff('st_yes_seat', staff_can_exam_seating=True)

        c_no = Client()
        c_no.force_login(st_no)
        self.assertEqual(c_no.get(reverse('food:exam_seating')).status_code, 403)
        self.assertEqual(c_no.get(reverse('food:exam_rooms_manage')).status_code, 403)

        c_yes = Client()
        c_yes.force_login(st_yes)
        self.assertEqual(c_yes.get(reverse('food:exam_seating')).status_code, 200)
        self.assertEqual(c_yes.get(reverse('food:exam_rooms_manage')).status_code, 200)

    def test_create_new_room_and_customize_seats(self):
        from food.models import ExamRoom, ExamSeat

        r_create = self.client.post(reverse('food:exam_rooms_manage'), {
            'action': 'create_room',
            'name': 'ห้องปฏิบัติการ 137',
            'has_computer': '1',
            'front_position': 'left',
            'front_label': 'หน้าห้อง',
            'default_fill_pattern': 'seq_asc',
            'default_alternate': '1',
            'sort_order': '5',
            'init_mode': 'empty',
        })
        self.assertEqual(r_create.status_code, 302)
        room = ExamRoom.objects.get(name='ห้องปฏิบัติการ 137')

        # เพิ่มแถวคู่ (double bench) 5 คู่ = 10 เครื่อง
        layout_url = reverse('food:exam_room_layout', args=[room.pk])
        self.client.post(layout_url, {
            'action': 'add_block',
            'bench_type': 'double',
            'count': '5',
            'col_offset': '0',
            'start_no': '1',
            'dir_top': 'lr',
            'dir_bottom': 'rl',
        })
        self.assertEqual(room.seats.count(), 10)

        # เพิ่มที่นั่งเดี่ยวอีก 1 จุด
        self.client.post(layout_url, {
            'action': 'add_seat',
            'block_index': '2',
            'row_index': '3',
            'col_index': '1',
            'seat_label': '11',
            'seat_number': '11',
            'default_status': 'reserve',
        })
        s11 = ExamSeat.objects.get(room=room, seat_label='11')
        self.assertEqual(s11.default_status, ExamSeat.STATUS_RESERVE)

        # แก้ไขและลบที่นั่ง 11
        self.client.post(layout_url, {
            'action': 'update_seat',
            'seat_id': s11.pk,
            'seat_label': '11A',
            'seat_number': '11',
            'default_status': 'normal',
        })
        s11.refresh_from_db()
        self.assertEqual(s11.seat_label, '11A')
        self.client.post(layout_url, {'action': 'delete_seat', 'seat_id': s11.pk})
        self.assertEqual(room.seats.count(), 10)

    def test_lab_alternating_seating_and_reserve_spacing(self):
        from food.exam_seating_service import get_or_create_plan_for_date
        from food.models import ExamRoom, ExamSeatAssignment

        plan = get_or_create_plan_for_date(self.exam_date2, actor=self.admin)
        r134 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        pr134 = plan.plan_rooms.get(room=r134)

        # จัด 8 คนแบบเครื่องเว้นเครื่อง + เว้นข้างเครื่องสำรอง (เครื่อง 1 เป็นสำรอง -> ต้องข้ามเครื่อง 2 ไปลงเครื่อง 3, 5, 7, 9, 11, 13, 15 แล้วไป 17/18)
        target_users = [m.pk for m in self.all_members[:8]]
        resp = self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_room',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_id': str(pr134.pk),
            'pattern': 'seq_asc',
            'alternate': '1',
            'skip_adjacent_to_reserve': '1',
            'user_ids': [str(u) for u in target_users],
        })
        self.assertEqual(resp.status_code, 302)

        assigned = list(
            ExamSeatAssignment.objects.filter(plan_room=pr134, status=ExamSeatAssignment.STATUS_ASSIGNED)
            .select_related('seat')
            .order_by('seat__seat_number')
        )
        self.assertEqual(len(assigned), 8)
        assigned_labels = [a.seat.seat_label for a in assigned]
        # เครื่อง 1 เป็นสำรอง -> ข้าม 2 -> 3, 5, 7, 9, 11, 13, 15 -> เครื่อง 16 เป็นสำรอง -> ข้าม 17 -> ลง 18!
        self.assertEqual(assigned_labels, ['3', '5', '7', '9', '11', '13', '15', '18'])

        # ยืนยันว่าไม่มีผู้เข้าสอบสองคนนั่งติดกันในแถวเดียวกันเลย
        coords = {(a.seat.row_index, a.seat.col_index) for a in assigned}
        for r, c in coords:
            self.assertNotIn((r, c + 1), coords)
            self.assertNotIn((r, c - 1), coords)

    def test_hall_snake_seating_depends_on_start_seat(self):
        from food.exam_seating_service import get_or_create_plan_for_date, order_room_seats
        from food.models import ExamRoom, ExamSeat, ExamSeatAssignment

        r_hall = ExamRoom.objects.get(name='ห้องประชุมเทิดพระเกียรติ ชั้น 2')
        seat_1 = ExamSeat.objects.get(room=r_hall, seat_label='1')
        seat_5 = ExamSeat.objects.get(room=r_hall, seat_label='5')
        seat_51 = ExamSeat.objects.get(room=r_hall, seat_label='51')

        # 1) เริ่มต้นที่ 1 (ซ้าย): แถว 1 ซ้าย->ขวา (1..5), แถว 2 ขวา->ซ้าย (10..6), แถว 3 ซ้าย->ขวา (11..15)
        seq_left = [s.seat_label for s in order_room_seats(r_hall, ExamRoom.FILL_SNAKE_AUTO, start_seat_id=seat_1.pk)[:12]]
        self.assertEqual(seq_left, ['1', '2', '3', '4', '5', '10', '9', '8', '7', '6', '11', '12'])

        # 2) เริ่มต้นที่ 5 (ขวา): แถว 1 ขวา->ซ้าย (5..1), แถว 2 ซ้าย->ขวา (6..10), แถว 3 ขวา->ซ้าย (15..11)
        seq_right = [s.seat_label for s in order_room_seats(r_hall, ExamRoom.FILL_SNAKE_AUTO, start_seat_id=seat_5.pk)[:12]]
        self.assertEqual(seq_right, ['5', '4', '3', '2', '1', '6', '7', '8', '9', '10', '15', '14'])

        # 3) เริ่มต้นที่ 51 (แถว 11 ด้านซ้าย): แถว 11 (51..55), แถว 12 (60..56)
        seq_row11 = [s.seat_label for s in order_room_seats(r_hall, ExamRoom.FILL_SNAKE_AUTO, start_seat_id=seat_51.pk)[:8]]
        self.assertEqual(seq_row11, ['51', '52', '53', '54', '55', '60', '59', '58'])

        # ทดสอบจัดจริงผ่าน view ลงห้องประชุมเทิดพระเกียรติ เริ่มที่นั่ง 5 แบบสลับฟันปลา
        plan = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        pr_hall = plan.plan_rooms.get(room=r_hall)
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_room',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_id': str(pr_hall.pk),
            'pattern': 'snake_auto',
            'start_seat_id': str(seat_5.pk),
            'user_ids': [str(m.pk) for m in self.all_members[:7]],
        })
        # คนที่ 1 (69DS01) ได้ 5, คนที่ 5 ได้ 1, คนที่ 6 ได้ 6, คนที่ 7 ได้ 7
        asg_map = {
            a.seat.seat_label: a.user.camp1_code
            for a in ExamSeatAssignment.objects.filter(plan_room=pr_hall, status=ExamSeatAssignment.STATUS_ASSIGNED).select_related('seat', 'user')
        }
        self.assertEqual(asg_map['5'], '69DS01')
        self.assertEqual(asg_map['1'], '69DS05')
        self.assertEqual(asg_map['6'], '69DS06')
        self.assertEqual(asg_map['7'], '69DS07')

    def test_manual_move_swap_cancel_and_cross_room_and_multi_assign(self):
        from food.exam_seating_service import get_or_create_plan_for_date
        from food.models import ExamRoom, ExamSeat, ExamSeatAssignment

        plan = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        r134 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        r135 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 135')
        pr134 = plan.plan_rooms.get(room=r134)
        pr135 = plan.plan_rooms.get(room=r135)

        s3 = ExamSeat.objects.get(room=r134, seat_label='3')
        s5 = ExamSeat.objects.get(room=r134, seat_label='5')
        s7 = ExamSeat.objects.get(room=r134, seat_label='7')
        s1_135 = ExamSeat.objects.get(room=r135, seat_label='1')

        seat_url = reverse('food:exam_seat_action')
        # จัด 69DS01 -> 134#3, 69DS02 -> 134#5
        self.client.post(seat_url, {'action': 'assign_seat', 'plan_room_id': pr134.pk, 'seat_id': s3.pk, 'user_id': self.m1.pk})
        self.client.post(seat_url, {'action': 'assign_seat', 'plan_room_id': pr134.pk, 'seat_id': s5.pk, 'user_id': self.m2.pk})

        a1 = ExamSeatAssignment.objects.get(plan_room=pr134, seat=s3)
        a2 = ExamSeatAssignment.objects.get(plan_room=pr134, seat=s5)
        self.assertEqual(a1.user, self.m1)
        self.assertEqual(a2.user, self.m2)

        # สลับที่นั่งภายในห้อง (134#3 <-> 134#5)
        r_swap = self.client.post(seat_url, {
            'action': 'move_or_swap',
            'source_assignment_id': a1.pk,
            'target_plan_room_id': pr134.pk,
            'target_seat_id': s5.pk,
        }, HTTP_X_REQUESTED_WITH='XMLHttpRequest')
        self.assertEqual(r_swap.status_code, 200)
        self.assertTrue(r_swap.json()['ok'])
        a1.refresh_from_db()
        a2.refresh_from_db()
        self.assertEqual(a1.user, self.m2)
        self.assertEqual(a2.user, self.m1)

        # ย้ายไปที่นั่งว่างในห้องเดิม (m1 จาก 134#5 -> 134#7)
        self.client.post(seat_url, {
            'action': 'move_or_swap',
            'source_assignment_id': a2.pk,
            'target_plan_room_id': pr134.pk,
            'target_seat_id': s7.pk,
        })
        a2.refresh_from_db()
        self.assertEqual(a2.seat_id, s7.pk)

        # ย้ายข้ามห้อง (m1 จาก 134#7 -> 135#1)
        self.client.post(seat_url, {
            'action': 'move_or_swap',
            'source_assignment_id': a2.pk,
            'target_plan_room_id': pr135.pk,
            'target_seat_id': s1_135.pk,
        })
        a2.refresh_from_db()
        self.assertEqual(a2.plan_room_id, pr135.pk)
        self.assertEqual(a2.seat_id, s1_135.pk)

        # ยกเลิกที่นั่งของ m2
        self.client.post(seat_url, {'action': 'cancel_seat', 'assignment_id': a1.pk})
        self.assertFalse(ExamSeatAssignment.objects.filter(pk=a1.pk).exists())

        # จัดกระจายหลายห้องสำหรับคนที่เหลือ (11 คน) -> ไม่มีซ้ำคน
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_multi',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_ids': [str(pr134.pk), str(pr135.pk)],
            'alternate': '1',
            'skip_adjacent_to_reserve': '1',
        })
        all_assigned = list(ExamSeatAssignment.objects.filter(plan=plan, status=ExamSeatAssignment.STATUS_ASSIGNED))
        self.assertEqual(len(all_assigned), 12)
        self.assertEqual(len({a.user_id for a in all_assigned}), 12)

    def test_proctors_plan_open_close_member_view_and_exports(self):
        from food.exam_seating_service import get_or_create_plan_for_date
        from food.models import CertificateRecipient, ExamRoom, ExamRoomProctor, ExamSeat

        plan = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        r134 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        pr134 = plan.plan_rooms.get(room=r134)
        s3 = ExamSeat.objects.get(room=r134, seat_label='3')

        lec1 = CertificateRecipient.objects.create(
            academic_year=2569, recipient_type=CertificateRecipient.TYPE_LECTURER, full_name='นายเสฏฐวุฒิ โมลานิล',
        )
        lec2 = CertificateRecipient.objects.create(
            academic_year=2569, recipient_type=CertificateRecipient.TYPE_LECTURER, full_name='นางสาวเมธาวี เดชชัย',
        )

        # 1. กำหนดวิทยากรคุมสอบ
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'save_proctors',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_id': str(pr134.pk),
            'recipient_ids': [str(lec1.pk), str(lec2.pk)],
            'name_style': 'short',
            'custom_names': 'ครูศุภพล',
        })
        proctor_names = list(ExamRoomProctor.objects.filter(plan_room=pr134).values_list('display_name', flat=True))
        self.assertEqual(proctor_names, ['ครูเสฏฐวุฒิ', 'ครูเมธาวี', 'ครูศุภพล'])

        # 2. จัดที่นั่งให้ m1 แต่แผนยังไม่เปิด -> นักเรียนต้องยังไม่เห็น
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': pr134.pk, 'seat_id': s3.pk, 'user_id': self.m1.pk,
        })
        mc = Client()
        mc.force_login(self.m1)
        m_resp_closed = mc.get(reverse('food:member_exam_seating')).content.decode('utf-8')
        self.assertIn('ยังไม่มีประกาศที่นั่งสอบในขณะนี้', m_resp_closed)
        prof_closed = mc.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertNotIn('profile-exam-seat-box', prof_closed)

        # 2.1 ยังจัดไม่ครบ (มีแค่ 1/12 คน) -> กดเปิดให้ผู้เข้าสอบดูต้องถูกปฏิเสธ และหน้าจอแสดงปุ่มล็อก
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr134.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)
        seating_incomplete_html = self.client.get(reverse('food:exam_seating') + f'?year=2569&date_id={self.exam_date.pk}').content.decode('utf-8')
        self.assertIn('ยังไม่พร้อมเปิด', seating_incomplete_html)

        # 3. จัดที่นั่งที่เหลืออีก 11 คนให้ครบลงห้อง 135 แล้วเปิดประกาศที่นั่งสอบ -> นักเรียนเห็นทั้งในหน้า profile และหน้า member_exam_seating
        pr135 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 135')
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_multi',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_ids': [str(pr135.pk)],
            'alternate': '1',
        })
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr134.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)
        m_resp_open = mc.get(reverse('food:member_exam_seating')).content.decode('utf-8')
        self.assertIn('ห้องปฏิบัติการคอมพิวเตอร์ 134', m_resp_open)
        self.assertIn('ครูเสฏฐวุฒิ', m_resp_open)
        prof_open = mc.get(reverse('composn:profile')).content.decode('utf-8')
        self.assertIn('profile-exam-seat-box', prof_open)

        # 4. ปิดเฉพาะห้อง 134 ที่มีผู้เข้าสอบอยู่ -> ระบบต้องปิดการเปิดให้ดูอัตโนมัติ และนักเรียนไม่เห็นที่นั่ง
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_room_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr134.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)
        m_resp_room_closed = mc.get(reverse('food:member_exam_seating')).content.decode('utf-8')
        self.assertIn('ยังไม่มีประกาศที่นั่งสอบในขณะนี้', m_resp_room_closed)
        # เปิดห้อง 134 และเปิดแผนกลับเพื่อตรวจ Export
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_room_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr134.pk),
        })
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr134.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)

        # 5. พิมพ์ HTML, Export PDF และ Export Excel
        p_resp = self.client.get(reverse('food:exam_seating_print', args=[plan.pk]))
        self.assertEqual(p_resp.status_code, 200)
        self.assertIn('ห้องปฏิบัติการคอมพิวเตอร์ 134', p_resp.content.decode('utf-8'))

        pdf_resp = self.client.get(reverse('food:exam_seating_export_pdf', args=[plan.pk]))
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertTrue(pdf_resp.content.startswith(b'%PDF-'))

        xlsx_resp = self.client.get(reverse('food:exam_seating_export_xlsx', args=[plan.pk]))
        self.assertEqual(xlsx_resp.status_code, 200)
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(xlsx_resp.content))
        self.assertIn('รายชื่อจัดที่นั่งสอบ', wb.sheetnames)
        self.assertTrue(any('134' in s for s in wb.sheetnames))

        # 6. ยกเลิกที่นั่ง 1 คนขณะที่แผนเปิดอยู่ -> แผนต้องปิดการแสดงผลอัตโนมัติทันที
        cancel_resp = self.client.post(
            reverse('food:exam_seat_action'),
            {'action': 'cancel_seat', 'plan_room_id': str(pr134.pk), 'seat_id': str(s3.pk)},
            HTTP_X_REQUESTED_WITH='XMLHttpRequest',
        )
        self.assertEqual(cancel_resp.status_code, 200)
        self.assertTrue(cancel_resp.json().get('plan_auto_closed'))
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

    def test_exam_type_from_training_date_per_date_open_colors_and_one_page_per_room(self):
        from openpyxl import load_workbook
        from food.exam_seating_service import get_or_create_plan_for_date
        from food.models import ExamRoom, ExamSeat

        # 1. ตั้งค่าประเภทการสอบที่ TrainingDate ผ่านหน้า score_categories -> exam_seating ต้องดึงมาแสดงและไม่มีฟอร์มเลือกประเภทการสอบ
        self.client.post(reverse('food:score_categories'), {
            'action': 'update_training_date',
            'year': '2569',
            'td_id': str(self.exam_date.pk),
            'event_type': TrainingDate.TYPE_EXAM,
            'exam_period': TrainingDate.EXAM_MORNING,
            'exam_type': TrainingDate.EXAM_TYPE_PRACTICE,
            'title': 'รอบเช้า',
        })
        self.exam_date.refresh_from_db()
        self.assertEqual(self.exam_date.exam_type, TrainingDate.EXAM_TYPE_PRACTICE)

        seating_html = self.client.get(reverse('food:exam_seating') + f'?year=2569&date_id={self.exam_date.pk}').content.decode('utf-8')
        self.assertIn('แผนผังที่นั่งสอบปฏิบัติ', seating_html)
        self.assertNotIn('name="exam_type"', seating_html)
        self.assertNotIn('createSessionModal', seating_html)

        # 2. สร้างนักเรียนใช้สิทธิ์ (68DS02) และจัดลงห้อง 134 พร้อมกับนักเรียนปัจจุบัน (69DS01) แล้วจัดที่นั่งที่เหลือลงห้อง 136 / ห้องประชุมให้ครบทุกคน
        r_stu = make_member('68DS02', year=2568, level='ม.3', first_name='ใช้สิทธิ์', last_name='สีเขียว')
        rr = rights.open_round(2569, 1, self.admin)
        rights.confirm_rights(r_stu, rr)

        plan1 = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        plan2 = get_or_create_plan_for_date(self.exam_date2, actor=self.admin)
        r134 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        pr134_1 = plan1.plan_rooms.get(room=r134)
        pr134_2 = plan2.plan_rooms.get(room=r134)

        s3 = ExamSeat.objects.get(room=r134, seat_label='3')
        s5 = ExamSeat.objects.get(room=r134, seat_label='5')
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': pr134_1.pk, 'seat_id': s3.pk, 'user_id': self.m1.pk,
        })
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': pr134_1.pk, 'seat_id': s5.pk, 'user_id': r_stu.pk,
        })
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': pr134_2.pk, 'seat_id': s3.pk, 'user_id': self.m1.pk,
        })

        # ก่อนจัดครบ -> เปิดไม่ได้
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan1.pk), 'date_id': str(self.exam_date.pk),
        })
        plan1.refresh_from_db()
        self.assertFalse(plan1.is_open)

        # จัดที่นั่งที่เหลือลง 136 และห้องประชุม (ปล่อย 135 ว่าง)
        pr136_1 = plan1.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 136')
        pr_hall_1 = plan1.plan_rooms.get(room__name='ห้องประชุมเทิดพระเกียรติ ชั้น 2')
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_multi',
            'year': '2569',
            'plan_id': str(plan1.pk),
            'plan_room_ids': [str(pr136_1.pk), str(pr_hall_1.pk)],
            'alternate': '1',
        })

        # 3. เปิดให้นักเรียนดูเฉพาะวันที่ 1 (plan1) ส่วนวันที่ 2 (plan2) ปิดไว้ -> นักเรียนต้องเห็นเฉพาะวันที่ 1
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan1.pk), 'date_id': str(self.exam_date.pk),
        })
        plan1.refresh_from_db()
        plan2.refresh_from_db()
        self.assertTrue(plan1.is_open)
        self.assertFalse(plan2.is_open)

        mc = Client()
        mc.force_login(self.m1)
        m_html = mc.get(reverse('food:member_exam_seating')).content.decode('utf-8')
        self.assertIn('13 ตุลาคม 2569', m_html)
        self.assertNotIn('14 ตุลาคม 2569', m_html)

        # 4. ตรวจสอบสีตัวอักษรบนหน้าเว็บและหน้าพิมพ์ (น้ำเงิน = นักเรียนปัจจุบัน, เขียวเข้ม = นักเรียนใช้สิทธิ์)
        print_html = self.client.get(reverse('food:exam_seating_print', args=[plan1.pk])).content.decode('utf-8')
        self.assertIn('seat-code-current">69DS01</div>', print_html)
        self.assertIn('seat-code-rights">68DS02</div>', print_html)
        # หน้าพิมพ์ต้องรวมทุกห้องที่เปิด (4 ห้อง)
        self.assertEqual(print_html.count('class="exam-print-page '), 4)

        # 5. PDF รวมทุกห้อง ห้องละ 1 แผ่น A4 (4 ห้องเปิด = 4 หน้า, ปิดห้องว่าง 1 ห้อง = 3 หน้า โดย plan1 ยังเปิดอยู่ได้)
        pdf_4 = self.client.get(reverse('food:exam_seating_export_pdf', args=[plan1.pk])).content
        self.assertTrue(pdf_4.startswith(b'%PDF-'))
        self.assertEqual(len(re.findall(rb'/Type\s*/Page(?!s)\b', pdf_4)), 4)

        r135 = ExamRoom.objects.get(name='ห้องปฏิบัติการคอมพิวเตอร์ 135')
        pr135_1 = plan1.plan_rooms.get(room=r135)
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_room_open', 'year': '2569', 'plan_id': str(plan1.pk), 'plan_room_id': str(pr135_1.pk),
        })
        plan1.refresh_from_db()
        self.assertTrue(plan1.is_open)
        pdf_3 = self.client.get(reverse('food:exam_seating_export_pdf', args=[plan1.pk])).content
        self.assertEqual(len(re.findall(rb'/Type\s*/Page(?!s)\b', pdf_3)), 3)

        # 6. Excel รวมทุกห้องที่เปิด + A4 1x1 page + สีตัวอักษรน้ำเงิน/เขียวเข้ม
        xlsx_bytes = self.client.get(reverse('food:exam_seating_export_xlsx', args=[plan1.pk])).content
        wb = load_workbook(io.BytesIO(xlsx_bytes))
        # ชีตรายชื่อ + 3 ห้องที่เปิดอยู่ (134, 136, ห้องประชุม)
        self.assertEqual(len(wb.sheetnames), 4)
        ws134 = next(wb[s] for s in wb.sheetnames if '134' in s)
        self.assertTrue(ws134.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(ws134.page_setup.fitToWidth, 1)
        self.assertEqual(ws134.page_setup.fitToHeight, 1)
        cell_colors = {}
        for row in ws134.iter_rows():
            for cell in row:
                if cell.value in ('69DS01', '68DS02') and cell.font and cell.font.color:
                    cell_colors[cell.value] = str(cell.font.color.rgb)[-6:].upper()
        self.assertEqual(cell_colors.get('69DS01'), '0000CC')
        self.assertEqual(cell_colors.get('68DS02'), '1B5E20')

    def test_cannot_open_incomplete_plan_and_auto_close(self):
        from food.exam_seating_service import check_plan_ready_to_open, get_or_create_plan_for_date
        from food.models import ExamRoom, ExamSeat

        plan = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        pr134 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        pr135 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 135')

        # 1. ยังไม่จัดใครเลย -> เปิดไม่ได้
        rd0 = check_plan_ready_to_open(plan)
        self.assertFalse(rd0['is_ready'])
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'date_id': str(self.exam_date.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

        # 2. จัดครบทั้งหมด 12 คนลงห้อง 134 -> พร้อมเปิดและเปิดได้
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_multi',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_ids': [str(pr134.pk)],
            'alternate': '1',
        })
        self.assertTrue(check_plan_ready_to_open(plan)['is_ready'])
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'date_id': str(self.exam_date.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)

        # 3. ย้ายคนหนึ่งไปห้อง 135 ที่ปิดอยู่ -> แผนต้องปิดอัตโนมัติ
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_room_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr135.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)  # 135 ยังว่าง แผนยังเปิดอยู่ได้
        one_asgn = plan.assignments.filter(plan_room=pr134, user__isnull=False).first()
        s_135 = ExamSeat.objects.filter(room=pr135.room, default_status=ExamSeat.STATUS_NORMAL).first()
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'move_or_swap',
            'source_assignment_id': str(one_asgn.pk),
            'target_plan_room_id': str(pr135.pk),
            'target_seat_id': str(s_135.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

        # 4. เปิดห้อง 135 กลับ -> พร้อมอีกครั้ง -> เปิดแผน -> ล้างห้อง 135 -> แผนปิดอัตโนมัติ
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_room_open', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr135.pk),
        })
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'date_id': str(self.exam_date.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'clear_room', 'year': '2569', 'plan_id': str(plan.pk), 'plan_room_id': str(pr135.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

        # 5. สั่ง auto_assign ลงห้อง 134 ใหม่ให้ครบ (ทุกคนกลับมาครบ) แล้วเปิดแผน -> clear_plan -> แผนปิดอัตโนมัติ
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_multi',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_ids': [str(pr134.pk)],
            'alternate': '1',
        })
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'toggle_plan_open', 'year': '2569', 'plan_id': str(plan.pk), 'date_id': str(self.exam_date.pk),
        })
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'clear_plan', 'year': '2569', 'plan_id': str(plan.pk),
        })
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

    def test_range_defaults_ui_cleanup_and_pdf_chart_rendering(self):
        from food.exam_seating_export import (
            _calc_page_geometry,
            _compute_uniform_seat_dimensions,
            _ensure_chart_fonts,
        )
        from food.exam_seating_service import (
            build_room_layout_data,
            get_exam_examinees,
            get_or_create_plan_for_date,
        )
        from food.models import ExamRoom, ExamSeat, ExamSeatAssignment

        plan = get_or_create_plan_for_date(self.exam_date, actor=self.admin)
        pr134 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 134')
        pr135 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 135')
        pr136 = plan.plan_rooms.get(room__name='ห้องปฏิบัติการคอมพิวเตอร์ 136')
        pr_hall = plan.plan_rooms.get(room__name='ห้องประชุมเทิดพระเกียรติ ชั้น 2')

        # 1. ตรวจสอบ UI cleanup + ค่าเริ่มต้นลำดับช่วงที่ยังไม่มีใครจัด (1..12)
        resp0 = self.client.get(
            reverse('food:exam_seating') + f'?year=2569&date_id={self.exam_date.pk}&plan_id={plan.pk}&room_id={pr134.pk}'
        )
        html0 = resp0.content.decode('utf-8')
        self.assertIn('เลือกช่วงนี้', html0)
        self.assertNotIn('ติ๊กช่วงนี้', html0)
        self.assertNotIn('toggle-show-names', html0)
        self.assertNotIn('>แสดงชื่อบนผัง<', html0)
        self.assertEqual(resp0.context['default_range_from'], 1)
        self.assertEqual(resp0.context['default_range_to'], 12)

        # 2. จัด 5 คนแรก (ลำดับ 1..5) ลงห้อง 136 แล้วสลับมาห้อง 135 -> ช่วงเริ่มต้นต้องเป็น 6 ถึง 10
        examinees = get_exam_examinees(2569)
        first_five_ids = [str(ex['user_id']) for ex in examinees[:5]]
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'auto_assign_room',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_id': str(pr136.pk),
            'pattern': ExamRoom.FILL_SEQ_ASC,
            'alternate': '1',
            'skip_adjacent_to_reserve': '1',
            'user_ids': first_five_ids,
        })
        resp1 = self.client.get(
            reverse('food:exam_seating') + f'?year=2569&date_id={self.exam_date.pk}&plan_id={plan.pk}&room_id={pr135.pk}'
        )
        self.assertEqual(resp1.context['default_range_from'], 6)
        self.assertEqual(resp1.context['default_range_to'], 10)

        # 3. ยกเลิกคนลำดับที่ 1 ออกจากห้อง 136 -> หน้าจัดห้อง 135 ต้องเริ่มที่ลำดับแรกที่ยังไม่มีที่นั่ง (1)
        asgn_first = ExamSeatAssignment.objects.get(plan=plan, user_id=examinees[0]['user_id'])
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'cancel_seat', 'assignment_id': str(asgn_first.pk),
        })
        resp2 = self.client.get(
            reverse('food:exam_seating') + f'?year=2569&date_id={self.exam_date.pk}&plan_id={plan.pk}&room_id={pr135.pk}'
        )
        self.assertEqual(resp2.context['default_range_from'], 1)

        # 4. ทดสอบ PDF รวม 4 ห้อง พร้อมนักเรียนปัจจุบัน นักเรียนใช้สิทธิ์ ที่นั่งสำรอง งดใช้ และวิทยากรคุมสอบ
        r_stu = make_member('68DS05', year=2568, level='ม.3', first_name='สิทธิ์', last_name='ทดสอบ')
        rr = rights.open_round(2569, 1, self.admin)
        rights.confirm_rights(r_stu, rr)

        s_blocked = ExamSeat.objects.filter(room=pr135.room, default_status=ExamSeat.STATUS_NORMAL).first()
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'set_seat_status', 'plan_room_id': str(pr135.pk), 'seat_id': str(s_blocked.pk), 'status': 'blocked',
        })
        s_hall_1 = ExamSeat.objects.get(room=pr_hall.room, seat_label='1')
        s_hall_2 = ExamSeat.objects.get(room=pr_hall.room, seat_label='2')
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': str(pr_hall.pk), 'seat_id': str(s_hall_1.pk), 'user_id': str(examinees[0]['user_id']),
        })
        self.client.post(reverse('food:exam_seat_action'), {
            'action': 'assign_seat', 'plan_room_id': str(pr_hall.pk), 'seat_id': str(s_hall_2.pk), 'user_id': str(r_stu.pk),
        })
        self.client.post(reverse('food:exam_plan_action'), {
            'action': 'save_proctors',
            'year': '2569',
            'plan_id': str(plan.pk),
            'plan_room_id': str(pr134.pk),
            'name_style': 'short',
            'custom_names': 'ครูเสฏฐวุฒิ\nครูณัฏฐ์กวี\nครูปัทม์',
        })

        from reportlab.lib.units import mm
        reg_f, bold_f, is_pua, scale = _ensure_chart_fonts()
        self.assertEqual(reg_f, 'THSarabunNew')
        self.assertEqual(bold_f, 'THSarabunNew-Bold')
        self.assertTrue(is_pua)

        # ตรวจขนาดกล่องที่นั่งเท่ากันทุกแผ่น โดยพิจารณาจากห้องแถวยาวที่สุด (15 ช่อง) และความสูงแถวห้องประชุม
        entries = []
        for pr in (pr134, pr135, pr136, pr_hall):
            layout = build_room_layout_data(pr.room)
            geom = _calc_page_geometry(pr.room, layout, bool(pr.proctors.exists()))
            entries.append({'plan_room': pr, 'layout': layout, 'geom': geom})
        dims = _compute_uniform_seat_dimensions(entries)
        expected_max_w = entries[0]['geom']['grid_w'] / 15.0
        self.assertAlmostEqual(dims['cell_w'], expected_max_w, places=2)
        hall_geom = entries[3]['geom']
        self.assertLessEqual(14 * dims['seat_unit_h'] + 13 * (2.4 * mm), hall_geom['grid_h'] + 0.01)

        pdf_resp = self.client.get(reverse('food:exam_seating_export_pdf', args=[plan.pk]))
        self.assertEqual(pdf_resp.status_code, 200)
        self.assertTrue(pdf_resp.content.startswith(b'%PDF-'))
        self.assertEqual(len(re.findall(rb'/Type\s*/Page(?!s)\b', pdf_resp.content)), 4)


class MenuReorganizationAndMergedPagesTests(FoodBaseTestCase):
    def test_menu_order_settings_rename_and_merged_pages(self):
        # 1. ตรวจสอบลำดับเมนู เวลาเรียน / คะแนน และ ตั้งค่าระบบ ทั้งใน Sidebar และ Top-nav และไม่มีตัวเลขนำหน้าเมนู
        dash_html = self.client.get(reverse('composn:admin_dashboard')).content.decode('utf-8')
        for numbered in ('1. ผู้ใช้', '2. เวลาเรียน / คะแนน', '3. อาหาร', '4. การใช้สิทธิ์', '5. ตั้งค่าระบบ'):
            self.assertNotIn(numbered, dash_html)
        self.assertIn('calc(0.78rem + 2px)', dash_html)
        self.assertIn('calc(0.84rem + 2px)', dash_html)
        self.assertIn('calc(0.86rem + 2px)', dash_html)

        m_scores = re.search(r'id="grp-scores">(.*?)id="grp-food"', dash_html, re.S)
        self.assertIsNotNone(m_scores)
        grp_scores_html = m_scores.group(1)
        scores_labels = ['บันทึกเวลาเรียน', 'บันทึกคะแนน', 'จัดที่นั่งสอบ', 'เกียรติบัตร', 'การประกาศผล']
        positions_scores = [grp_scores_html.index(lbl) for lbl in scores_labels]
        self.assertEqual(positions_scores, sorted(positions_scores))

        m_settings = re.search(r'id="grp-settings">(.*?)<div class="sidebar-section">บัญชี</div>', dash_html, re.S)
        self.assertIsNotNone(m_settings)
        grp_settings_html = m_settings.group(1)
        settings_labels = ['ผู้ใช้', 'ตั้งค่าสี ข้อความ', 'กำหนดปีการศึกษา', 'รายการอาหาร', 'บันทึกสำรอง']
        positions_settings = [grp_settings_html.index(lbl) for lbl in settings_labels]
        self.assertEqual(positions_settings, sorted(positions_settings))
        self.assertNotIn('เกียรติบัตร', grp_settings_html)
        self.assertNotIn('วันอบรม และคะแนน', grp_settings_html)

        # ตรวจ Top-nav ตั้งค่าระบบ และ เวลาเรียน / คะแนน
        top_nav = dash_html[dash_html.index('class="top-nav-pills"'):dash_html.index('class="topbar-user')]
        m_top_scores = re.search(r'เวลาเรียน / คะแนน.*?</button>\s*<ul class="dropdown-menu[^"]*">(.*?)</ul>', top_nav, re.S)
        self.assertIsNotNone(m_top_scores)
        top_scores_html = m_top_scores.group(1)
        top_scores_pos = [top_scores_html.index(lbl) for lbl in scores_labels]
        self.assertEqual(top_scores_pos, sorted(top_scores_pos))

        m_top_settings = re.search(r'ตั้งค่าระบบ.*?</button>\s*<ul class="dropdown-menu[^"]*">(.*?)</ul>', top_nav, re.S)
        self.assertIsNotNone(m_top_settings)
        top_settings_html = m_top_settings.group(1)
        top_settings_pos = [top_settings_html.index(lbl) for lbl in settings_labels]
        self.assertEqual(top_settings_pos, sorted(top_settings_pos))
        self.assertNotIn('เกียรติบัตร', top_settings_html)

        # 2. หน้า composn:admin_settings เป็น "ตั้งค่าสี ข้อความ" ไม่มี Modal เพิ่มปีการศึกษา / current_academic_year และบันทึกโดยไม่กระทบปีการศึกษา
        set_html = self.client.get(reverse('composn:admin_settings')).content.decode('utf-8')
        self.assertIn('ตั้งค่าสี ข้อความ', set_html)
        self.assertNotIn('id="addAcademicYearModal"', set_html)
        self.assertNotIn('name="current_academic_year"', set_html)

        r_save_theme = self.client.post(reverse('composn:admin_settings'), {
            'action': 'save_all_settings',
            'site_title': 'ทดสอบชื่อใหม่',
            'site_title_line2': 'ศูนย์คอมพิวเตอร์',
            'auto_logout_minutes': '10',
        })
        self.assertEqual(r_save_theme.status_code, 302)
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.site_title, 'ทดสอบชื่อใหม่')
        self.assertEqual(self.setting.current_academic_year, 2569)

        # 3. หน้า food:score_categories เป็น "กำหนดปีการศึกษา" มี 2 แท็บ (ไม่มี tab-qual-btn), จัดการปีการศึกษาและล็อก max_absences หลังอนุมัติ
        cat_html = self.client.get(reverse('food:score_categories')).content.decode('utf-8')
        self.assertIn('กำหนดปีการศึกษา', cat_html)
        self.assertIn('name="current_academic_year"', cat_html)
        self.assertIn('id="addAcademicYearModal"', cat_html)
        self.assertIn('id="tab-scores-btn"', cat_html)
        self.assertIn('id="tab-dates-btn"', cat_html)
        self.assertNotIn('id="tab-qual-btn"', cat_html)

        self.client.post(reverse('food:score_categories'), {
            'action': 'add_academic_year',
            'new_academic_year': '2571',
            'set_as_current': '1',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.current_academic_year, 2571)
        self.assertIn(2571, self.setting.get_academic_years())
        self.client.post(reverse('food:score_categories'), {
            'action': 'save_academic_year_settings',
            'current_academic_year': '2569',
            'max_absences': '4',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.current_academic_year, 2569)
        self.assertEqual(self.setting.max_absences, 4)

        approve_year_scores(2569, self.admin)
        self.client.post(reverse('food:score_categories'), {
            'action': 'save_academic_year_settings',
            'current_academic_year': '2569',
            'max_absences': '9',
        })
        self.setting.refresh_from_db()
        self.assertEqual(self.setting.max_absences, 4)

        # 4. หน้า food:score_announcement แสดง 2 แท็บสำหรับ Admin และไม่อนุญาตให้ Staff เรียกคำสั่งแท็บอนุมัติผล
        ann_admin_html = self.client.get(reverse('food:score_announcement') + '?year=2569&tab=qualification').content.decode('utf-8')
        self.assertIn('id="tab-doc-btn"', ann_admin_html)
        self.assertIn('id="tab-qual-btn"', ann_admin_html)
        self.assertIn('id="approval-status-card"', ann_admin_html)

        staff = make_staff(code='staff_menu')
        sc = Client()
        sc.force_login(staff)
        ann_staff_html = sc.get(reverse('food:score_announcement') + '?year=2569&tab=qualification').content.decode('utf-8')
        self.assertNotIn('id="tab-qual-btn"', ann_staff_html)
        for act in ('approve_scores', 'unapprove_scores', 'save_qualification'):
            r_forb = sc.post(reverse('food:score_announcement'), {'action': act, 'year': '2569'})
            self.assertEqual(r_forb.status_code, 403)

        # 5. หน้า admin_certificates ไม่มีแถบ 2 แท็บนอกที่ซ้ำกับเมนูด้านข้างอีกต่อไป
        cert_html = self.client.get(reverse('food:admin_certificates')).content.decode('utf-8')
        self.assertNotIn('เอกสารประกาศผล</a>', cert_html)


class DashboardQuickTogglesTests(FoodBaseTestCase):
    def test_dashboard_quick_toggles_all_systems(self):
        from food.exam_seating_service import (
            assign_single_seat, ensure_default_exam_rooms, get_or_create_plan_for_date, sync_plan_rooms,
        )
        from food.models import (
            CampSurvey, CertificateYearStatus, ExamSeat, FoodOrder,
        )

        dash_url = reverse('composn:admin_dashboard')

        # 1. สร้างข้อมูลทดสอบ: วันสอบ + แผนสอบ, รายการอาหาร, แบบสอบถาม
        td_exam = TrainingDate.objects.create(
            academic_year=2569,
            date=timezone.localdate() + dt.timedelta(days=4),
            event_type=TrainingDate.TYPE_EXAM,
            exam_type=TrainingDate.EXAM_TYPE_THEORY,
            title='สอบปลายค่าย',
        )
        ensure_default_exam_rooms()
        plan = get_or_create_plan_for_date(td_exam, actor=self.admin)
        plan_rooms = sync_plan_rooms(plan, actor=self.admin)

        order = FoodOrder.objects.create(
            academic_year=2569,
            food_date=timezone.localdate() + dt.timedelta(days=1),
            restaurant_name='ครัวทดสอบ',
            edit_toggle_open=True,
        )
        survey = CampSurvey.objects.create(
            academic_year=2569,
            title='แบบสอบถามความพึงพอใจ',
            is_open=True,
        )

        html = self.client.get(dash_url + '?year=2569').content.decode('utf-8')
        self.assertIn('id="widget-quick_toggles"', html)
        for txt in ('เกียรติบัตร', 'การประกาศผล', 'ระบบเลือกอาหาร', 'แก้ไขข้อมูลส่วนตัว', 'ที่นั่งสอบ', 'รอบการใช้สิทธิ์'):
            self.assertIn(txt, html)

        # 2. เปิด/ปิด เกียรติบัตร
        self.assertFalse(CertificateYearStatus.is_year_published(2569))
        self.client.post(dash_url, {'action': 'dash_toggle_cert_publish', 'year': '2569'})
        self.assertTrue(CertificateYearStatus.is_year_published(2569))
        self.client.post(dash_url, {'action': 'dash_toggle_cert_publish', 'year': '2569'})
        self.assertFalse(CertificateYearStatus.is_year_published(2569))

        # 3. เปิด/ปิด ที่นั่งสอบ (ยังจัดห้องไม่ครบต้องเปิดไม่ได้, จัดครบแล้วเปิด-ปิดได้)
        self.client.post(dash_url, {'action': 'dash_toggle_exam_plan', 'year': '2569', 'plan_id': str(plan.pk)})
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

        open_pr = next((pr for pr in plan_rooms if pr.is_open), plan_rooms[0])
        if not open_pr.is_open:
            open_pr.is_open = True
            open_pr.save()
        normal_seats = list(ExamSeat.objects.filter(
            room=open_pr.room, is_active=True, default_status=ExamSeat.STATUS_NORMAL,
        ).order_by('seat_number')[:2])
        assign_single_seat(open_pr, normal_seats[0], self.m1, actor=self.admin)
        assign_single_seat(open_pr, normal_seats[1], self.m2, actor=self.admin)

        self.client.post(dash_url, {'action': 'dash_toggle_exam_plan', 'year': '2569', 'plan_id': str(plan.pk)})
        plan.refresh_from_db()
        self.assertTrue(plan.is_open)
        self.client.post(dash_url, {'action': 'dash_toggle_exam_plan', 'year': '2569', 'plan_id': str(plan.pk)})
        plan.refresh_from_db()
        self.assertFalse(plan.is_open)

        # 4. เปิด/ปิด ระบบเลือกอาหาร, รายการอาหาร, แบบสอบถาม, แก้ไขข้อมูลส่วนตัว
        self.assertTrue(SystemSetting.get().food_system_open)
        self.client.post(dash_url, {'action': 'dash_toggle_food_system', 'year': '2569'})
        self.assertFalse(SystemSetting.get().food_system_open)
        self.client.post(dash_url, {'action': 'dash_toggle_food_system', 'year': '2569'})
        self.assertTrue(SystemSetting.get().food_system_open)

        self.client.post(dash_url, {'action': 'dash_toggle_food_order', 'year': '2569', 'order_id': str(order.pk)})
        order.refresh_from_db()
        self.assertFalse(order.edit_toggle_open)
        self.client.post(dash_url, {'action': 'dash_toggle_food_order', 'year': '2569', 'order_id': str(order.pk)})
        order.refresh_from_db()
        self.assertTrue(order.edit_toggle_open)

        self.client.post(dash_url, {'action': 'dash_toggle_survey', 'year': '2569', 'survey_id': str(survey.pk)})
        survey.refresh_from_db()
        self.assertFalse(survey.is_open)
        self.client.post(dash_url, {'action': 'dash_toggle_survey', 'year': '2569', 'survey_id': str(survey.pk)})
        survey.refresh_from_db()
        self.assertTrue(survey.is_open)

        self.assertTrue(bool(SystemSetting.get().get_editable_fields()))
        self.client.post(dash_url, {'action': 'dash_toggle_profile_edit_all', 'year': '2569'})
        self.assertEqual(SystemSetting.get().get_editable_fields(), [])
        self.client.post(dash_url, {'action': 'dash_toggle_profile_edit_all', 'year': '2569'})
        self.assertTrue(bool(SystemSetting.get().get_editable_fields()))

        # 5. เปิด/ปิด หัวข้อคะแนน (is_open และ is_visible)
        self.client.post(dash_url, {
            'action': 'dash_toggle_score_category', 'year': '2569',
            'cat_id': str(self.main_cat.pk), 'field': 'is_open',
        })
        self.main_cat.refresh_from_db()
        self.assertFalse(self.main_cat.is_open)
        self.client.post(dash_url, {
            'action': 'dash_toggle_score_category', 'year': '2569',
            'cat_id': str(self.main_cat.pk), 'field': 'is_open',
        })
        self.main_cat.refresh_from_db()
        self.assertTrue(self.main_cat.is_open)

        self.client.post(dash_url, {
            'action': 'dash_toggle_score_category', 'year': '2569',
            'cat_id': str(self.main_cat.pk), 'field': 'is_visible',
        })
        self.main_cat.refresh_from_db()
        self.assertFalse(self.main_cat.is_visible)
        self.client.post(dash_url, {
            'action': 'dash_toggle_score_category', 'year': '2569',
            'cat_id': str(self.main_cat.pk), 'field': 'is_visible',
        })
        self.main_cat.refresh_from_db()
        self.assertTrue(self.main_cat.is_visible)

        # 6. เปิด/ปิด/เปิดใหม่ รอบการใช้สิทธิ์
        self.client.post(dash_url, {'action': 'dash_toggle_rights_round', 'year': '2569', 'camp_level': '1'})
        rr = RightsRound.objects.get(academic_year=2569, camp_level=1)
        self.assertTrue(rr.is_open)
        self.client.post(dash_url, {'action': 'dash_toggle_rights_round', 'year': '2569', 'round_id': str(rr.pk)})
        rr.refresh_from_db()
        self.assertFalse(rr.is_open)
        self.client.post(dash_url, {'action': 'dash_toggle_rights_round', 'year': '2569', 'round_id': str(rr.pk)})
        rr.refresh_from_db()
        self.assertTrue(rr.is_open)

        # 7. อนุมัติ / ยกเลิกการประกาศผล (ต้องใส่รหัสผ่าน Admin ที่ถูกต้อง, และช่วงอนุมัติจะล็อกการสลับหัวข้อคะแนน)
        self.client.post(dash_url, {'action': 'dash_approve_scores', 'year': '2569'})
        self.assertTrue(ScoreApproval.is_year_approved(2569))

        self.client.post(dash_url, {
            'action': 'dash_toggle_score_category', 'year': '2569',
            'cat_id': str(self.main_cat.pk), 'field': 'is_open',
        })
        self.main_cat.refresh_from_db()
        self.assertTrue(self.main_cat.is_open)

        self.client.post(dash_url, {'action': 'dash_unapprove_scores', 'year': '2569', 'confirm_password': 'wrong'})
        self.assertTrue(ScoreApproval.is_year_approved(2569))
        self.client.post(dash_url, {'action': 'dash_unapprove_scores', 'year': '2569', 'confirm_password': '1234'})
        self.assertFalse(ScoreApproval.is_year_approved(2569))

        # 8. Staff ไม่สามารถเข้าถึงหรือสั่งเปิด/ปิดบน Dashboard (403)
        staff = make_staff(code='staff_dash_blk')
        sc = Client()
        sc.force_login(staff)
        self.assertEqual(sc.get(dash_url).status_code, 403)
        self.assertEqual(sc.post(dash_url, {'action': 'dash_toggle_cert_publish', 'year': '2569'}).status_code, 403)



