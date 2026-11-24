from django.conf import settings

from .models import (
    DEFAULT_ACADEMIC_YEARS, MAX_ACADEMIC_YEAR, MIN_ACADEMIC_YEAR,
    RightsRound, SchoolChangeRequest, SystemSetting, UseRecord, User,
)


def _add_valid_years(target_set, values):
    for raw in values:
        if raw is None:
            continue
        try:
            y = int(raw)
        except (TypeError, ValueError):
            continue
        if MIN_ACADEMIC_YEAR <= y <= MAX_ACADEMIC_YEAR:
            target_set.add(y)


def get_academic_year_choices(cur_year=2569, setting=None):
    """รวบรวมปีการศึกษาที่เปิดใช้งานในระบบ (เริ่ม 2566) และปีที่มีข้อมูลจริงในระบบ พร้อมปีปัจจุบัน"""
    years = set(DEFAULT_ACADEMIC_YEARS)
    try:
        setting_obj = setting if setting is not None else SystemSetting.get()
        if setting_obj is not None:
            _add_valid_years(years, setting_obj.get_academic_years())
    except Exception:
        pass
    _add_valid_years(years, [cur_year])
    try:
        from .models import ArchiveRecord
        from food.models import CampSurvey, CertificateRecipient, ExamSeatingPlan, FoodOrder, ScoreCategory, TrainingDate
        for model_cls in (
            User, UseRecord, RightsRound, ArchiveRecord,
            TrainingDate, ScoreCategory, CampSurvey, FoodOrder, CertificateRecipient, ExamSeatingPlan,
        ):
            _add_valid_years(
                years,
                model_cls.objects.exclude(academic_year__isnull=True).values_list('academic_year', flat=True).distinct(),
            )
    except Exception:
        pass
    return sorted(years, reverse=True)


def site_context(request):
    from django.utils import timezone
    ctx = {
        'AUTO_LOGOUT_SECONDS': settings.AUTO_LOGOUT_SECONDS,
        'AUTO_LOGOUT_MINUTES': max(0, int(settings.AUTO_LOGOUT_SECONDS // 60)),
        'pending_food_orders': [],
        'pending_food_orders_count': 0,
        'pending_surveys': [],
        'pending_surveys_count': 0,
        'member_training_years': [],
        'member_open_exam_seats_count': 0,
    }
    setting = None
    try:
        setting = SystemSetting.get()
        ctx['site_setting'] = setting
        mins = max(0, int(getattr(setting, 'auto_logout_minutes', 5)))
        ctx['AUTO_LOGOUT_MINUTES'] = mins
        ctx['AUTO_LOGOUT_SECONDS'] = mins * 60
        ctx['academic_year_choices'] = get_academic_year_choices(setting.current_academic_year, setting=setting)
    except Exception:  # ตารางยังไม่ถูกสร้าง (ก่อน migrate)
        ctx['site_setting'] = None
        ctx['academic_year_choices'] = sorted(DEFAULT_ACADEMIC_YEARS, reverse=True)
    user = getattr(request, 'user', None)
    if user is not None and user.is_authenticated:
        if user.is_admin_role:
            try:
                ctx['pending_school_requests'] = SchoolChangeRequest.objects.filter(
                    status=SchoolChangeRequest.STATUS_PENDING).count()
            except Exception:
                ctx['pending_school_requests'] = 0
        elif user.role == User.ROLE_MEMBER:
            if setting:
                base_y = user.base_academic_year
                ctx['is_current_member'] = bool(base_y and int(base_y) >= int(setting.current_academic_year))
            try:
                cohort_y = user.base_academic_year
                used_map = {
                    int(r.academic_year): r.camp_level
                    for r in UseRecord.objects.filter(user=user, is_used=True).only('academic_year', 'camp_level')
                }
                yr_set = set(used_map.keys())
                if cohort_y:
                    yr_set.add(int(cohort_y))
                ctx['member_training_years'] = [
                    {
                        'year': y,
                        'is_training_year': bool(cohort_y and int(y) == int(cohort_y)),
                        'camp_level': 1 if (cohort_y and int(y) == int(cohort_y)) else used_map.get(y, 1),
                    }
                    for y in sorted(yr_set, reverse=True)
                ]
            except Exception:
                ctx['member_training_years'] = []
            if setting and setting.food_system_open:
                try:
                    from food.models import FoodOrder, FoodSelection
                    cur_year = int(setting.current_academic_year)
                    today = timezone.localdate()
                    open_orders = list(
                        FoodOrder.objects.filter(
                            academic_year=cur_year,
                            edit_toggle_open=True,
                            food_date__gte=today,
                        ).order_by('food_date', 'id')
                    )
                    if open_orders:
                        is_rights = UseRecord.objects.filter(
                            user=user, academic_year=cur_year, is_used=True
                        ).exists()
                        visible = [o for o in open_orders if o.is_visible_to(user, year=cur_year, is_rights_user=is_rights)]
                        if visible:
                            chosen_ids = set(
                                FoodSelection.objects.filter(user=user, food_order__in=visible)
                                .values_list('food_order_id', flat=True)
                            )
                            pending = [o for o in visible if o.pk not in chosen_ids]
                            ctx['pending_food_orders'] = pending
                            ctx['pending_food_orders_count'] = len(pending)
                except Exception:
                    pass
            if setting:
                try:
                    from food.models import CampSurvey, CampSurveyAnswer
                    cur_year = int(setting.current_academic_year)
                    open_surveys = list(CampSurvey.objects.filter(academic_year=cur_year, is_open=True))
                    if open_surveys:
                        is_rights_sv = UseRecord.objects.filter(
                            user=user, academic_year=cur_year, is_used=True
                        ).exists()
                        visible_surveys = [
                            s for s in open_surveys
                            if s.is_visible_to(user, year=cur_year, is_rights_user=is_rights_sv)
                        ]
                        if visible_surveys:
                            answered_ids = set(
                                CampSurveyAnswer.objects.filter(user=user, survey__in=visible_surveys)
                                .values_list('survey_id', flat=True)
                            )
                            pending_sv = [s for s in visible_surveys if s.pk not in answered_ids]
                            ctx['pending_surveys'] = pending_sv
                            ctx['pending_surveys_count'] = len(pending_sv)
                except Exception:
                    pass
            try:
                from food.models import ExamSeatAssignment
                ctx['member_open_exam_seats_count'] = ExamSeatAssignment.objects.filter(
                    user=user,
                    status=ExamSeatAssignment.STATUS_ASSIGNED,
                    plan__is_open=True,
                    plan_room__is_open=True,
                ).count()
            except Exception:
                ctx['member_open_exam_seats_count'] = 0
    return ctx
