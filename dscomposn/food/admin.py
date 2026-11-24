from django.contrib import admin

from composn.admin import SoftDeleteAdmin

from .models import (AbsenceReason, Attendance, CertificateRecipient, CertificateSetting,
                     ExamRoom, ExamRoomProctor, ExamSeat, ExamSeatAssignment, ExamSeatingPlan,
                     ExamSeatingPlanRoom, FoodMenu, FoodOrder, FoodSelection, Score, ScoreCategory,
                     WarningColor)


@admin.register(Attendance)
class AttendanceAdmin(SoftDeleteAdmin):
    list_display = ('user', 'date', 'session_period', 'status', 'reason', 'is_deleted')
    search_fields = ('user__camp1_code',)


@admin.register(Score)
class ScoreAdmin(SoftDeleteAdmin):
    list_display = ('user', 'category', 'exam_no', 'score', 'max_score', 'is_deleted')


admin.site.register([
    AbsenceReason, WarningColor, ScoreCategory, FoodOrder, FoodMenu, FoodSelection,
    CertificateSetting, CertificateRecipient,
    ExamRoom, ExamSeat, ExamSeatingPlan, ExamSeatingPlanRoom, ExamRoomProctor, ExamSeatAssignment,
])
