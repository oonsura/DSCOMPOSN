from django.urls import path

from . import views, views_exam_seating as ves, views_food as vf

app_name = 'food'

urlpatterns = [
    # 3. เวลาเรียน / คะแนน
    path('admin/score-categories/', views.score_categories, name='score_categories'),
    path('admin/reasons/', views.reasons_manage, name='reasons_manage'),
    path('admin/colors/', views.colors_manage, name='colors_manage'),
    path('staff/scores/', views.staff_scores, name='staff_scores'),
    path('staff/scores/summary/', views.score_summary, name='score_summary'),
    path('staff/scores/announcement/', views.score_announcement, name='score_announcement'),
    path('staff/certificates/', views.admin_certificates, name='admin_certificates'),
    path('staff/certificates/download/<int:user_id>/', views.certificate_download, name='certificate_download'),
    path('staff/certificates/export-all/', views.certificates_export_all, name='certificates_export_all'),
    path('staff/certificates/export-zip/', views.certificates_export_zip, name='certificates_export_zip'),
    path('my-certificate/', views.my_certificate, name='my_certificate'),
    path('certificate/verify/', views.certificate_verify, name='certificate_verify_index'),
    path('certificate/verify/<str:ref_code>/', views.certificate_verify, name='certificate_verify'),
    path('admin/certificate-settings/', views.certificate_settings_redirect, name='certificate_settings'),
    path('staff/certificates/preview-pdf/', views.certificate_preview_pdf, name='certificate_preview_pdf'),
    path('staff/certificates/recipient/<int:pk>/download/', views.certificate_recipient_download, name='certificate_recipient_download'),
    path('staff/certificates/recipients/export-pdf/', views.certificate_recipients_export_pdf, name='certificate_recipients_export_pdf'),
    path('staff/certificates/recipients/export-xlsx/', views.certificate_recipients_export_xlsx, name='certificate_recipients_export_xlsx'),
    path('staff/certificates/recipients/template/', views.certificate_recipients_template, name='certificate_recipients_template'),
    path('staff/attendance/', views.staff_attendance, name='staff_attendance'),
    path('staff/attendance/save/', views.attendance_save_api, name='attendance_save'),
    path('staff/attendance/summary/', views.attendance_summary, name='attendance_summary'),
    path('staff/attendance/signin-sheet/', views.signin_sheet, name='signin_sheet'),
    path('staff/attendance/<int:user_id>/', views.attendance_history, name='attendance_history'),

    # 2. จัดที่นั่งสอบ
    path('staff/exam-seating/', ves.exam_seating_index, name='exam_seating'),
    path('staff/exam-seating/plan/action/', ves.exam_plan_action, name='exam_plan_action'),
    path('staff/exam-seating/seat/action/', ves.exam_seat_action, name='exam_seat_action'),
    path('staff/exam-seating/rooms/', ves.exam_rooms_manage, name='exam_rooms_manage'),
    path('staff/exam-seating/rooms/<int:room_id>/layout/', ves.exam_room_layout, name='exam_room_layout'),
    path('staff/exam-seating/print/<int:plan_id>/', ves.exam_seating_print, name='exam_seating_print'),
    path('staff/exam-seating/export/pdf/<int:plan_id>/', ves.exam_seating_export_pdf, name='exam_seating_export_pdf'),
    path('staff/exam-seating/export/xlsx/<int:plan_id>/', ves.exam_seating_export_xlsx, name='exam_seating_export_xlsx'),
    path('exam-seating/', ves.member_exam_seating, name='member_exam_seating'),

    # 3. อาหาร & แบบสอบถามในค่าย
    path('admin/food/settings/', vf.food_settings, name='food_settings'),
    path('staff/food/surveys/', vf.surveys_manage, name='surveys_manage'),
    path('staff/food/', vf.orders, name='orders'),
    path('staff/food/new/', vf.order_create, name='order_create'),
    path('staff/food/<int:pk>/', vf.order_detail, name='order_detail'),
    path('staff/food/<int:pk>/edit/', vf.order_edit, name='order_edit'),
    path('staff/food/<int:pk>/copy/', vf.order_copy, name='order_copy'),
    path('staff/food/<int:pk>/report/', vf.order_report, name='order_report'),
    path('staff/food/<int:pk>/print/', vf.order_print, name='order_print'),
    path('staff/food/<int:pk>/action/', vf.order_action, name='order_action'),
    path('staff/food/<int:pk>/export/', vf.order_export, name='order_export'),
    path('api/ai-food-images/', vf.api_generate_food_images, name='api_ai_food_images'),

    # อาหาร & แบบสอบถามในค่าย (Member)
    path('food/', vf.my_food, name='my_food'),
    path('food/surveys/', vf.member_surveys, name='member_surveys'),
    path('food/my-selections/', vf.my_selections, name='my_selections'),
    path('food/<int:pk>/', vf.my_food_detail, name='my_food_detail'),
]
