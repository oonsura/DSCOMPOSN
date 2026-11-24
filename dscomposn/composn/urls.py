from django.urls import path

from . import views, views_admin as va

app_name = 'composn'

urlpatterns = [
    path('', views.home, name='home'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('change-password/', views.change_password, name='change_password'),
    path('session-ping/', views.session_ping, name='session_ping'),

    # Member
    path('profile/', views.profile, name='profile'),
    path('training/', views.training_scores, name='training_scores'),
    path('training/<int:year>/', views.training_scores, name='training_scores_year'),
    path('confirm-rights/', views.confirm_rights, name='confirm_rights'),
    path('cancel-rights/', views.cancel_rights, name='cancel_rights'),
    path('school-request/', views.school_request, name='school_request'),
    path('profile/history.xlsx', views.my_history_export, name='my_history_export'),

    # Display (ดูได้หลัง Login)
    path('display/', views.display, name='display'),
    path('display/download/', views.display_download, name='display_download'),
    path('display/<str:center>/', views.display, name='display_center'),

    # Admin - 1. แดชบอร์ด
    path('admin/', va.dashboard, name='admin_dashboard'),

    # Admin - 2. ผู้ใช้
    path('admin/users/settings/', va.users_settings, name='admin_users_settings'),
    path('admin/admins/', va.admins_list, name='admin_admins'),
    path('admin/admins/new/', va.admin_create, name='admin_admin_create'),
    path('admin/users/', va.users, name='admin_users'),
    path('admin/users/new/', va.user_create, name='admin_user_create'),
    path('admin/users/bulk-action/', va.users_bulk_action, name='admin_users_bulk_action'),
    path('admin/users/<int:pk>/', va.user_detail, name='admin_user_detail'),
    path('admin/users/<int:pk>/edit/', va.user_edit, name='admin_user_edit'),
    path('admin/users/<int:pk>/delete/', va.user_delete, name='admin_user_delete'),
    path('admin/users/<int:pk>/toggle/', va.user_toggle, name='admin_user_toggle'),
    path('admin/users/<int:pk>/reset-password/', va.user_reset_password, name='admin_user_reset_password'),
    path('admin/users/<int:pk>/set-password/', va.user_set_password, name='admin_user_set_password'),
    path('admin/import-students/', va.import_students_view, name='admin_import_students'),
    path('admin/import-camp2/', va.import_camp2_view, name='admin_import_camp2'),
    path('admin/delete-camp2/', va.delete_camp2_view, name='admin_delete_camp2'),
    path('admin/templates/students.xlsx', va.template_students, name='admin_template_students'),
    path('admin/templates/camp2.xlsx', va.template_camp2, name='admin_template_camp2'),
    path('admin/calculate-level/', va.calculate_level, name='admin_calculate_level'),
    path('admin/school-requests/', va.school_requests, name='admin_school_requests'),

    # Admin - 4. การใช้สิทธิ์
    path('admin/rights/settings/', va.rights_settings, name='admin_rights_settings'),
    path('admin/import-rights/', va.import_rights_users_view, name='admin_import_rights'),
    path('admin/templates/rights.xlsx', va.template_rights, name='admin_template_rights'),
    path('admin/rounds/', va.rounds, name='admin_rounds'),
    path('admin/open-round/', va.open_round, name='admin_open_round'),
    path('admin/close-round/<int:pk>/', va.close_round, name='admin_close_round'),
    path('admin/reopen-round/<int:pk>/', va.reopen_round, name='admin_reopen_round'),
    path('admin/rounds/<int:pk>/schedule/', va.round_edit_schedule, name='admin_round_schedule'),
    path('admin/rounds/<int:pk>/', va.round_detail, name='admin_round_detail'),
    path('admin/rounds/report/', va.year_report, name='admin_year_report'),
    path('admin/use-records/', va.use_records_list, name='admin_use_records'),
    path('admin/use-records/<int:pk>/delete/', va.use_record_delete, name='admin_use_record_delete'),
    path('admin/archive-records/<int:pk>/delete/', va.archive_record_delete, name='admin_archive_record_delete'),

    # Admin - 6. ตั้งค่าระบบ & Backup / Audit
    path('admin/settings/', va.system_settings, name='admin_settings'),
    path('admin/backups/', va.backups, name='admin_backups'),
    path('admin/backups/download/<str:name>', va.backup_download, name='admin_backup_download'),
    path('admin/exports/', va.exports, name='admin_exports'),
    path('admin/audit-logs/', va.audit_logs, name='admin_audit_logs'),
    path('admin/trash/', va.trash, name='admin_trash'),
]
