from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin

from .models import (ArchiveRecord, AuditLog, RightsRound, SchoolChangeRequest, SystemSetting, UseRecord, User)


@admin.register(User)
class UserAdmin(BaseUserAdmin):
    list_display = ('camp1_code', 'camp2_code', 'first_name', 'last_name', 'camp_center', 'role', 'is_active')
    list_filter = ('role', 'camp_center', 'is_active')
    search_fields = ('camp1_code', 'camp2_code', 'first_name', 'last_name', 'school_name')
    ordering = ('camp1_code',)
    fieldsets = (
        (None, {'fields': ('camp1_code', 'password')}),
        ('ข้อมูลส่วนตัว', {'fields': ('camp2_code', 'first_name', 'last_name', 'school_name', 'birth_date',
                                      'camp_center', 'academic_year', 'note')}),
        ('ระดับชั้น', {'fields': ('current_level', 'year_plus_1', 'year_plus_2', 'year_plus_3', 'year_plus_4',
                                  'year_plus_5')}),
        ('สิทธิ์', {'fields': ('role', 'is_active', 'is_first_login', 'is_superuser', 'groups',
                               'user_permissions')}),
    )
    add_fieldsets = ((None, {'classes': ('wide',), 'fields': ('camp1_code', 'role', 'password1', 'password2')}),)


class SoftDeleteAdmin(admin.ModelAdmin):
    def get_queryset(self, request):
        return self.model.all_objects.get_queryset()

    list_filter = ('is_deleted',)


@admin.register(UseRecord)
class UseRecordAdmin(SoftDeleteAdmin):
    list_display = ('rights_code', 'user', 'academic_year', 'camp_level', 'is_used', 'confirmed_at', 'is_deleted')
    search_fields = ('rights_code', 'user__camp1_code')


@admin.register(ArchiveRecord)
class ArchiveRecordAdmin(SoftDeleteAdmin):
    list_display = ('rights_code', 'user', 'academic_year', 'camp_level', 'used_at', 'is_deleted')
    search_fields = ('rights_code', 'user__camp1_code')


@admin.register(RightsRound)
class RightsRoundAdmin(admin.ModelAdmin):
    list_display = ('__str__', 'is_open', 'opened_at', 'closed_at')


@admin.register(SchoolChangeRequest)
class SchoolChangeRequestAdmin(admin.ModelAdmin):
    list_display = ('user', 'new_school_name', 'status', 'is_cleared', 'created_at')


@admin.register(AuditLog)
class AuditLogAdmin(admin.ModelAdmin):
    list_display = ('created_at', 'user', 'action', 'target_model', 'target_id')
    list_filter = ('action',)
    search_fields = ('user__camp1_code', 'new_value')

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


admin.site.register(SystemSetting)
