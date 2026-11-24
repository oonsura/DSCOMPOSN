"""
Views ระบบสั่งอาหาร (รวมจากระบบ food เดิมอย่างครบถ้วน)
- Admin/Staff: สร้าง/แก้ไข/คัดลอก/ลบรายการอาหาร, จัดการเมนู (พร้อมค้นหารูป AI + แก้ไขเมนู),
               ดูสรุปยอด, พิมพ์รายงาน A4 (5 ชื่อต่อบรรทัด), Export Excel
- Member: เลือกเมนูตามเงื่อนไข (จำนวนสูงสุด, ไม่ซ้ำ/ซ้ำได้, ขอบเขตศูนย์), ดูรายการที่ตนเองเลือก
"""
import datetime as dt
import ssl
import urllib.parse
import urllib.request

from django import forms
from django.contrib import messages
from django.core.files.base import ContentFile
from django.db import transaction
from django.db.models import Count, Q, Sum
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from composn.audit import log_action
from composn.decorators import admin_required, member_required, staff_perm_required
from composn.excel import build_workbook, workbook_response
from composn.forms import BootstrapMixin
from composn.models import SystemSetting, User
from composn.templatetags.composn_tags import format_date_thai_full, get_thai_day_name

from .models import (CampSurvey, CampSurveyAnswer, CampSurveyQuestion, FoodMenu, FoodOrder, FoodSelection,
                     build_survey_summary)


@admin_required
def food_settings(request):
    """ตั้งค่าหน้าอาหาร (เปิด/ปิดระบบเลือกอาหารสำหรับสมาชิก และตั้งค่าสีการ์ดอาหารประจำวัน)"""
    setting = SystemSetting.get()
    if request.method == 'POST':
        action = request.POST.get('action', '')
        if action == 'save_all_food_settings':
            setting.food_system_open = bool(request.POST.get('food_system_open'))
            setting.food_closed_message = request.POST.get('food_closed_message', setting.food_closed_message).strip()
            for f in ('color_monday', 'color_tuesday', 'color_wednesday', 'color_thursday',
                      'color_friday', 'color_saturday', 'color_sunday'):
                if f in request.POST:
                    setattr(setting, f, request.POST.get(f, '').strip())
            setting.save()
            log_action(request.user, 'update_settings', target=setting,
                       new_value='บันทึกการตั้งค่าหน้าอาหารทั้งหมด', request=request)
            messages.success(request, 'บันทึกการตั้งค่าหน้าอาหารเรียบร้อย')
        elif action == 'toggle_food':
            setting.food_system_open = not setting.food_system_open
            setting.food_closed_message = request.POST.get('food_closed_message', setting.food_closed_message).strip()
            setting.save(update_fields=['food_system_open', 'food_closed_message', 'updated_at'])
            log_action(request.user, 'update_settings', target=setting,
                       new_value={'food_system_open': setting.food_system_open}, request=request)
            messages.success(request, f'{"เปิด" if setting.food_system_open else "ปิด"}ระบบเลือกอาหารเรียบร้อย')
        elif action == 'save_colors':
            for f in ('color_monday', 'color_tuesday', 'color_wednesday', 'color_thursday',
                      'color_friday', 'color_saturday', 'color_sunday'):
                if f in request.POST:
                    setattr(setting, f, request.POST.get(f, '').strip())
            setting.save()
            log_action(request.user, 'update_settings', target=setting,
                       new_value='อัปเดตสีการ์ดอาหารตามวัน', request=request)
            messages.success(request, 'บันทึกสีการ์ดรายการอาหารตามวันในสัปดาห์เรียบร้อย')
        return redirect('food:food_settings')
    return render(request, 'food/food_settings.html', {'setting': setting})


class FoodOrderForm(BootstrapMixin, forms.ModelForm):
    class Meta:
        model = FoodOrder
        fields = ['food_date', 'restaurant_name', 'max_selection_count', 'selection_mode',
                  'allow_current_students', 'allow_rights_students',
                  'edit_toggle_open', 'overall_image', 'additional_detail']
        widgets = {'food_date': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
                   'additional_detail': forms.Textarea(attrs={'rows': 2})}

    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self._bootstrap()

    def clean_max_selection_count(self):
        v = self.cleaned_data['max_selection_count']
        if v < 1:
            raise forms.ValidationError('ต้องมากกว่า 0')
        return v

    def clean(self):
        cleaned = super().clean()
        if not cleaned.get('allow_current_students') and not cleaned.get('allow_rights_students'):
            raise forms.ValidationError('กรุณาเลือกกลุ่มผู้มีสิทธิ์สั่งอาหารอย่างน้อย 1 กลุ่ม (นักเรียนปัจจุบัน หรือ นักเรียนใช้สิทธิ์)')
        return cleaned


class FoodMenuForm(BootstrapMixin, forms.ModelForm):
    class Meta:
        model = FoodMenu
        fields = ['menu_name', 'menu_description', 'menu_image', 'sort_order']
        widgets = {'menu_description': forms.Textarea(attrs={'rows': 2})}

    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self._bootstrap()


def _download_image_from_url(url, timeout=15):
    if not url:
        return None
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(
            url,
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
            },
        )
        with urllib.request.urlopen(req, context=ctx, timeout=timeout) as resp:
            return resp.read()
    except Exception:
        return None


def _decorate_order(o, setting=None):
    setting = setting or SystemSetting.get()
    o.day_color = setting.get_day_color(o.food_date.weekday())
    o.day_name = get_thai_day_name(o.food_date.weekday())
    o.date_full = format_date_thai_full(o.food_date)
    return o


def _attach_order_selection_counts(orders_list):
    """แนบจำนวนนักเรียนที่เลือกแล้ว (selected_student_count) และจำนวนนักเรียนทั้งหมดที่มีสิทธิ์ (eligible_student_count)"""
    if not orders_list:
        return
    from composn.models import UseRecord
    order_ids = [o.pk for o in orders_list]
    chosen_map = dict(
        FoodSelection.objects.filter(food_order_id__in=order_ids, user_id__isnull=False)
        .values('food_order_id')
        .annotate(n=Count('user_id', distinct=True))
        .values_list('food_order_id', 'n')
    )
    elig_cache = {}
    rights_cache = {}
    for o in orders_list:
        o.selected_student_count = chosen_map.get(o.pk, 0)
        key = (int(o.academic_year), o.scope, bool(o.allow_current_students), bool(o.allow_rights_students))
        if key not in elig_cache:
            y, sc, allow_cur, allow_r = key
            if not allow_cur and not allow_r:
                elig_cache[key] = 0
            else:
                qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True)
                if sc in ('DS', 'KM'):
                    qs = qs.filter(camp_center=sc)
                cond = Q()
                if allow_cur:
                    yy = f'{y % 100:02d}'
                    cond |= Q(camp1_code__istartswith=yy) | Q(academic_year=y)
                if allow_r:
                    if y not in rights_cache:
                        rights_cache[y] = set(
                            UseRecord.objects.filter(academic_year=y, is_used=True).values_list('user_id', flat=True)
                        )
                    cond |= Q(pk__in=rights_cache[y])
                elig_cache[key] = qs.filter(cond).distinct().count()
        o.eligible_student_count = elig_cache[key]


# ------------------------------- Staff/Admin ---------------------------------
@staff_perm_required('staff_can_manage_food')
def orders(request):
    setting = SystemSetting.get()
    year_req = request.GET.get('year', '').strip()
    try:
        year = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year = int(setting.current_academic_year)
    qs = FoodOrder.objects.prefetch_related('menus').filter(academic_year=year).order_by('food_date', 'id')
    q = request.GET.get('q', '').strip()
    scope = request.GET.get('scope', '')
    flt = request.GET.get('status', '')
    if q:
        qs = qs.filter(restaurant_name__icontains=q)
    if scope in ('all', 'DS', 'KM'):
        qs = qs.filter(scope=scope)
    items = [_decorate_order(o, setting) for o in qs if not flt or o.get_status() == flt]
    _attach_order_selection_counts(items)
    return render(request, 'food/orders.html', {
        'orders': items, 'flt': flt, 'q': q, 'scope': scope, 'year': year,
    })


@staff_perm_required('staff_can_manage_food')
def order_create(request):
    setting = SystemSetting.get()
    form = FoodOrderForm(request.POST or None, request.FILES or None,
                         initial={'food_date': timezone.localdate()})
    existing_restaurants = list(
        FoodOrder.objects.values_list('restaurant_name', flat=True).distinct().order_by('restaurant_name')
    )
    if request.method == 'POST' and form.is_valid():
        o = form.save(commit=False)
        o.academic_year = setting.current_academic_year
        o.created_by = request.user
        o.save()
        log_action(request.user, 'food_order', target=o, new_value={'action': 'create', 'name': o.restaurant_name,
                                                                    'date': str(o.food_date)}, request=request)
        messages.success(request, 'สร้างรายการอาหารแล้ว เพิ่มเมนูได้ที่หน้านี้')
        return redirect('food:order_detail', pk=o.pk)
    return render(request, 'food/order_form.html', {
        'form': form, 'title': 'สร้างรายการอาหาร', 'existing_restaurants': existing_restaurants,
    })


@staff_perm_required('staff_can_manage_food')
def order_edit(request, pk):
    o = get_object_or_404(FoodOrder, pk=pk)
    form = FoodOrderForm(request.POST or None, request.FILES or None, instance=o)
    existing_restaurants = list(
        FoodOrder.objects.values_list('restaurant_name', flat=True).distinct().order_by('restaurant_name')
    )
    if request.method == 'POST' and form.is_valid():
        form.save()
        log_action(request.user, 'food_order', target=o, new_value={'action': 'edit'}, request=request)
        messages.success(request, 'บันทึกการแก้ไขรายการอาหารแล้ว')
        return redirect('food:order_detail', pk=o.pk)
    return render(request, 'food/order_form.html', {
        'form': form, 'title': f'แก้ไข {o.restaurant_name}', 'o': o,
        'existing_restaurants': existing_restaurants,
    })


@staff_perm_required('staff_can_manage_food')
def order_copy(request, pk):
    """คัดลอกรายการอาหารพร้อมเมนูทั้งหมดเป็นรายการใหม่"""
    orig = get_object_or_404(FoodOrder, pk=pk)
    initial = {
        'food_date': timezone.localdate(),
        'restaurant_name': orig.restaurant_name,
        'scope': orig.scope,
        'allow_current_students': orig.allow_current_students,
        'allow_rights_students': orig.allow_rights_students,
        'max_selection_count': orig.max_selection_count,
        'selection_mode': orig.selection_mode,
        'edit_toggle_open': True,
        'additional_detail': orig.additional_detail,
    }
    form = FoodOrderForm(request.POST or None, request.FILES or None, initial=initial)
    if request.method == 'POST' and form.is_valid():
        with transaction.atomic():
            new_o = form.save(commit=False)
            new_o.academic_year = SystemSetting.get().current_academic_year
            new_o.created_by = request.user
            if not new_o.overall_image and orig.overall_image:
                new_o.overall_image = orig.overall_image
            new_o.save()
            for m in orig.menus.all():
                FoodMenu.objects.create(
                    food_order=new_o,
                    menu_name=m.menu_name,
                    menu_description=m.menu_description,
                    menu_image=m.menu_image,
                    sort_order=m.sort_order,
                )
        log_action(request.user, 'food_order', target=new_o,
                   new_value={'action': 'copy_from', 'source_id': orig.pk, 'name': new_o.restaurant_name},
                   request=request)
        messages.success(request, f'คัดลอกรายการอาหารพร้อมเมนู {orig.menus.count()} รายการเรียบร้อย')
        return redirect('food:order_detail', pk=new_o.pk)
    return render(request, 'food/order_copy.html', {'form': form, 'original': orig})


@staff_perm_required('staff_can_manage_food')
def order_detail(request, pk):
    o = _decorate_order(get_object_or_404(FoodOrder, pk=pk))
    _attach_order_selection_counts([o])
    center_filter = request.GET.get('center', '').upper()
    if center_filter not in ('DS', 'KM'):
        center_filter = ''

    menu_form = FoodMenuForm(request.POST or None, request.FILES or None, prefix='m',
                             initial={'sort_order': o.menus.count() + 1})
    if request.method == 'POST' and request.POST.get('action') == 'add_menu':
        # รองรับทั้งแบบใช้ prefix 'm-' และแบบส่งตรงจากฟอร์มใหม่
        m_name = (request.POST.get('m-menu_name') or request.POST.get('menu_name') or '').strip()
        m_desc = (request.POST.get('m-menu_description') or request.POST.get('menu_description') or '').strip()
        try:
            m_sort = int(request.POST.get('m-sort_order') or request.POST.get('sort_order') or (o.menus.count() + 1))
        except ValueError:
            m_sort = o.menus.count() + 1
        if not m_name:
            messages.error(request, 'กรุณากรอกชื่อเมนู')
        else:
            m = FoodMenu(food_order=o, menu_name=m_name, menu_description=m_desc, sort_order=m_sort)
            img_file = request.FILES.get('m-menu_image') or request.FILES.get('menu_image')
            ai_url = (request.POST.get('ai_image_url') or '').strip()
            if img_file:
                m.menu_image = img_file
            elif ai_url:
                data = _download_image_from_url(ai_url)
                if data:
                    m.menu_image.save(f'ai_menu_{o.pk}_{int(timezone.now().timestamp())}.jpg',
                                      ContentFile(data), save=False)
            m.save()
            log_action(request.user, 'food_order', target=o,
                       new_value={'action': 'add_menu', 'menu': m.menu_name}, request=request)
            messages.success(request, f'เพิ่มเมนู "{m.menu_name}" แล้ว')
            return redirect(request.path)

    selections_qs = o.selections.select_related('user', 'food_menu').order_by('camp_code_snapshot')
    if center_filter:
        selections_qs = selections_qs.filter(center_snapshot=center_filter)

    menu_stats = []
    grand_total = 0
    for idx, m in enumerate(o.menus.all(), 1):
        m_sels = [s for s in selections_qs if s.food_menu_id == m.pk]
        tot = sum(s.quantity for s in m_sels)
        grand_total += tot
        m.total = tot
        m.member_selections = m_sels
        menu_stats.append(m)

    return render(request, 'food/order_detail.html', {
        'o': o, 'menu_form': menu_form, 'summary': menu_stats,
        'selections': selections_qs, 'grand_total': grand_total, 'center_filter': center_filter,
    })


def _build_order_report_data(o, center_filter=''):
    _decorate_order(o)
    _attach_order_selection_counts([o])
    menus = o.menus.all().order_by('sort_order', 'id')
    menu_data = []
    grand_total = 0
    for idx, menu in enumerate(menus, 1):
        qs = FoodSelection.objects.filter(food_menu=menu)
        if center_filter in ('DS', 'KM'):
            qs = qs.filter(center_snapshot=center_filter)
        members = list(qs.order_by('camp_code_snapshot'))
        total = sum(m.quantity for m in members)
        grand_total += total
        name_list = [
            f'{m.camp_code_snapshot} {m.full_name_snapshot}' + (f'(x{m.quantity})' if m.quantity > 1 else '')
            for m in members
        ]
        lines = [', '.join(name_list[i:i + 5]) for i in range(0, len(name_list), 5)]
        menu_data.append({
            'number': idx,
            'menu': menu,
            'total': total,
            'members': members,
            'name_lines': lines,
        })
    return menu_data, grand_total


@staff_perm_required('staff_can_manage_food')
def order_report(request, pk):
    """หน้ารายงานสรุปยอดการสั่งอาหาร (จัดแสดง 5 รายชื่อต่อบรรทัด พร้อมปุ่มพิมพ์ A4)"""
    o = get_object_or_404(FoodOrder, pk=pk)
    center = request.GET.get('center', '').upper()
    if center not in ('DS', 'KM'):
        center = ''
    menu_data, grand_total = _build_order_report_data(o, center)
    return render(request, 'food/order_report.html', {
        'o': o, 'menu_data': menu_data, 'grand_total': grand_total, 'center': center,
    })


@staff_perm_required('staff_can_manage_food')
def order_print(request, pk):
    """หน้าพิมพ์รายงาน A4 (1 แผ่น) สรุปรายการอาหารและรายชื่อผู้เลือก"""
    o = get_object_or_404(FoodOrder, pk=pk)
    center = request.GET.get('center', '').upper()
    if center not in ('DS', 'KM'):
        center = ''
    menu_data, grand_total = _build_order_report_data(o, center)
    log_action(request.user, 'export', target=o,
               new_value={'print_a4': o.restaurant_name, 'center': center or 'ALL'}, request=request)
    return render(request, 'food/order_print.html', {
        'o': o, 'menu_data': menu_data, 'grand_total': grand_total, 'center': center,
    })


@staff_perm_required('staff_can_manage_food')
@require_POST
def order_action(request, pk):
    o = get_object_or_404(FoodOrder, pk=pk)
    action = request.POST.get('action')
    if action == 'toggle':
        o.edit_toggle_open = not o.edit_toggle_open
        o.save(update_fields=['edit_toggle_open', 'updated_at'])
        log_action(request.user, 'food_order', target=o, new_value={'action': 'toggle',
                                                                    'open': o.edit_toggle_open}, request=request)
        messages.success(request, 'เปิดให้เลือกอาหาร' if o.edit_toggle_open else 'ปิดการเลือกอาหาร')
        nxt = (request.POST.get('next') or '').strip()
        if nxt:
            from django.utils.http import url_has_allowed_host_and_scheme
            if url_has_allowed_host_and_scheme(nxt, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
                return redirect(nxt)
    elif action == 'edit_menu':
        m = get_object_or_404(FoodMenu, pk=request.POST.get('menu_id'), food_order=o)
        m.menu_name = (request.POST.get('menu_name') or m.menu_name).strip()
        m.menu_description = (request.POST.get('menu_description') or '').strip()
        try:
            m.sort_order = int(request.POST.get('sort_order', m.sort_order))
        except ValueError:
            pass
        if 'menu_image' in request.FILES:
            m.menu_image = request.FILES['menu_image']
        elif request.POST.get('edit_ai_image_url'):
            data = _download_image_from_url(request.POST.get('edit_ai_image_url').strip())
            if data:
                m.menu_image.save(f'ai_menu_{o.pk}_{int(timezone.now().timestamp())}.jpg',
                                  ContentFile(data), save=False)
        elif request.POST.get('clear_menu_image'):
            m.menu_image = None
        m.save()
        log_action(request.user, 'food_order', target=o,
                   new_value={'action': 'edit_menu', 'menu': m.menu_name}, request=request)
        messages.success(request, f'แก้ไขเมนู "{m.menu_name}" เรียบร้อย')
    elif action == 'delete_menu':
        m = get_object_or_404(FoodMenu, pk=request.POST.get('menu_id'), food_order=o)
        m.delete()
        log_action(request.user, 'food_order', target=o, new_value={'action': 'delete_menu', 'menu': m.menu_name},
                   request=request)
        messages.success(request, 'ลบเมนูแล้ว')
    elif action == 'delete':
        log_action(request.user, 'food_order', target=o, old_value={'action': 'delete', 'name': str(o)},
                   request=request)
        o.delete()
        messages.success(request, 'ลบรายการอาหารแล้ว')
        return redirect('food:orders')
    return redirect('food:order_detail', pk=o.pk)


@staff_perm_required('staff_can_manage_food')
def order_export(request, pk):
    o = get_object_or_404(FoodOrder, pk=pk)
    rows = [[s.camp_code_snapshot, s.full_name_snapshot, s.food_menu.menu_name, s.quantity,
             s.selected_at] for s in o.selections.select_related('food_menu').order_by('camp_code_snapshot')]
    summary = [[m.menu_name, m.total_selected()] for m in o.menus.all()]
    wb = build_workbook([
        ('สรุปเมนู', ['เมนู', 'จำนวนรวม'], summary),
        ('รายชื่อ', ['รหัสค่าย 1', 'ชื่อ-นามสกุล', 'เมนู', 'จำนวน', 'เวลาเลือก'], rows)])
    log_action(request.user, 'export', target=o, new_value='รายงานอาหาร', request=request)
    return workbook_response(wb, f'อาหาร_{o.food_date}_{o.restaurant_name}.xlsx')


@staff_perm_required('staff_can_manage_food')
def api_generate_food_images(request):
    """API: คืนค่ารายการรูปภาพตัวอย่าง/AI ให้เลือกสำหรับเมนูอาหาร"""
    menu_name = request.GET.get('name', '').strip()
    menu_desc = request.GET.get('desc', '').strip()
    if not menu_name and not menu_desc:
        return JsonResponse({'images': []})
    combined_text = f'{menu_name} {menu_desc}'.strip()
    food_keywords = [
        ('กะเพรา', 'thai basil stir fry pad kra pao', '1562967914-608f82629710', '1540420773420-3366772f4999'),
        ('กระเพรา', 'thai basil stir fry pad kra pao', '1562967914-608f82629710', '1540420773420-3366772f4999'),
        ('ข้าวผัด', 'thai fried rice', '1603133872878-684f208fb84b', '1512058564366-18510be2db19'),
        ('ต้มยำ', 'tom yum spicy soup', '1548943487-a2e4e43b4853', '1569718212165-3a8278d5f624'),
        ('ส้มตำ', 'som tum thai papaya salad', '1559847844-5315695dadae', '1540420773420-3366772f4999'),
        ('ไก่ทอด', 'crispy golden fried chicken', '1626082927389-6cd097cdc6ec', '1562967914-608f82629710'),
        ('ไก่ย่าง', 'thai grilled chicken', '1598515214211-89d3c73ae83b', '1626082927389-6cd097cdc6ec'),
        ('หมูทอด', 'crispy deep fried pork', '1604908176997-125f25cc6f3d', '1544025162-d76694265947'),
        ('หมูกรอบ', 'crispy pork belly', '1544025162-d76694265947', '1604908176997-125f25cc6f3d'),
        ('ข้าวมันไก่', 'hainanese chicken rice with soup', '1512058564366-18510be2db19', '1626082927389-6cd097cdc6ec'),
        ('ข้าวขาหมู', 'thai braised pork leg with rice', '1544025162-d76694265947', '1604908176997-125f25cc6f3d'),
        ('ผัดไทย', 'authentic pad thai noodles', '1559314809-0d155014e29e', '1569718212165-3a8278d5f624'),
        ('แกงเขียวหวาน', 'thai green curry with chicken', '1455619452474-d2be8b1e70cd', '1540420773420-3366772f4999'),
        ('ก๋วยเตี๋ยว', 'thai noodle soup bowl', '1569718212165-3a8278d5f624', '1552611052-33e04de081de'),
        ('บะหมี่', 'egg noodle soup bowl', '1569718212165-3a8278d5f624', '1552611052-33e04de081de'),
        ('ผัดซีอิ๊ว', 'stir fried flat rice noodles with egg', '1585032226651-759b368d7246', '1559314809-0d155014e29e'),
        ('ไข่เจียว', 'fluffy thai omelette over rice', '1525351484163-7529414344d8', '1562967914-608f82629710'),
    ]
    matched_kw = f'delicious thai food dish {combined_text}'
    photo_id_1 = '1546069901-ba9599a7e63c'
    photo_id_2 = '1565299624946-b28f40a0ae38'
    for th_word, en_kw, p1, p2 in food_keywords:
        if th_word in combined_text:
            matched_kw, photo_id_1, photo_id_2 = en_kw, p1, p2
            break
    p1_enc = urllib.parse.quote(f'gourmet {matched_kw}, thai food: {combined_text}, close up culinary photography')
    p2_enc = urllib.parse.quote(f'delicious fresh {matched_kw}, appetizing presentation, photorealistic')
    images = [
        {'id': 1, 'title': 'แบบที่ 1 (ภาพถ่ายจัดจาน)',
         'url': f'https://images.unsplash.com/photo-{photo_id_1}?auto=format&fit=crop&w=512&h=512&q=80'},
        {'id': 2, 'title': 'แบบที่ 2 (ภาพถ่ายมุมมองทางเลือก)',
         'url': f'https://images.unsplash.com/photo-{photo_id_2}?auto=format&fit=crop&w=512&h=512&q=80'},
        {'id': 3, 'title': 'แบบที่ 3 (AI สร้างตามชื่อ 1)',
         'url': f'https://image.pollinations.ai/prompt/{p1_enc}?width=512&height=512&nologo=true&seed=42'},
        {'id': 4, 'title': 'แบบที่ 4 (AI สร้างตามชื่อ 2)',
         'url': f'https://image.pollinations.ai/prompt/{p2_enc}?width=512&height=512&nologo=true&seed=99'},
    ]
    return JsonResponse({'images': images})


# --------------------------------- Member ------------------------------------
@member_required
def my_food(request):
    setting = SystemSetting.get()
    year = setting.current_academic_year
    today = timezone.localdate()
    flt = request.GET.get('status', 'all')
    from composn.models import UseRecord
    is_rights_user = UseRecord.objects.filter(
        user=request.user, academic_year=year, is_used=True
    ).exists()
    all_visible = [
        _decorate_order(o, setting)
        for o in FoodOrder.objects.prefetch_related('menus').filter(
            academic_year=year,
            food_date__gte=today - dt.timedelta(days=30)
        ).order_by('food_date', 'id')
        if o.is_visible_to(request.user, year=year, is_rights_user=is_rights_user)
    ]
    if flt == 'active':
        items = [o for o in all_visible if o.get_status() == 'active']
    else:
        items = all_visible
    _attach_order_selection_counts(items)
    mine = {}
    for s in FoodSelection.objects.filter(user=request.user, food_order__in=items).select_related('food_menu'):
        mine.setdefault(s.food_order_id, []).append(s)
    for o in items:
        o.my_selections = mine.get(o.pk, [])
    open_surveys_all = [
        s for s in CampSurvey.objects.filter(academic_year=year, is_open=True).prefetch_related('questions')
        if s.is_visible_to(request.user, year=year, is_rights_user=is_rights_user)
    ]
    answered_survey_ids = set(
        CampSurveyAnswer.objects.filter(user=request.user, survey__in=open_surveys_all)
        .values_list('survey_id', flat=True)
    )
    open_surveys = [s for s in open_surveys_all if s.pk not in answered_survey_ids]
    my_selections_list = list(
        FoodSelection.objects.filter(
            user=request.user,
            food_order__academic_year=year,
        ).select_related('food_order', 'food_menu')
        .order_by('food_order__food_date', 'food_menu__sort_order')
    )
    return render(request, 'food/my_food.html', {
        'orders': items, 'flt': flt, 'year': year,
        'open_surveys': open_surveys,
        'selections': my_selections_list,
    })


@member_required
def my_selections(request):
    """หน้ารายการอาหารทั้งหมดที่สมาชิกเคยเลือกไว้ในปีการศึกษาปัจจุบัน"""
    year = SystemSetting.get().current_academic_year
    sels = FoodSelection.objects.filter(
        user=request.user,
        food_order__academic_year=year,
    ).select_related('food_order', 'food_menu').order_by('food_order__food_date', 'food_menu__sort_order')
    return render(request, 'food/my_selections.html', {'selections': sels, 'year': year})


@member_required
def my_food_detail(request, pk):
    o = _decorate_order(get_object_or_404(FoodOrder, pk=pk))
    _attach_order_selection_counts([o])
    if not o.is_visible_to(request.user):
        messages.error(request, 'รายการนี้ไม่เปิดให้กลุ่มหรือศูนย์ของท่าน')
        return redirect('food:my_food')
    editable = o.is_editable_by_member()
    existing_qs = list(FoodSelection.objects.filter(food_order=o, user=request.user).select_related('food_menu'))
    current = {s.food_menu_id: s.quantity for s in existing_qs}
    existing_map = {s.food_menu_id: s for s in existing_qs}
    total_selected = sum(current.values())

    if request.method == 'POST':
        if not editable:
            messages.error(request, 'รายการนี้ปิดการแก้ไขแล้ว')
            return redirect(request.path)

        action = request.POST.get('action', '')
        # รองรับปุ่มกดเลือก/ยกเลิก/เพิ่ม-ลดจำนวนทีละเมนู (รูปแบบระบบ food เดิม)
        if action in ('select', 'remove', 'update_quantity'):
            if action == 'select':
                menu = get_object_or_404(FoodMenu, pk=request.POST.get('menu_id'), food_order=o)
                if total_selected >= o.max_selection_count:
                    messages.error(request, f'เลือกได้สูงสุด {o.max_selection_count} อย่าง')
                else:
                    sel, created = FoodSelection.objects.get_or_create(
                        food_order=o, food_menu=menu, user=request.user, defaults={'quantity': 1}
                    )
                    if not created and o.selection_mode == FoodOrder.MODE_REPEAT:
                        sel.quantity += 1
                        sel.save(update_fields=['quantity', 'updated_at'])
                    log_action(request.user, 'food_select', target=o,
                               new_value={'select': menu.menu_name}, request=request)
                    messages.success(request, f'เลือกเมนู "{menu.menu_name}" เรียบร้อย')
            elif action == 'remove':
                sel = FoodSelection.objects.filter(pk=request.POST.get('selection_id'),
                                                   food_order=o, user=request.user).first()
                if sel:
                    mname = sel.food_menu.menu_name
                    sel.delete()
                    log_action(request.user, 'food_select', target=o,
                               new_value={'remove': mname}, request=request)
                    messages.success(request, f'ยกเลิกเมนู "{mname}" แล้ว')
            elif action == 'update_quantity':
                sel = FoodSelection.objects.filter(pk=request.POST.get('selection_id'),
                                                   food_order=o, user=request.user).first()
                if sel:
                    try:
                        new_q = int(request.POST.get('quantity') or 0)
                    except ValueError:
                        new_q = sel.quantity
                    other_total = total_selected - sel.quantity
                    if new_q <= 0:
                        sel.delete()
                        messages.success(request, 'นำรายการออกแล้ว')
                    elif other_total + new_q > o.max_selection_count:
                        messages.error(request, f'เลือกได้สูงสุด {o.max_selection_count} อย่าง')
                    else:
                        sel.quantity = new_q
                        sel.save(update_fields=['quantity', 'updated_at'])
                        messages.success(request, 'อัปเดตจำนวนเรียบร้อย')
            return redirect(request.path)

        # รองรับการส่งแบบฟอร์มรวม qty_<id> (คงความเข้ากันได้กับชุดทดสอบอัตโนมัติ)
        chosen = {}
        for m in o.menus.all():
            try:
                q = int(request.POST.get(f'qty_{m.pk}') or 0)
            except ValueError:
                q = 0
            if q > 0:
                chosen[m.pk] = q if o.selection_mode == FoodOrder.MODE_REPEAT else 1
        total = sum(chosen.values())
        if total > o.max_selection_count:
            messages.error(request, f'เลือกได้สูงสุด {o.max_selection_count} รายการ (ท่านเลือก {total})')
            return redirect(request.path)
        with transaction.atomic():
            FoodSelection.objects.filter(food_order=o, user=request.user).exclude(food_menu_id__in=chosen).delete()
            for mid, q in chosen.items():
                FoodSelection.objects.update_or_create(food_order=o, food_menu_id=mid, user=request.user,
                                                       defaults={'quantity': q})
        log_action(request.user, 'food_select', target=o, new_value=chosen, request=request)
        messages.success(request, 'บันทึกการเลือกอาหารเรียบร้อย')
        return redirect(request.path)

    menus = list(o.menus.annotate(group_count=Sum('selections__quantity')))
    return render(request, 'food/my_food_detail.html', {
        'o': o, 'editable': editable, 'current': current,
        'existing_selections': existing_qs, 'existing_map': existing_map,
        'total_selected': total_selected, 'menus': menus,
    })


# =============================================================================
# แบบสอบถามสำหรับนักเรียนในค่าย & สรุปภาพรวมการสั่งอาหาร (Admin / Member)
# =============================================================================
@staff_perm_required('staff_can_manage_food')
def surveys_manage(request):
    """Admin/Staff สร้างและจัดการแบบสอบถามสำหรับนักเรียนในค่าย พร้อมดู Dashboard สรุปภาพรวมเกี่ยวกับการสั่งอาหาร"""
    from django.db.models import Q
    setting = SystemSetting.get()
    year_req = request.GET.get('year') or request.POST.get('year')
    try:
        year = int(year_req) if year_req else int(setting.current_academic_year)
    except (ValueError, TypeError):
        year = int(setting.current_academic_year)

    if request.method == 'POST':
        action = request.POST.get('action', '')
        if action == 'create_survey':
            title = (request.POST.get('title') or '').strip()
            desc = (request.POST.get('description') or '').strip()
            if request.POST.get('target_groups_submitted') == '1':
                allow_current = bool(request.POST.get('allow_current_students'))
                allow_rights = bool(request.POST.get('allow_rights_students'))
            else:
                allow_current = True
                allow_rights = False
            if not title:
                messages.error(request, 'กรุณาระบุหัวข้อแบบสอบถาม')
            elif not allow_current and not allow_rights:
                messages.error(request, 'กรุณาเลือกกลุ่มผู้มีสิทธิ์ตอบแบบสอบถามอย่างน้อย 1 กลุ่ม (นักเรียนปัจจุบัน หรือ นักเรียนใช้สิทธิ์)')
            else:
                s = CampSurvey.objects.create(
                    academic_year=year, title=title, description=desc,
                    allow_current_students=allow_current,
                    allow_rights_students=allow_rights,
                    is_open=bool(request.POST.get('is_open', True)), created_by=request.user,
                )
                q_text = (request.POST.get('first_question') or '').strip()
                if q_text:
                    CampSurveyQuestion.objects.create(
                        survey=s,
                        question_text=q_text,
                        question_type=request.POST.get('first_question_type') or CampSurveyQuestion.TYPE_CHOICE,
                        choices_text=(request.POST.get('first_choices') or '').strip(),
                        sort_order=1,
                    )
                log_action(request.user, 'food_order', target=s,
                           new_value={'create_survey': s.title, 'year': year, 'groups': s.target_groups_label()},
                           request=request)
                messages.success(request, f'สร้างแบบสอบถาม "{s.title}" เรียบร้อย')
                return redirect(f'{request.path}?year={year}')

        elif action == 'update_survey_groups':
            s = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'))
            allow_current = bool(request.POST.get('allow_current_students'))
            allow_rights = bool(request.POST.get('allow_rights_students'))
            if not allow_current and not allow_rights:
                messages.error(request, 'กรุณาเลือกกลุ่มผู้มีสิทธิ์ตอบแบบสอบถามอย่างน้อย 1 กลุ่ม')
            else:
                s.allow_current_students = allow_current
                s.allow_rights_students = allow_rights
                s.save(update_fields=['allow_current_students', 'allow_rights_students', 'updated_at'])
                log_action(request.user, 'food_order', target=s,
                           new_value={'update_survey_groups': s.target_groups_label()}, request=request)
                messages.success(request, f'บันทึกกลุ่มผู้มีสิทธิ์ตอบ "{s.title}" เรียบร้อยแล้ว ({s.target_groups_label()})')
            return redirect(f'{request.path}?year={year}')

        elif action == 'create_default_food_survey':
            with transaction.atomic():
                s = CampSurvey.objects.create(
                    academic_year=year,
                    title=f'แบบสอบถามสำรวจการสั่งอาหารและความต้องการอาหารในค่าย (ปีการศึกษา {year})',
                    description='สำหรับสำรวจประเภทอาหาร เมนูที่ต้องการ และข้อจำกัดด้านอาหารของนักเรียนในค่าย',
                    allow_current_students=True,
                    allow_rights_students=False,
                    is_open=True,
                    created_by=request.user,
                )
                CampSurveyQuestion.objects.create(
                    survey=s, sort_order=1,
                    question_text='ประเภทอาหารหลักที่ท่านรับประทานในค่าย',
                    question_type=CampSurveyQuestion.TYPE_CHOICE,
                    choices_text='อาหารทั่วไป, อาหารฮาลาล (อิสลาม), อาหารมังสวิรัติ/เจ, ไม่ทานเนื้อวัว, ไม่ทานหมู',
                )
                CampSurveyQuestion.objects.create(
                    survey=s, sort_order=2,
                    question_text='ประเภทเมนูอาหารที่ท่านชื่นชอบ (เลือกได้หลายข้อ)',
                    question_type=CampSurveyQuestion.TYPE_MULTIPLE,
                    choices_text='กะเพรา/ผัดพริกแกง, ข้าวผัด, ข้าวมันไก่/ข้าวหมูแดง, ก๋วยเตี๋ยว/เส้น, แกงจืด/ต้มยำ, ไก่ทอด/หมูทอด',
                )
                CampSurveyQuestion.objects.create(
                    survey=s, sort_order=3,
                    question_text='ความพึงพอใจต่อระบบการสั่งอาหารและปริมาณอาหารในค่าย',
                    question_type=CampSurveyQuestion.TYPE_RATING,
                )
                CampSurveyQuestion.objects.create(
                    survey=s, sort_order=4, is_required=False,
                    question_text='ข้อเสนอแนะเพิ่มเติมเกี่ยวกับการสั่งอาหาร / เมนูที่อยากให้จัดเพิ่ม',
                    question_type=CampSurveyQuestion.TYPE_TEXT,
                )
            log_action(request.user, 'food_order', target=s,
                       new_value={'create_default_food_survey': s.title}, request=request)
            messages.success(request, 'สร้างแบบสอบถามสำรวจการสั่งอาหารมาตรฐานเรียบร้อยแล้ว')
            return redirect(f'{request.path}?year={year}')

        elif action == 'add_question':
            s = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'))
            q_text = (request.POST.get('question_text') or '').strip()
            q_type = request.POST.get('question_type') or CampSurveyQuestion.TYPE_CHOICE
            choices_text = (request.POST.get('choices_text') or '').strip()
            if not q_text:
                messages.error(request, 'กรุณาระบุข้อความคำถาม')
            else:
                CampSurveyQuestion.objects.create(
                    survey=s,
                    question_text=q_text,
                    question_type=q_type,
                    choices_text=choices_text,
                    is_required=bool(request.POST.get('is_required')),
                    sort_order=s.questions.count() + 1,
                )
                messages.success(request, 'เพิ่มคำถามในแบบสอบถามเรียบร้อย')
            return redirect(f'{request.path}?year={year}')

        elif action == 'delete_question':
            q = get_object_or_404(CampSurveyQuestion, pk=request.POST.get('question_id'))
            q.delete()
            messages.success(request, 'ลบคำถามเรียบร้อย')
            return redirect(f'{request.path}?year={year}')

        elif action == 'toggle_survey':
            s = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'))
            s.is_open = not s.is_open
            s.save(update_fields=['is_open', 'updated_at'])
            messages.success(request, f'{"เปิด" if s.is_open else "ปิด"}แบบสอบถาม "{s.title}" แล้ว')
            return redirect(f'{request.path}?year={year}')

        elif action == 'delete_survey':
            s = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'))
            s.delete()
            messages.success(request, 'ลบแบบสอบถามเรียบร้อย')
            return redirect(f'{request.path}?year={year}')

    surveys = list(CampSurvey.objects.filter(academic_year=year).prefetch_related('questions', 'answers__user'))
    survey_summaries = [build_survey_summary(s) for s in surveys]

    # ข้อมูลภาพรวมการสั่งอาหารและข้อจำกัดด้านอาหารของนักเรียนในปีการศึกษานี้
    yy = f'{year % 100:02d}'
    members_qs = User.objects.filter(role=User.ROLE_MEMBER, is_active=True).filter(
        Q(camp1_code__istartswith=yy) | Q(academic_year=year)
    ).order_by('camp1_code')
    muslim_students = list(members_qs.filter(is_muslim=True))
    allergy_students = list(members_qs.exclude(food_allergy=''))
    disease_students = list(members_qs.exclude(chronic_disease=''))

    orders_qs = FoodOrder.objects.filter(academic_year=year).prefetch_related('menus', 'selections')
    order_stats = []
    for o in orders_qs[:15]:
        total_sel = sum(sel.quantity for sel in o.selections.all())
        user_cnt = len({sel.user_id for sel in o.selections.all() if sel.user_id})
        order_stats.append({
            'order': o,
            'total_quantity': total_sel,
            'user_count': user_cnt,
            'menu_count': o.menus.count(),
        })

    return render(request, 'food/surveys_manage.html', {
        'year': year,
        'survey_summaries': survey_summaries,
        'type_choices': CampSurveyQuestion.TYPE_CHOICES,
        'member_total': members_qs.count(),
        'muslim_students': muslim_students,
        'allergy_students': allergy_students,
        'disease_students': disease_students,
        'order_stats': order_stats,
    })


@member_required
def member_surveys(request):
    """หน้านักเรียนทำแบบสอบถามในค่าย (เกี่ยวกับการสั่งอาหาร/ความต้องการในค่าย)"""
    from composn.models import UseRecord
    year = int(SystemSetting.get().current_academic_year)
    is_rights_user = UseRecord.objects.filter(
        user=request.user, academic_year=year, is_used=True
    ).exists()
    surveys = [
        s for s in CampSurvey.objects.filter(academic_year=year, is_open=True).prefetch_related('questions')
        if s.is_visible_to(request.user, year=year, is_rights_user=is_rights_user)
    ]

    if request.method == 'POST':
        survey = get_object_or_404(CampSurvey, pk=request.POST.get('survey_id'), academic_year=year, is_open=True)
        if not survey.is_visible_to(request.user, year=year, is_rights_user=is_rights_user):
            messages.error(request, 'แบบสอบถามนี้ไม่เปิดให้กลุ่มของท่าน')
            return redirect('food:member_surveys')
        with transaction.atomic():
            for q in survey.questions.all():
                if q.question_type == CampSurveyQuestion.TYPE_MULTIPLE:
                    vals = [v.strip() for v in request.POST.getlist(f'q_{q.pk}') if v.strip()]
                    ans_text = ', '.join(vals)
                else:
                    ans_text = (request.POST.get(f'q_{q.pk}') or '').strip()
                CampSurveyAnswer.objects.update_or_create(
                    question=q,
                    user=request.user,
                    defaults={'survey': survey, 'answer_text': ans_text},
                )
        log_action(request.user, 'food_select', target=survey,
                   new_value={'survey_submit': survey.title}, request=request)
        messages.success(request, f'บันทึกคำตอบแบบสอบถาม "{survey.title}" เรียบร้อยแล้ว')
        return redirect('food:member_surveys')

    my_answers = {
        a.question_id: a.answer_text
        for a in CampSurveyAnswer.objects.filter(user=request.user, survey__in=surveys)
    }
    for s in surveys:
        s.q_items = []
        for q in s.questions.all():
            cur = my_answers.get(q.pk, '')
            cur_list = [p.strip() for p in cur.split(',') if p.strip()]
            s.q_items.append({
                'q': q,
                'choices': q.get_choices_list(),
                'current': cur,
                'current_list': cur_list,
            })

    return render(request, 'food/member_surveys.html', {
        'surveys': surveys,
        'year': year,
    })
