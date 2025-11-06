# staff/views.py
import re
import time
import importlib
from collections import defaultdict
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.http import HttpResponse, HttpResponseRedirect, JsonResponse
from django.db.models import Count, Q
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_GET
from django.templatetags.static import static
from django.conf import settings
from django.utils import timezone
from openpyxl import load_workbook
import requests

from .models import Employee, Equipment
from .forms import EmployeeForm, EquipmentForm

# Оставляем только основные функции управления сотрудниками и оборудованием
# УДАЛИТЬ все функции: generate_config, provisioning_dashboard, provisioning_create_ad_account,
# get_sip_phone_info, provisioning_get_sip_phone_ip, get_sip_screenshot, 
# provisioning_send_welcome_email, send_sip_command


try:
    import requests
except ModuleNotFoundError:
    def missing_requests_warning(request):
        return JsonResponse({
            'success': False,
            'message': '❌ Ошибка: библиотека "requests" не установлена. '
                       'Установите её командой: pip install requests'
        })
    requests = None





@login_required
def dashboard(request):
    active = Employee.objects.filter(status='active').count()
    dismissed = Employee.objects.filter(status='dismissed').count()
    maternity = Employee.objects.filter(status='maternity').count()
    total_employees = Employee.objects.count()

    free_equipment_summary_raw = Equipment.objects.filter(
        employee__isnull=True,
        disposed=False
    ).values('type', 'office').annotate(count=Count('id')).order_by('type', 'office')

    free_summary = {
        'types': set(),
        'offices': set(),
        'data': defaultdict(lambda: defaultdict(int))
    }

    for item in free_equipment_summary_raw:
        type_key = item['type']
        office_key = item['office']
        count = item['count']
        free_summary['types'].add(type_key)
        free_summary['offices'].add(office_key)
        free_summary['data'][type_key][office_key] = count

    free_summary['types'] = sorted(list(free_summary['types']), key=lambda x: dict(Equipment.EQUIPMENT_TYPES).get(x, x))
    free_summary['offices'] = sorted(list(free_summary['offices']), key=lambda x: dict(Employee.OFFICE_CHOICES).get(x, x))
    free_summary['data'] = {k: dict(v) for k, v in free_summary['data'].items()}

    return render(request, 'staff/dashboard.html', {
        'active': active,
        'dismissed': dismissed,
        'maternity': maternity,
        'total_employees': total_employees,
        'free_equipment_summary': free_summary,
        'equipment_type_labels': dict(Equipment.EQUIPMENT_TYPES),
        'office_labels': dict(Employee.OFFICE_CHOICES),
    })



@login_required
def employees_list(request):
    selected_id = request.GET.get('selected_id')
    employees = Employee.objects.all().order_by('fio')

    paginator = Paginator(employees, 1000)  # отдаём много — фронт сам фильтрует
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    selected_employee = None
    if selected_id:
        try:
            selected_employee = Employee.objects.prefetch_related('equipment_set').get(id=selected_id)
        except Employee.DoesNotExist:
            selected_employee = None

    # HTMX запрос -> только правая панель
    if request.headers.get('HX-Request') == 'true':
        if selected_employee:
            return render(request, 'staff/partials/right_column_content.html', {
                'selected_employee': selected_employee
            })
        return render(request, 'staff/partials/right_column_empty.html')

    # Обычная загрузка -> отдаём всю страницу и ВСЕХ без фильтрации
    return render(request, 'staff/employees_list.html', {
        'page_obj': page_obj,
        'selected_employee': selected_employee,
        'query': '',              # пусто
        'status_filter': 'all',   # без фильтра
        'office_filter': 'all',   # без фильтра
        'per_page': 1000,
    })


@login_required
def employee_create(request):
    if request.method == 'POST':
        form = EmployeeForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('employees_list')
    else:
        form = EmployeeForm()
    return render(request, 'staff/employee_form.html', {'form': form})


@login_required
def employee_edit(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    if request.method == 'POST':
        form = EmployeeForm(request.POST, instance=employee)
        if form.is_valid():
            form.save()
            if request.headers.get('x-requested-with') == 'XMLHttpRequest':
                return JsonResponse({'status': 'success'})
            return redirect('employees_list')
        else:
            if request.headers.get('x-requested-with') == 'XMLHttpRequest':
                return JsonResponse({'status': 'error', 'errors': form.errors}, status=400)
    else:
        form = EmployeeForm(instance=employee)
    return render(request, 'staff/employee_form.html', {'form': form})


@login_required
def equipment_list(request):
    query = request.GET.get('q', '')
    type_filter = request.GET.get('type', '')
    office_filter = request.GET.get('office', '')
    disposed_filter = request.GET.get('disposed', '')
    assigned_filter = request.GET.get('assigned', '')  
    employee_filter = request.GET.get('employee', '')  
    per_page = request.GET.get('per_page', 20)
    equipment = Equipment.objects.select_related('employee').all()
    
    if query:
        equipment = equipment.filter(
            Q(serial_number__icontains=query) |
            Q(model__icontains=query) |
            Q(mac_address__icontains=query) |
            Q(employee__fio__icontains=query)
        )
    if type_filter:
        equipment = equipment.filter(type=type_filter)
    if office_filter:
        equipment = equipment.filter(office=office_filter)
    if disposed_filter:
        equipment = equipment.filter(disposed=disposed_filter == 'true')
    
    # ✅ Фильтр по закреплению
    if assigned_filter == 'true':  # Закрепленное
        equipment = equipment.exclude(employee__isnull=True)
    elif assigned_filter == 'false':  # Свободное
        equipment = equipment.filter(employee__isnull=True)
    
    if employee_filter:  # ✅ Фильтр по сотруднику
        try:
            employee_id = int(employee_filter)
            equipment = equipment.filter(employee_id=employee_id)
        except (ValueError, TypeError):
            pass    


    try:
        per_page = int(per_page)
    except:
        per_page = 20
    per_page = min(max(per_page, 10), 100)
    equipment = equipment.order_by('-id')
    paginator = Paginator(equipment, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # ✅ Передаем всех сотрудников для выпадающего списка
    all_employees = Employee.objects.filter(status='active').order_by('fio')    
    
    return render(request, 'staff/equipment_list.html', {
        'page_obj': page_obj,
        'query': query,
        'type_filter': type_filter,
        'office_filter': office_filter,
        'disposed_filter': disposed_filter,
        'assigned_filter': assigned_filter,  # ✅ Передаем в шаблон
        'per_page': per_page,
        'equipment_types': Equipment.EQUIPMENT_TYPES,
        'office_choices': Employee.OFFICE_CHOICES,
        'all_employees': all_employees,  # ✅ Всегда передаём
    })

@login_required
def equipment_create(request):
    if request.method == 'POST':
        form = EquipmentForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('equipment_list')
    else:
        form = EquipmentForm()
    return render(request, 'staff/equipment_form.html', {'form': form})

@login_required
def equipment_edit(request, pk):
    equipment = get_object_or_404(Equipment, pk=pk)
    if request.method == 'POST':
        form = EquipmentForm(request.POST, instance=equipment)
        if form.is_valid():
            form.save()
            return redirect('equipment_list')
    else:
        form = EquipmentForm(instance=equipment)
    return render(request, 'staff/equipment_form.html', {'form': form})

@login_required
def equipment_delete(request, pk):
    equipment = get_object_or_404(Equipment, pk=pk)
    if request.method == 'POST':
        equipment.delete()
        return redirect('equipment_list')
    return render(request, 'staff/equipment_confirm_delete.html', {'equipment': equipment})


@login_required
def employee_change_status(request, pk):
    """
    Изменяет статус сотрудника через AJAX.
    """
    employee = get_object_or_404(Employee, pk=pk)
    
    if request.method == 'POST':
        new_status = request.POST.get('status')
        if new_status in dict(Employee.STATUS_CHOICES):
            employee.status = new_status
            employee.save()
            return JsonResponse({
                'success': True,
                'message': f'Статус сотрудника {employee.fio} изменён на {employee.get_status_display()}',
                'status_badge_class': employee.get_status_badge_class(),
                'status_display': employee.get_status_display()
            })
        else:
            return JsonResponse({
                'success': False,
                'message': 'Недопустимый статус'
            }, status=400)
    
    return JsonResponse({
        'success': False,
        'message': 'Метод не разрешён'
    }, status=405)

@login_required
def equipment_assign_employee(request, pk):
    """
    Изменяет сотрудника, закрепленного за оборудованием.
    """
    equipment = get_object_or_404(Equipment, pk=pk)
    
    if request.method == 'POST':
        employee_id = request.POST.get('employee_id')
        next_url = request.POST.get('next', '/equipment/')
        
        if employee_id:
            try:
                employee = Employee.objects.get(id=employee_id)
                equipment.employee = employee
                equipment.save()
            except Employee.DoesNotExist:
                pass
        else:
            equipment.employee = None
            equipment.save()
        
        return HttpResponseRedirect(next_url)
    
    return HttpResponseRedirect('/equipment/')


















