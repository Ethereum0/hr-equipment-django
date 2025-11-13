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
from .ad_service import ADService
from .models import Employee, Equipment
from .forms import EmployeeForm, EquipmentForm

from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse
import json
from django.core.cache import cache

from django.http import StreamingHttpResponse
import json

import uuid


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

def get_ad_status(request, employee_id):
    """AJAX: Получить статус AD"""
    employee = get_object_or_404(Employee, pk=employee_id)
    if not employee.ad_login:
        return JsonResponse({'error': 'No AD login'})
    
    ad = ADService()
    status = ad.get_user_status(employee.ad_login)
    return JsonResponse(status)

def unlock_ad_account(request, employee_id):
    """AJAX: Разблокировать учетку AD"""
    employee = get_object_or_404(Employee, pk=employee_id)
    if not employee.ad_login:
        return JsonResponse({'error': 'No AD login'})
    
    ad = ADService()
    result = ad.unlock_user(employee.ad_login)
    return JsonResponse(result)

# staff/views.py - добавить эту функцию

def ad_events_stream(request, employee_id):
    """Server-Sent Events поток для обновлений AD статуса"""
    employee = get_object_or_404(Employee, pk=employee_id)
    
    def event_stream():
        # Отправляем начальный статус
        yield f"data: {json.dumps({'type': 'connected', 'username': employee.ad_login})}\n\n"
        
        # Здесь можно добавить логику долгоживущего соединения
        # Но для простоты просто закрываем соединение
        yield f"data: {json.dumps({'type': 'complete'})}\n\n"
    
    response = StreamingHttpResponse(
        event_stream(),
        content_type='text/event-stream'
    )
    response['Cache-Control'] = 'no-cache'
    response['Connection'] = 'keep-alive'
    return response

@csrf_exempt
def ad_lockout_webhook(request):
    """Webhook для получения событий блокировки из AD"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)
    
    try:
        data = json.loads(request.body)
        username = data.get('username')
        event_type = data.get('event_type')
        
        print(f"🔔 Webhook received: {username} - {event_type}")
        
        if not username:
            return JsonResponse({'error': 'Username required'}, status=400)
        
        # Сбрасываем кэш для этого пользователя
        cache_key = f'ad_status_{username}'
        cache.delete(cache_key)
        
        # Сохраняем уведомление
        notification_id = str(uuid.uuid4())
        notification_key = f'ad_notification_{notification_id}'
        
        # Ищем сотрудника по AD логину
        try:
            employee = Employee.objects.get(ad_login=username)
            employee_name = employee.fio
            employee_id = employee.id
        except Employee.DoesNotExist:
            employee_name = username
            employee_id = None
        
        notification_data = {
            'id': notification_id,
            'username': username,
            'employee_name': employee_name,
            'employee_id': employee_id,
            'event_type': event_type,
            'timestamp': timezone.now().isoformat(),
            'source_computer': data.get('source_computer', 'Unknown')
        }
        
        # Сохраняем уведомление на 1 час
        cache.set(notification_key, notification_data, 3600)
        
        # Добавляем ID в список активных уведомлений
        notifications_list = cache.get('ad_notifications_list', [])
        notifications_list.append(notification_id)
        cache.set('ad_notifications_list', notifications_list, 3600)
        
        print(f"✅ Notification created for: {username}")
        
        return JsonResponse({
            'success': True, 
            'message': f'Notification created for {username}',
            'username': username,
            'employee_id': employee_id
        })
        
    except Exception as e:
        print(f"❌ Webhook error: {e}")
        return JsonResponse({'error': str(e)}, status=500)
    

def check_ad_event(request, employee_id):
    """Проверяет есть ли события для сотрудника"""
    employee = get_object_or_404(Employee, pk=employee_id)
    
    if not employee.ad_login:
        return JsonResponse({'has_event': False})
    
    event_key = f'ad_event_{employee.ad_login}'
    event_data = cache.get(event_key)
    
    if event_data:
        # Удаляем событие после отправки
        cache.delete(event_key)
        return JsonResponse({
            'has_event': True,
            'event_data': event_data
        })
    
    return JsonResponse({'has_event': False})

def check_notifications(request):
    """Проверяет новые уведомления о блокировках"""
    notifications_list = cache.get('ad_notifications_list', [])
    new_notifications = []
    
    for notification_id in notifications_list:
        notification_key = f'ad_notification_{notification_id}'
        notification_data = cache.get(notification_key)
        
        if notification_data:
            new_notifications.append(notification_data)
            # НЕ удаляем уведомление - оставляем пока не истечет TTL (1 час)
    
    return JsonResponse({
        'notifications': new_notifications,
        'count': len(new_notifications)
    })

@csrf_exempt
def remove_notification(request):
    """Удаляет уведомление при закрытии"""
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            notification_id = data.get('notification_id')
            
            if notification_id:
                # Удаляем уведомление из кэша
                notification_key = f'ad_notification_{notification_id}'
                cache.delete(notification_key)
                
                # Удаляем ID из списка активных уведомлений
                notifications_list = cache.get('ad_notifications_list', [])
                if notification_id in notifications_list:
                    notifications_list.remove(notification_id)
                    cache.set('ad_notifications_list', notifications_list, 3600)
                
                print(f"✅ Notification removed: {notification_id}")
                return JsonResponse({'success': True})
            
        except Exception as e:
            print(f"❌ Remove notification error: {e}")
    
    return JsonResponse({'success': False})


