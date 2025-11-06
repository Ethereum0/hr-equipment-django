# staff/views_provisioning.py
import os
import re
import shutil
import time
import logging
import requests
from django.http import JsonResponse, HttpResponse
from django.conf import settings
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth.decorators import login_required
from django.views.decorators.http import require_POST
from django.shortcuts import render, get_object_or_404
from django.utils import timezone

from staff.models import Employee, Equipment

logger = logging.getLogger(__name__)

# === 📊 Основная страница Provisioning ===
@login_required
def provisioning_dashboard(request):
    """
    Страница Provisioning: выбор сотрудника и выполнение действий.
    """
    # Получаем всех активных сотрудников
    active_employees = Employee.objects.filter(status='active').order_by('fio')
    
    # Получаем выбранного сотрудника (если передан ID)
    selected_employee_id = request.GET.get('employee_id')
    selected_employee = None
    sip_phones = []  # Список для SIP-телефонов
    
    if selected_employee_id:
        try:
            # Получаем сотрудника с оборудованием
            selected_employee = Employee.objects.prefetch_related('equipment_set').get(
                id=selected_employee_id, 
                status='active'
            )
            # Фильтруем только SIP-телефоны
            sip_phones = selected_employee.equipment_set.filter(type='SIP')
        except Employee.DoesNotExist:
            pass

    return render(request, 'staff/provisioning_dashboard.html', {
        'active_employees': active_employees,
        'selected_employee': selected_employee,
        'sip_phones': sip_phones,
    })

# === ⚙️ Генерация конфигурации ===
@csrf_exempt
@login_required
def generate_config(request):
    if request.method != "POST":
        return JsonResponse({'success': False, 'message': 'Метод не разрешён'})

    try:
        employee_id = request.POST.get("employee_id")
        equipment_id = request.POST.get("equipment_id")  # Новый параметр
        
        if not employee_id:
            return JsonResponse({'success': False, 'message': 'Не выбран сотрудник.'})

        employee = Employee.objects.prefetch_related('equipment_set').get(id=employee_id)

        # Получаем конкретное оборудование или первое SIP
        if equipment_id:
            sip_phone = employee.equipment_set.filter(id=equipment_id, type='SIP').first()
        else:
            sip_phone = employee.equipment_set.filter(type='SIP').first()
            
        if not sip_phone:
            return JsonResponse({'success': False, 'message': 'У сотрудника нет SIP-телефона.'})

        internal_number = getattr(employee, 'sip_number', None) or "Не заполнен"
        mac_address = getattr(sip_phone, 'mac_address', None) or "Не заполнен"
        mac_clean = mac_address.replace(':', '').replace('-', '').replace('.', '').upper()
        filename = f"{mac_clean}.cfg"

        config_lines = [
            "#!version:1.0.0.1",
            "account.1.enable = 1",
            f"account.1.label = {internal_number} | {employee.last_name} {employee.first_name}",
            f"account.1.display_name = {employee.last_name} {employee.first_name}",
            f"account.1.auth_name = Profelectro{internal_number}",
            f"account.1.user_name = {internal_number}",
            f"account.1.password = {employee.pass_3cx}",
            # f"# Оборудование: {sip_phone.model or 'N/A'}",
            # f"# MAC: {mac_address}",
            "",
        ]
        config_content = "\n".join(config_lines)

        return JsonResponse({
            'success': True,
            'message': '✅ Файл успешно сформирован.',
            'config': config_content,
            'filename': filename,
            'employee_fio': employee.fio,
            'internal_number': internal_number,
            'mac_address': mac_address,
        })

    except Employee.DoesNotExist:
        return JsonResponse({'success': False, 'message': 'Сотрудник не найден.'})
    except Exception as e:
        return JsonResponse({'success': False, 'message': f'Ошибка: {e}'})

# === 📷 Получение скриншота SIP-телефона ===
@login_required
def provisioning_sip_screen(request):
    """
    Получает скриншот SIP-телефона и чистит старые снимки.
    """
    try:
        employee_id = request.GET.get('employee_id')
        equipment_id = request.GET.get('equipment_id')
        
        if not employee_id:
            return JsonResponse({'success': False, 'message': 'Не указан сотрудник.'})

        employee = Employee.objects.prefetch_related('equipment_set').get(id=employee_id)
        
        # Получаем конкретное оборудование или первое SIP
        if equipment_id:
            sip_phone = employee.equipment_set.filter(id=equipment_id, type='SIP').first()
        else:
            sip_phone = employee.equipment_set.filter(type='SIP').first()
            
        if not sip_phone:
            return JsonResponse({'success': False, 'message': 'У сотрудника нет SIP-телефона.'})

        ip_field = sip_phone.ip_or_anydesk or ""
        match = re.search(r'(\d{1,3}(?:\.\d{1,3}){3})', ip_field)
        if not match:
            return JsonResponse({'success': False, 'message': f"Не удалось извлечь IP из '{ip_field}'."})
        ip = match.group(1)

        login = getattr(sip_phone, 'login', 'admin')
        password = getattr(sip_phone, 'password', 'admin')

        url = f"http://{login}:{password}@{ip}/servlet?m=mod_action&command=screenshot"
        response = requests.get(url, timeout=5)
        if response.status_code != 200:
            return JsonResponse({'success': False, 'message': f'Ошибка HTTP {response.status_code} при обращении к {ip}.'})

        screenshots_dir = os.path.join(settings.MEDIA_ROOT, 'sip_screenshots')
        os.makedirs(screenshots_dir, exist_ok=True)

        # 🧹 Удаляем старые файлы (старше 10 секунд)
        now = time.time()
        for f in os.listdir(screenshots_dir):
            path = os.path.join(screenshots_dir, f)
            if os.path.isfile(path) and now - os.path.getmtime(path) > 10:
                os.remove(path)

        # 💾 Сохраняем новый
        filename = f"sip_{employee.id}_{equipment_id or 'default'}_{timezone.now().strftime('%Y%m%d_%H%M%S')}.png"
        filepath = os.path.join(screenshots_dir, filename)
        with open(filepath, 'wb') as f:
            f.write(response.content)

        image_url = settings.MEDIA_URL + 'sip_screenshots/' + filename
        return JsonResponse({'success': True, 'image_url': image_url})

    except Exception as e:
        return JsonResponse({'success': False, 'message': f'Ошибка: {e}'})

# === 🎛 Отправка команд SIP-телефону ===
@csrf_exempt
@login_required
def send_sip_command(request):
    """
    Отправка команд SIP-телефону (Reboot, AutoP, Reset и др.)
    """
    logger.info(f"SIP command request: method={request.method}, POST={request.POST}")
    
    if request.method != "POST":
        return JsonResponse({"success": False, "message": "Метод не разрешён."})

    employee_id = request.POST.get("employee_id")
    equipment_id = request.POST.get("equipment_id")
    cmd = request.POST.get("cmd")
    login = request.POST.get("login") or "admin"
    password = request.POST.get("password") or "admin"

    if not employee_id or not cmd:
        return JsonResponse({"success": False, "message": "Не указаны обязательные параметры."})

    try:
        employee = Employee.objects.prefetch_related("equipment_set").get(id=employee_id)
        
        # Получаем конкретное оборудование или первое SIP
        if equipment_id:
            sip_phone = employee.equipment_set.filter(id=equipment_id, type="SIP").first()
        else:
            sip_phone = employee.equipment_set.filter(type="SIP").first()
            
        if not sip_phone:
            return JsonResponse({"success": False, "message": "У сотрудника нет SIP-телефона."})

        # Извлекаем IP
        ip_match = re.search(r"http://([^/]+)", sip_phone.ip_or_anydesk or "")
        if not ip_match:
            return JsonResponse({"success": False, "message": "Не удалось извлечь IP из IP/AnyDesk."})

        ip = ip_match.group(1)
        command_url = f"http://{login}:{password}@{ip}/servlet?key={cmd}"

        response = requests.get(command_url, timeout=10)
        logger.info(f"Отправляется команда на SIP: {command_url}, команда: {cmd}")
        
        if response.status_code == 200:
            return JsonResponse({"success": True, "message": f"✅ Команда {cmd} отправлена успешно."})
        else:
            logger.error(f"Ошибка отправки команды на SIP: {cmd}, URL: {command_url}")
            return JsonResponse({
                "success": False,
                "message": f"⚠️ Ошибка HTTP {response.status_code} при отправке команды.",
            })
    except Exception as e:
        return JsonResponse({"success": False, "message": f"Ошибка: {e}"})

# === 📋 Получение информации о SIP-телефоне ===
@login_required
def get_sip_phone_info(request):
    """
    Возвращает информацию о SIP-телефоне сотрудника (включая IP).
    """
    if request.method == 'POST':
        employee_id = request.POST.get('employee_id')
        equipment_id = request.POST.get('equipment_id')
        
        try:
            employee = Employee.objects.prefetch_related('equipment_set').get(id=employee_id, status='active')
            
            # Получаем конкретное оборудование или первое SIP
            if equipment_id:
                sip_phone = employee.equipment_set.filter(id=equipment_id, type='SIP').first()
            else:
                sip_phone = employee.equipment_set.filter(type='SIP').first()
                
            if not sip_phone:
                return JsonResponse({
                    'success': False,
                    'message': 'У сотрудника нет SIP-телефона.'
                })

            # Возвращаем данные
            return JsonResponse({
                'success': True,
                'sip_phone': {
                    'id': sip_phone.id,
                    'model': sip_phone.model,
                    'mac_address': sip_phone.mac_address,
                    'ip_or_anydesk': sip_phone.ip_or_anydesk,
                    'serial_number': sip_phone.serial_number,
                }
            })
        except Employee.DoesNotExist:
            return JsonResponse({
                'success': False,
                'message': 'Сотрудник не найден или не активен.'
            })
    return JsonResponse({
        'success': False,
        'message': 'Метод не разрешён'
    })

# === ✏️ Inline редактирование оборудования ===
@login_required
@require_POST
def update_equipment_inline(request):
    """
    AJAX: обновить поля equipment: model, serial_number, mac_address, ip_or_anydesk
    Ожидает form-data: id, model, serial_number, mac_address, ip_or_anydesk
    """
    
    eq_id = request.POST.get('id')
    if not eq_id:
        return JsonResponse({'success': False, 'message': 'Не указан id оборудования'}, status=400)
    
    try:
        eq = Equipment.objects.get(id=eq_id)
    except Equipment.DoesNotExist:
        return JsonResponse({'success': False, 'message': 'Оборудование не найдено'}, status=404)

    # Получаем и обрезаем значения
    model = request.POST.get('model', '').strip()
    serial = request.POST.get('serial_number', '').strip()
    mac = request.POST.get('mac_address', '').strip()
    ip = request.POST.get('ip_or_anydesk', '').strip()

    # Сохраняем изменения
    eq.model = model or None
    eq.serial_number = serial or None
    eq.mac_address = mac or None
    eq.ip_or_anydesk = ip or None
    eq.save()

    updated = {
        'model': eq.model or '',
        'serial_number': eq.serial_number or '',
        'mac_address': eq.mac_address or '',
        'ip_or_anydesk': eq.ip_or_anydesk or '',
    }

    return JsonResponse({'success': True, 'updated': updated})

# === 📧 Заглушки для будущего функционала ===
@login_required
def provisioning_create_ad_account(request, employee_id):
    """
    Заглушка для создания учетной записи в AD.
    """
    if request.method == 'POST':
        try:
            employee = Employee.objects.get(id=employee_id, status='active')
            # TODO: Реализовать логику создания учетной записи в AD
            message = f"Учетная запись в AD для {employee.fio} создана."
            return JsonResponse({'success': True, 'message': message})
        except Employee.DoesNotExist:
            return JsonResponse({'success': False, 'message': 'Сотрудник не найден или не активен.'})
    return JsonResponse({'success': False, 'message': 'Метод не разрешён'})

@login_required
def provisioning_send_welcome_email(request, employee_id):
    """
    Заглушка для отправки приветственного email.
    """
    print("views_provisioning.py: provisioning_send_welcome_email loaded")
    # TODO: Реализовать логику отправки email
    return JsonResponse({'success': False, 'message': 'Функция в разработке'})