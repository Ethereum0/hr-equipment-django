# staff/views_ajax.py

from django.shortcuts import render, get_object_or_404
from django.contrib.auth.decorators import login_required
from .models import Employee, Equipment

from django.http import JsonResponse
from django.core import serializers


@login_required
def employee_equipment_partial(request, employee_id):
    """
    Возвращает частичный HTML с таблицей оборудования для конкретного сотрудника.
    Используется для загрузки в модальное окно.
    """
    employee = get_object_or_404(Employee, pk=employee_id)
    equipment_list = employee.equipment_set.all().order_by('-id')

    return render(request, 'staff/partials/employee_equipment_table.html', {
        'employee': employee,
        'equipment_list': equipment_list,
    })

def employee_equipment_json(request, employee_id):
    """Возвращает оборудование сотрудника в формате JSON"""
    try:
        equipment = Equipment.objects.filter(employee_id=employee_id)
        equipment_data = []
        
        for item in equipment:
            equipment_data.append({
                'type': item.get_type_display() if item.type else '',
                'model': item.model or '',
                'serial_number': item.serial_number or '',
                'inventory_number': item.inventory_number or '',
                'status': item.get_status_display() if item.status else '',
                'issue_date': item.issue_date.strftime('%d.%m.%Y') if item.issue_date else '',
                'comment': item.comment or ''
            })
        
        return JsonResponse(equipment_data, safe=False)
        
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)