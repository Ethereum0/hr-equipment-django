# staff/views_ajax.py

from django.shortcuts import render, get_object_or_404
from django.contrib.auth.decorators import login_required
from .models import Employee, Equipment

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