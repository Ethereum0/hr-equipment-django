# staff/admin.py

from django.contrib import admin
from .models import Employee, Equipment
from .forms import EmployeeForm

@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    form = EmployeeForm  # Используем нашу кастомную форму
    list_display = ('fio', 'position', 'office', 'status', 'phone', 'ad_login')
    list_filter = ('office', 'status', 'hire_date', 'dismissal_date')
    search_fields = ('fio', 'ad_login', 'email', 'phone')
    list_editable = ('status',)

@admin.register(Equipment)
class EquipmentAdmin(admin.ModelAdmin):
    list_display = ('type', 'model', 'serial_number', 'employee', 'office', 'disposed')
    list_filter = ('type', 'office', 'in_domain', 'disposed')
    search_fields = ('serial_number', 'model', 'mac_address', 'employee__fio')
    raw_id_fields = ('employee',)