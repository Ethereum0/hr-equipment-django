# staff/forms.py

from django import forms
from .models import Employee, Equipment

# staff/forms.py

from django import forms
from .models import Employee

class EmployeeForm(forms.ModelForm):
    class Meta:
        model = Employee
        fields = '__all__'
        widgets = {
            'office': forms.Select(choices=Employee.OFFICE_CHOICES),
            'status': forms.Select(choices=Employee.STATUS_CHOICES),
            'hire_date': forms.DateInput(attrs={'type': 'date'}),
            'dismissal_date': forms.DateInput(attrs={'type': 'date'}),
            # Добавляем атрибуты для JavaScript
            'ad_login': forms.TextInput(attrs={'id': 'id_ad_login', 'oninput': 'updateFields()'}),
            'email': forms.EmailInput(attrs={'id': 'id_email', 'oninput': 'updateFields()'}),
            'ad_password': forms.PasswordInput(attrs={'id': 'id_ad_password', 'oninput': 'updateFields()'}),
            'email_password': forms.PasswordInput(attrs={'id': 'id_email_password', 'oninput': 'updateFields()'}),
            'b24_login': forms.TextInput(attrs={'id': 'id_b24_login', 'oninput': 'updateFields()'}),
            'b24_password': forms.PasswordInput(attrs={'id': 'id_b24_password', 'oninput': 'updateFields()'}),
        }

class EquipmentForm(forms.ModelForm):
    class Meta:
        model = Equipment
        fields = '__all__'
        widgets = {
            'office': forms.Select(choices=Employee.OFFICE_CHOICES),
            'type': forms.Select(choices=Equipment.EQUIPMENT_TYPES),
            'in_domain': forms.CheckboxInput(),
            'disposed': forms.CheckboxInput(),
        }
