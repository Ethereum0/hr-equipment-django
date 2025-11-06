# staff/urls.py
from django.urls import path
from . import views
from . import views_ajax
from . import views_provisioning
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    # === Основные маршруты ===
    path('', views.dashboard, name='dashboard'),
    path('employees/', views.employees_list, name='employees_list'),
    path('employee/new/', views.employee_create, name='employee_create'),
    path('employee/<int:pk>/edit/', views.employee_edit, name='employee_edit'),
    path('employee/<int:pk>/change-status/', views.employee_change_status, name='employee_change_status'),
    
    # === Оборудование ===
    path('equipment/', views.equipment_list, name='equipment_list'),
    path('equipment/new/', views.equipment_create, name='equipment_create'),
    path('equipment/<int:pk>/edit/', views.equipment_edit, name='equipment_edit'),
    path('equipment/<int:pk>/delete/', views.equipment_delete, name='equipment_delete'),
    path('equipment/<int:pk>/assign/', views.equipment_assign_employee, name='equipment_assign_employee'),
    
    # === AJAX маршруты ===
    path('employee/<int:employee_id>/equipment/', views_ajax.employee_equipment_partial, name='employee_equipment_partial'),
    
    # === Provisioning маршруты ===
    path('provisioning/', views_provisioning.provisioning_dashboard, name='provisioning_dashboard'),
    path('provisioning/generate_config/', views_provisioning.generate_config, name='generate_config'),
    path('provisioning/sip_screen/', views_provisioning.provisioning_sip_screen, name='provisioning_sip_screen'),
    path('provisioning/send_sip_command/', views_provisioning.send_sip_command, name='send_sip_command'),
    path('provisioning/get_sip_phone_info/', views_provisioning.get_sip_phone_info, name='get_sip_phone_info'),
    path('provisioning/update_equipment_inline/', views_provisioning.update_equipment_inline, name='update_equipment_inline'),
    path('provisioning/create_ad_account/<int:employee_id>/', views_provisioning.provisioning_create_ad_account, name='provisioning_create_ad_account'),
    path('provisioning/send_welcome_email/<int:employee_id>/', views_provisioning.provisioning_send_welcome_email, name='provisioning_send_welcome_email'),
    
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)