# staff/models.py
from django.db import models
from django.db.models.signals import pre_save
from django.dispatch import receiver

class Employee(models.Model):
    OFFICE_CHOICES = [
        ('ORL', 'Орёл'),
        ('MSK', 'Москва'),
        ('SKLAD', 'Склад'),
        ('remote', 'Удалёнка'),
        ('ЭМКО', 'ЭМКО'),
        ('ORUS', 'ОРУС'), # Добавлен недостающий офис из файла
    ]

    STATUS_CHOICES = [
        ('active', 'Активный'),
        ('dismissed', 'Уволен'),
        ('maternity', 'Декрет'),
    ]

    # Основная информация
    fio = models.CharField("ФИО", max_length=255)
    first_name = models.CharField("Имя", max_length=100, blank=True)
    last_name = models.CharField("Фамилия", max_length=100, blank=True)
    middle_name = models.CharField("Отчество", max_length=100, blank=True)
    
    # Учетные данные
    ad_login = models.CharField("AD-логин", max_length=100, unique=True, blank=True, null=True)
    ad_password = models.CharField("Пароль AD", max_length=128, blank=True)
    email = models.EmailField("Email", blank=True, null=True)
    email_password = models.CharField("Пароль Email", max_length=128, blank=True)
    
    # Телефония
    sip_number = models.CharField("SIP-номер", max_length=20, blank=True)
    web_3cx = models.CharField("3CX WEB", max_length=255, blank=True) # Новое поле
    pass_3cx = models.CharField("Пароль 3CX", max_length=128, blank=True) # Новое поле
    
    # Битрикс24
    b24_login = models.CharField("Логин Б24", max_length=100, blank=True)
    b24_password = models.CharField("Пароль Б24", max_length=128, blank=True)
    
    # Дополнительная информация
    phone = models.CharField("Телефон", max_length=20, blank=True)
    position = models.CharField("Должность", max_length=255, blank=True)
    department = models.CharField("Отдел", max_length=255, blank=True)
    supervisor = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Руководитель")
    info = models.TextField("Инфо", blank=True) # Новое поле
    
    # Офис и статус
    office = models.CharField("Офис", max_length=20, choices=OFFICE_CHOICES)
    status = models.CharField("Статус", max_length=20, choices=STATUS_CHOICES, default='active')
    
    # Новые поля для телефонии и Б24 из файла
    tel_b24_login = models.CharField("Тел Б24 логин", max_length=100, blank=True) # Новое поле
    tel_b24_password = models.CharField("Тел Б24 пароль", max_length=128, blank=True) # Новое поле
    vats_login = models.CharField("Ватс логин", max_length=100, blank=True) # Новое поле
    vats_password = models.CharField("ВАТС пароль", max_length=128, blank=True) # Новое поле

    # Даты (опционально)
    hire_date = models.DateField("Дата приёма", null=True, blank=True)
    dismissal_date = models.DateField("Дата увольнения/выхода в декрет", null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.fio

    @property
    def photo_url(self):
        if self.ad_login:
            return f"https://p-el.ru/mail-photos/{self.ad_login}.jpg"
        return None

    def get_status_badge_class(self):
        return {
            'active': 'bg-green-100 text-green-800',
            'dismissed': 'bg-red-100 text-red-800',
            'maternity': 'bg-orange-100 text-orange-800',
        }.get(self.status, 'bg-gray-100 text-gray-800')

    def as_markdown(self):
        lines = [
            f"**ФИО**: {self.fio}",
            f"**Должность**: {self.position or '—'}",
            f"**Офис**: {dict(Employee.OFFICE_CHOICES).get(self.office, self.office)}",
            f"**Телефон**: {self.phone or '—'}",
            "",
            "### Учётные данные",
            f"- **AD**: `CORP\\{self.ad_login}`" if self.ad_login else "- **AD**: —",
            f"- **Пароль AD**: `{self.ad_password}`" if self.ad_password else "- **Пароль AD**: —",
            f"- **Email**: `{self.email}`" if self.email else "- **Email**: —",
            f"- **Пароль Email**: `{self.email_password}`" if self.email_password else "- **Пароль Email**: —",
            f"- **SIP**: `{self.sip_number}`" if self.sip_number else "- **SIP**: —",
            f"- **Б24 логин**: `{self.b24_login}`" if self.b24_login else "- **Б24 логин**: —",
            f"- **Б24 пароль**: `{self.b24_password}`" if self.b24_password else "- **Б24 пароль**: —",
        ]
        return "\n".join(lines)
    

    def laptop_anydesk_id(self):
        """
        Возвращает ID AnyDesk для ноутбука сотрудника.
        Ищет оборудование типа 'Ноутбук' и возвращает значение из поля ip_or_anydesk,
        если оно является числовым ID.
        """
        try:
            # Ищем оборудование типа "Ноутбук", привязанное к сотруднику
            laptop = self.equipment_set.get(type='Ноутбук')
            # Проверяем, что значение ip_or_anydesk является числовым ID
            if laptop.ip_or_anydesk and laptop.ip_or_anydesk.isdigit():
                return laptop.ip_or_anydesk
            else:
                return None
        except Equipment.DoesNotExist:
            # Ноутбук не найден
            return None
        except Equipment.MultipleObjectsReturned:
            # Найдено несколько ноутбуков, возвращаем ID первого
            laptops = self.equipment_set.filter(type='Ноутбук')
            for laptop in laptops:
                if laptop.ip_or_anydesk and laptop.ip_or_anydesk.isdigit():
                    return laptop.ip_or_anydesk
            return None

class Equipment(models.Model):
    EQUIPMENT_TYPES = [
        ('SIP', 'SIP-телефон'),
        ('Монитор', 'Монитор'),
        ('Ноутбук', 'Ноутбук'),
        ('Android', 'Android-устройство'),
        ('PC', 'Стационарный ПК'),
        ('-', 'Прочее'),
    ]

    employee = models.ForeignKey(Employee, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Сотрудник")
    type = models.CharField("Тип ТМЦ", max_length=50, choices=EQUIPMENT_TYPES)
    model = models.CharField("Модель", max_length=255, blank=True)
    serial_number = models.CharField("Серийный номер", max_length=255, unique=True, blank=True, null=True)
    mac_address = models.CharField("MAC-адрес", max_length=17, blank=True)
    ip_or_anydesk = models.CharField("IP / AnyDesk", max_length=100, blank=True)
    comment = models.TextField("Комментарий", blank=True)
    in_domain = models.BooleanField("В домене", default=False)
    office = models.CharField("Офис", max_length=20, choices=Employee.OFFICE_CHOICES)
    position = models.CharField("Должность", max_length=255, blank=True) # Добавлено ранее
    disposed = models.BooleanField("Списано", default=False) # Добавлено ранее

    def __str__(self):
        return f"{self.type} ({self.serial_number or 'б/н'})"

    def as_markdown(self):
        lines = [
            f"**Тип**: {self.get_type_display()}",
            f"**Модель**: {self.model or '—'}",
            f"**Серийный номер**: {self.serial_number or '—'}",
            f"**MAC-адрес**: {self.mac_address or '—'}",
            f"**IP / AnyDesk**: {self.ip_or_anydesk or '—'}",
            f"**Комментарий**: {self.comment or '—'}",
            f"**В домене**: {'Да' if self.in_domain else 'Нет'}",
            f"**Офис**: {dict(Employee.OFFICE_CHOICES).get(self.office, self.office)}",
            f"**Сотрудник**: {self.employee.fio if self.employee else 'Свободное'}",
        ]
        return "\n".join(lines)
# === Сигналы ===

@receiver(pre_save, sender=Employee)
def free_equipment_on_status_change(sender, instance, **kwargs):
    """
    Освобождает оборудование сотрудника при изменении статуса на 'dismissed' или 'maternity'.
    """
    if instance.pk:  # Проверяем, что это обновление, а не создание
        try:
            old_instance = Employee.objects.get(pk=instance.pk)
            # Если статус изменился на 'dismissed' или 'maternity'
            if instance.status in ['dismissed', 'maternity'] and old_instance.status != instance.status:
                # Освобождаем все оборудование, привязанное к сотруднику
                Equipment.objects.filter(employee=instance).update(employee=None)
                print(f"Оборудование сотрудника {instance.fio} освобождено.")
        except Employee.DoesNotExist:
            pass  # Это новый сотрудник, ничего не делаем    