# staff/templatetags/form_extras.py

from django import template
from django.forms.widgets import Widget

register = template.Library()

@register.filter
def add_class(field, css_class):
    """
    Добавляет CSS-класс к полю формы.
    Использование: {{ form.field|add_class:"my-class" }}
    """
    if hasattr(field, 'as_widget') and callable(field.as_widget):
        # Это поле формы
        widget = field.field.widget
        existing_classes = widget.attrs.get('class', '')
        if existing_classes:
            css_class = f"{existing_classes} {css_class}"
        return field.as_widget(attrs={"class": css_class})
    else:
        # Это не поле формы, возвращаем как есть
        return field
def get_item(dictionary, key):
    """
    Фильтр для получения значения из словаря по ключу в шаблонах.
    Использование: {{ my_dict|get_item:key_variable }}
    """
    return dictionary.get(key)    