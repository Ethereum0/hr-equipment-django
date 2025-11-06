# staff/templatetags/form_extras.py

from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """
    Фильтр для получения значения из словаря по ключу в шаблонах.
    Использование: {{ my_dict|get_item:key_variable }}
    """
    return dictionary.get(key)