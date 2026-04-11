from django import template

register = template.Library()


@register.filter
def get_item(dictionary, key):
    """Return dict[key], works with string and int keys."""
    if not isinstance(dictionary, dict):
        return ''
    return dictionary.get(str(key), dictionary.get(key, ''))


@register.filter
def get_dict(dictionary, key):
    """Alias for get_item — used to look up work type labels."""
    if not isinstance(dictionary, dict):
        return key
    return dictionary.get(str(key), dictionary.get(key, key))
