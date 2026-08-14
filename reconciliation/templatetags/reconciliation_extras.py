from django import template

register = template.Library()


@register.filter
def money(value):
    """Format a number with a fixed comma-thousands / period-decimal style,
    independent of LANGUAGE_CODE locale formatting (which for 'ar' swaps the
    separators and disables grouping - confusing for financial figures that
    should read the same way they do in the source Excel/image documents).
    """
    if value is None or value == '':
        return '-'
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if number == int(number):
        return f'{int(number):,}'
    return f'{number:,.2f}'


IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png')


@register.filter
def is_image_ext(file_field):
    """True for an image file, False for anything else (e.g. .xlsx) - lets
    the upload dropzone decide thumbnail-vs-icon preview from the actual
    stored file rather than a caller-supplied flag, now that the pharmacy
    submission slot accepts either a photo or a real spreadsheet.
    """
    name = getattr(file_field, 'name', '') or ''
    return name.lower().endswith(IMAGE_EXTENSIONS)
