import os

from django.conf import settings

ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png'}
ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}


def validate_pharmacy_file_upload(file):
    """The pharmacy's daily cash-box report: either a photo of the sheet or
    a real spreadsheet in the same layout - vision_service handles the
    former, pharmacy_excel_service the latter.
    """
    if file is None:
        return 'الرجاء اختيار صورة أو ملف إكسل لتقرير الصندوق.'
    ext = os.path.splitext(file.name)[1].lower()
    if ext not in ALLOWED_IMAGE_EXTENSIONS and ext not in ALLOWED_EXCEL_EXTENSIONS:
        return 'صيغة الملف غير مدعومة. الرجاء رفع صورة (JPG/PNG) أو ملف إكسل (XLSX/XLS).'
    if file.size > settings.MAX_IMAGE_UPLOAD_SIZE:
        max_mb = settings.MAX_IMAGE_UPLOAD_SIZE // (1024 * 1024)
        return f'حجم الملف أكبر من الحد المسموح ({max_mb} ميغابايت).'
    return None


def validate_excel_upload(file):
    if file is None:
        return 'الرجاء اختيار ملف إكسل.'
    ext = os.path.splitext(file.name)[1].lower()
    if ext not in ALLOWED_EXCEL_EXTENSIONS:
        return 'صيغة الملف غير مدعومة. الرجاء رفع ملف Excel بصيغة XLSX أو XLS فقط.'
    if file.size > settings.MAX_EXCEL_UPLOAD_SIZE:
        max_mb = settings.MAX_EXCEL_UPLOAD_SIZE // (1024 * 1024)
        return f'حجم الملف أكبر من الحد المسموح ({max_mb} ميغابايت).'
    return None
