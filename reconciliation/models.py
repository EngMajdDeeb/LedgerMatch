import re

from django.db import models
from django.utils import timezone

_UNSAFE_PATH_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_folder_name(name: str) -> str:
    """Sanitize a pharmacy name for use as a filesystem folder name.

    Keeps Arabic/Latin letters, digits, spaces and dashes (all valid on
    modern filesystems including NTFS) and only strips characters that are
    actually illegal in Windows paths.
    """
    cleaned = _UNSAFE_PATH_CHARS.sub('', name).strip()
    return cleaned or 'صيدلية'


def _dated_folder(pharmacy_name: str, date) -> str:
    date = date or timezone.now().date()
    return f"{safe_folder_name(pharmacy_name)}/{date.strftime('%Y-%m')}/{date.strftime('%d')}"


def _dated_filename(pharmacy_name: str, date, label: str, ext: str) -> str:
    date = date or timezone.now().date()
    base = f"{safe_folder_name(pharmacy_name)} - {date.isoformat()} - {label}"
    return f"{base}.{ext}"


def _upload_path(instance, filename, label):
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'bin'
    folder = _dated_folder(instance.pharmacy.name, instance.date)
    name = _dated_filename(instance.pharmacy.name, instance.date, label, ext)
    return f"{folder}/{name}"


def image_upload_path(instance, filename):
    return _upload_path(instance, filename, 'تقرير الصندوق')


def excel_upload_path(instance, filename):
    return _upload_path(instance, filename, 'بيان محاسبي')


def converted_excel_upload_path(instance, filename):
    reconciliation = instance.reconciliation
    folder = _dated_folder(reconciliation.pharmacy.name, reconciliation.date)
    name = _dated_filename(reconciliation.pharmacy.name, reconciliation.date, 'نسخة اكسل من الصورة', 'xlsx')
    return f"{folder}/{name}"


class Pharmacy(models.Model):
    name = models.CharField(max_length=255, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'صيدلية'
        verbose_name_plural = 'الصيدليات'

    def __str__(self):
        return self.name


class DailyReconciliation(models.Model):
    class Status(models.TextChoices):
        PENDING = 'pending', 'قيد الانتظار'
        ANALYZED = 'analyzed', 'تمت المطابقة'
        ERROR = 'error', 'حدث خطأ'

    pharmacy = models.ForeignKey(Pharmacy, related_name='reconciliations', on_delete=models.CASCADE)
    date = models.DateField(default=timezone.localdate)
    # Despite the field name, this holds whatever the pharmacy actually
    # submitted for the day's cash-box sheet - a photo (read via vision_service)
    # or a real spreadsheet in the same layout (read via pharmacy_excel_service).
    # views.run_comparison picks the extraction path from the file's extension.
    image_file = models.FileField(upload_to=image_upload_path, blank=True, null=True)
    excel_file = models.FileField(upload_to=excel_upload_path, blank=True, null=True)
    status = models.CharField(max_length=20, choices=Status.choices, default=Status.PENDING)
    error_message = models.TextField(blank=True, default='')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-date', 'pharmacy__name']
        constraints = [
            models.UniqueConstraint(fields=['pharmacy', 'date'], name='unique_pharmacy_per_day'),
        ]
        verbose_name = 'تسوية يومية'
        verbose_name_plural = 'التسويات اليومية'

    def __str__(self):
        return f"{self.pharmacy.name} - {self.date}"

    @property
    def ready_to_compare(self):
        return bool(self.image_file) and bool(self.excel_file)


class ComparisonResult(models.Model):
    reconciliation = models.OneToOneField(
        DailyReconciliation, related_name='comparison_result', on_delete=models.CASCADE
    )

    # Full structured output of the AI comparison call.
    differences = models.JSONField(default=list, blank=True)
    matched_rows_count = models.PositiveIntegerField(default=0)
    mismatched_rows_count = models.PositiveIntegerField(default=0)
    missing_in_image = models.JSONField(default=list, blank=True)
    missing_in_excel = models.JSONField(default=list, blank=True)
    totals_diff = models.JSONField(default=dict, blank=True)
    summary_ar = models.TextField(blank=True, default='')

    # Debugging / audit trail of the two upstream AI/pandas calls.
    raw_image_extraction = models.JSONField(default=dict, blank=True)
    raw_excel_table = models.JSONField(default=dict, blank=True)

    # Downloadable .xlsx rebuilt from the image's extracted JSON, for
    # manual spot-checking next to the real El-Bayan file.
    converted_excel_file = models.FileField(
        upload_to=converted_excel_upload_path, blank=True, null=True
    )

    has_differences = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'نتيجة مطابقة'
        verbose_name_plural = 'نتائج المطابقة'

    def __str__(self):
        return f"نتيجة {self.reconciliation}"
