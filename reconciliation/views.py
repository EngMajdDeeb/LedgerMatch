import datetime
import os
from urllib.parse import quote

from django.core.files.base import ContentFile
from django.http import Http404, HttpResponse, HttpResponseBadRequest
from django.shortcuts import get_object_or_404, render
from django.template.loader import render_to_string
from django.utils import timezone
from django.views.decorators.http import require_POST

from .models import ComparisonResult, DailyReconciliation, Pharmacy
from .services import comparison_service, excel_export_service, excel_service, pharmacy_excel_service, vision_service
from .services.excel_service import ExcelParsingError
from .services.openai_client import AIServiceError
from .services.pharmacy_excel_service import PharmacyExcelParsingError
from .validators import validate_excel_upload, validate_pharmacy_file_upload

PHARMACY_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}


def _today():
    return timezone.localdate()


def _parse_date(date_str):
    try:
        return datetime.date.fromisoformat(date_str)
    except (TypeError, ValueError):
        raise Http404('تاريخ غير صالح')


def _selected_date_from_request(request):
    raw = request.GET.get('date')
    if not raw:
        return _today()
    try:
        return datetime.date.fromisoformat(raw)
    except ValueError:
        return _today()


def _get_or_create_reconciliation(pharmacy, date):
    obj, _ = DailyReconciliation.objects.get_or_create(pharmacy=pharmacy, date=date)
    return obj


def _build_dashboard_context(selected_date):
    pharmacies = list(Pharmacy.objects.order_by('name'))
    reconciliations = {
        r.pharmacy_id: r
        for r in DailyReconciliation.objects.filter(date=selected_date).select_related('comparison_result')
    }
    cards = [{'pharmacy': p, 'reconciliation': reconciliations.get(p.id)} for p in pharmacies]

    result_cards = [
        c for c in cards
        if c['reconciliation'] and c['reconciliation'].status in (
            DailyReconciliation.Status.ANALYZED, DailyReconciliation.Status.ERROR,
        )
    ]
    analyzed = [c for c in result_cards if c['reconciliation'].status == DailyReconciliation.Status.ANALYZED]
    matched_count = sum(1 for c in analyzed if not c['reconciliation'].comparison_result.has_differences)
    diff_count = sum(1 for c in analyzed if c['reconciliation'].comparison_result.has_differences)
    error_count = len(result_cards) - len(analyzed)

    return {
        'today': _today(),
        'selected_date': selected_date,
        'is_today': selected_date == _today(),
        'prev_date': selected_date - datetime.timedelta(days=1),
        'next_date': selected_date + datetime.timedelta(days=1),
        'cards': cards,
        'result_cards': result_cards,
        'matched_count': matched_count,
        'diff_count': diff_count,
        'error_count': error_count,
    }


def dashboard(request):
    selected_date = _selected_date_from_request(request)
    return render(request, 'reconciliation/dashboard.html', _build_dashboard_context(selected_date))


def download_combined_excel(request):
    """One "ميزان مراجعة" workbook combining every analyzed pharmacy's
    closing-balance comparison for the selected day - see
    excel_export_service.build_combined_results_excel.
    """
    selected_date = _selected_date_from_request(request)
    reconciliations = (
        DailyReconciliation.objects
        .filter(date=selected_date, status=DailyReconciliation.Status.ANALYZED)
        .select_related('pharmacy', 'comparison_result')
        .order_by('pharmacy__name')
    )

    rows = []
    for r in reconciliations:
        result = r.comparison_result
        totals = result.totals_diff or {}
        account_code = (result.raw_excel_table or {}).get('account_code')
        rows.append({
            'account_code': account_code,
            'pharmacy_name': r.pharmacy.name,
            'bayan_balance': totals.get('bayan_closing_balance'),
            'report_balance': totals.get('image_closing_balance'),
            'note': '',
        })

    buffer = excel_export_service.build_combined_results_excel(selected_date, rows)
    filename = f"ميزان مراجعة - {selected_date.isoformat()}.xlsx"
    response = HttpResponse(
        buffer.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f"attachment; filename*=UTF-8''{quote(filename)}"
    return response


@require_POST
def add_pharmacy(request):
    name = (request.POST.get('name') or '').strip()
    if not name:
        return HttpResponseBadRequest('اسم الصيدلية مطلوب')
    selected_date = _selected_date_from_request(request)
    pharmacy, _ = Pharmacy.objects.get_or_create(name=name)
    reconciliation = DailyReconciliation.objects.filter(
        pharmacy=pharmacy, date=selected_date
    ).select_related('comparison_result').first()
    context = {'pharmacy': pharmacy, 'reconciliation': reconciliation, 'view_date': selected_date}
    return render(request, 'reconciliation/_pharmacy_card.html', context)


@require_POST
def upload_image(request, pharmacy_id, date_str):
    pharmacy = get_object_or_404(Pharmacy, pk=pharmacy_id)
    view_date = _parse_date(date_str)
    reconciliation = _get_or_create_reconciliation(pharmacy, view_date)
    file = request.FILES.get('image_file')
    error = validate_pharmacy_file_upload(file)
    if error:
        context = {'pharmacy': pharmacy, 'reconciliation': reconciliation, 'view_date': view_date, 'upload_error': error}
        return render(request, 'reconciliation/_pharmacy_card.html', context)

    reconciliation.image_file = file
    _reset_stale_result(reconciliation)
    reconciliation.save()
    context = {'pharmacy': pharmacy, 'reconciliation': reconciliation, 'view_date': view_date}
    return render(request, 'reconciliation/_pharmacy_card.html', context)


@require_POST
def upload_excel(request, pharmacy_id, date_str):
    pharmacy = get_object_or_404(Pharmacy, pk=pharmacy_id)
    view_date = _parse_date(date_str)
    reconciliation = _get_or_create_reconciliation(pharmacy, view_date)
    file = request.FILES.get('excel_file')
    error = validate_excel_upload(file)
    if error:
        context = {'pharmacy': pharmacy, 'reconciliation': reconciliation, 'view_date': view_date, 'upload_error': error}
        return render(request, 'reconciliation/_pharmacy_card.html', context)

    reconciliation.excel_file = file
    _reset_stale_result(reconciliation)
    reconciliation.save()
    context = {'pharmacy': pharmacy, 'reconciliation': reconciliation, 'view_date': view_date}
    return render(request, 'reconciliation/_pharmacy_card.html', context)


def _reset_stale_result(reconciliation):
    """A fresh upload invalidates any previous analysis for this reconciliation."""
    if reconciliation.pk and hasattr(reconciliation, 'comparison_result'):
        reconciliation.comparison_result.delete()
    reconciliation.status = DailyReconciliation.Status.PENDING
    reconciliation.error_message = ''


def _extract_pharmacy_submission(reconciliation):
    """Route to the right extraction path based on what the pharmacy actually
    uploaded: a real spreadsheet is parsed deterministically (more reliable
    than OCR), a photo goes through the vision pipeline.
    """
    ext = os.path.splitext(reconciliation.image_file.name)[1].lower()
    if ext in PHARMACY_EXCEL_EXTENSIONS:
        return pharmacy_excel_service.parse_pharmacy_excel(
            reconciliation.image_file.path, pharmacy_name=reconciliation.pharmacy.name
        )

    reconciliation.image_file.open('rb')
    try:
        return vision_service.extract_image_table(reconciliation.image_file)
    finally:
        reconciliation.image_file.close()


@require_POST
def run_comparison(request, reconciliation_id):
    reconciliation = get_object_or_404(
        DailyReconciliation.objects.select_related('pharmacy'), pk=reconciliation_id
    )

    if not reconciliation.ready_to_compare:
        return _fail(request, reconciliation, 'الرجاء رفع الصورة وملف الإكسل أولا قبل المطابقة.', persist=False)

    try:
        excel_table = excel_service.parse_el_bayan_excel(reconciliation.excel_file.path)
        image_extraction = _extract_pharmacy_submission(reconciliation)
        comparison = comparison_service.run_comparison(excel_table, image_extraction)
    except (ExcelParsingError, PharmacyExcelParsingError, AIServiceError) as exc:
        return _fail(request, reconciliation, str(exc))
    except Exception:
        return _fail(request, reconciliation, 'حدث خطأ غير متوقع أثناء المطابقة. الرجاء المحاولة مرة أخرى.')

    audit_buffer = excel_export_service.build_audit_excel(image_extraction, comparison)

    result, _ = ComparisonResult.objects.update_or_create(
        reconciliation=reconciliation,
        defaults={
            'differences': comparison['categories'],
            'matched_rows_count': comparison['matched_rows_count'],
            'mismatched_rows_count': comparison['mismatched_rows_count'],
            'missing_in_image': comparison['missing_in_image'],
            'missing_in_excel': comparison['missing_in_excel'],
            'totals_diff': comparison['totals_diff'],
            'summary_ar': comparison['summary_ar'],
            'raw_image_extraction': image_extraction,
            'raw_excel_table': excel_table,
            'has_differences': comparison['has_differences'],
        },
    )
    # The actual stored filename is fully determined by converted_excel_upload_path
    # (pharmacy name + date); the name passed here is only used to infer the extension.
    result.converted_excel_file.save(
        'converted.xlsx',
        ContentFile(audit_buffer.read()),
        save=True,
    )

    reconciliation.status = DailyReconciliation.Status.ANALYZED
    reconciliation.error_message = ''
    reconciliation.save(update_fields=['status', 'error_message', 'updated_at'])

    return _respond_with_updated_state(request, reconciliation, result=result)


def _fail(request, reconciliation, message, persist=True):
    if persist:
        reconciliation.status = DailyReconciliation.Status.ERROR
        reconciliation.error_message = message
        reconciliation.save(update_fields=['status', 'error_message', 'updated_at'])
        return _respond_with_updated_state(request, reconciliation, error=message)

    # Not ready to compare yet - don't persist an error state, just report it inline.
    pharmacy_card_html = render_to_string(
        request=request,
        template_name='reconciliation/_pharmacy_card.html',
        context={
            'pharmacy': reconciliation.pharmacy, 'reconciliation': reconciliation,
            'view_date': reconciliation.date, 'upload_error': message,
        },
    )
    return HttpResponse(pharmacy_card_html)


def _respond_with_updated_state(request, reconciliation, error=None, result=None):
    """Primary swap = the pharmacy's upload card; OOB swap = *only this
    pharmacy's* result slot.

    "Compare all" fires one independent request per ready pharmacy, and
    those requests can finish and arrive back at the browser in any order
    (only a few worker processes handle them, in batches). An earlier design
    had each response regenerate and swap in the *entire* results section as
    a snapshot - whichever response happened to arrive last would overwrite
    everyone else's results with whatever the database looked like at that
    response's own snapshot time, silently discarding results that had
    already rendered correctly. Touching only this pharmacy's own DOM slot
    makes every response's update independent and order-safe: it can never
    clobber another pharmacy's result, no matter what order they arrive in.
    """
    pharmacy_card_html = render_to_string(
        request=request,
        template_name='reconciliation/_pharmacy_card.html',
        context={'pharmacy': reconciliation.pharmacy, 'reconciliation': reconciliation, 'view_date': reconciliation.date},
    )
    result_slot_html = render_to_string(
        request=request,
        template_name='reconciliation/_result_slot.html',
        context={
            'pharmacy': reconciliation.pharmacy,
            'reconciliation': reconciliation,
            'error': error,
            'result': result,
        },
    )
    return HttpResponse(pharmacy_card_html + result_slot_html)
