"""
Step 3 of the pipeline: the actual comparison engine.

This is the only place that decides "does El-Bayan match the image". It is
a single OpenAI text call that receives two already-cleaned JSON tables -
the El-Bayan ledger (from excel_service) and the image's category totals
(from vision_service + aggregation) - and returns a structured category by
category diff plus an Arabic narrative summary.

The only work Python does before/after this call is:
  * handing the model plain facts it cannot get wrong by itself (the
    deterministic opening/closing balances already parsed from the ledger,
    and the arithmetic discrepancy between the two closing balances) so it
    doesn't have to re-derive them from 39+ transaction rows,
  * filtering/counting the model's own `status` field per category into the
    matched/mismatched/missing_* buckets the UI renders - it never
    re-judges what the model decided, and
  * attaching each category's `image_note` (the pharmacy's handwritten note
    for that row, read by vision_service's third call) - a plain dict
    lookup, not something the comparison model needs to reason about.
"""
from __future__ import annotations

import json

from django.conf import settings
from openai import APIError, APITimeoutError, OpenAIError

from ..constants import ALL_CATEGORY_KEYS, CATEGORY_LABELS_AR, FLOW_CATEGORY_KEYS
from .aggregation import aggregate_image_totals
from .openai_client import AIServiceError, friendly_api_error, get_client

STATUS_MATCH = 'match'
STATUS_MISMATCH = 'mismatch'
STATUS_MISSING_IN_BAYAN = 'missing_in_bayan'
STATUS_MISSING_IN_IMAGE = 'missing_in_image'
STATUS_NOT_APPLICABLE = 'not_applicable'
VALID_STATUSES = [
    STATUS_MATCH, STATUS_MISMATCH, STATUS_MISSING_IN_BAYAN,
    STATUS_MISSING_IN_IMAGE, STATUS_NOT_APPLICABLE,
]

SYSTEM_PROMPT = """أنت محاسب خبير تُطابق صندوق صيدلية بين مصدرين لنفس اليوم:
1) "el_bayan": كشف حساب حقيقي (قيود محاسبية تفصيلية) من برنامج البيان المحاسبي، وهو مصدر الحقيقة المحاسبي.
2) "image": جدول تسوية يومي أرسلته الصيدلية (مُستخرج من صورة فوتوغرافية)، يحتوي مجاميع جاهزة لكل فئة.

مهمتك بخطوتين:

الخطوة الأولى - تصنيف قيود el_bayan.transactions إلى الفئات العشر التالية (بالاعتماد على وصف كل قيد "description" وعمودي مدين/دائن):
- sales (مبيعات): قيود مدينة "debit" ضمن وصفها "فاتورة بيع" (وليست مردودة).
- sales_returns (مرتجع مبيعات): قيود دائنة "credit" ضمن وصفها "مردود بيع" أو "مرتجع بيع".
- cash_purchases (مشتريات / نقدا): الجزء المدفوع نقدا فورا من مشتريات/دفعات لموردين (غالبا قيود دائنة). إن ذكر الوصف "باقي" أو "متبقي" فهذا يعني أن جزءا من المبلغ آجل - افصل الجزء المدفوع فورا (نقدي) عن الجزء المتبقي (آجل).
- credit_purchases (مشتريات / آجل): الجزء الآجل (غير المدفوع بعد) من نفس نوع القيود أعلاه، كما يظهر من عبارات مثل "باقي 4545".
- purchase_returns (مرتجع مشتريات): قيود مدينة لإرجاع بضاعة لمورد.
- supplier_payments (دفعات موردين): تسديد رصيد سابق لمورد لا يرتبط بفاتورة شراء ضمن هذا الكشف نفسه.
- company_funding (تمويل شركة): تمويل/دعم وارد من الشركة الأم أو الإدارة للصندوق.
- expenses_1 / expenses_2 / expenses_3 (مصاريف): قيود دائنة لمصاريف تشغيلية عامة (رواتب، كهرباء، صيانة...). ضع أول نوع مصروف واضح في expenses_1، والثاني في expenses_2 إن وجد، وهكذا. اتركها null إن لم تجد ما يقابلها.
إن لم يقع قيد ما بوضوح ضمن أي فئة، تجاهله في المجاميع واذكره باختصار ضمن summary_ar بدل حذفه بصمت.

تنبيه مهم جدا حول الفصل بين الفئات: مرتجع مشتريات (purchase_returns) و مبيعات (sales) فئتان منفصلتان تماما ولا علاقة بينهما إطلاقا،
حتى لو كان القيدان من نفس نوع الحركة (مدين). لا يجوز أن يدخل أي مبلغ من مرتجع مشتريات ضمن مجموع مبيعات، ولا أن يُطرح منه أو
يُضاف إليه بأي شكل - كل قيد يُصنَّف مرة واحدة فقط في الفئة التي يخصها حصرا بحسب وصفه، ولا يُكرَّر أو يُخلط مع فئة أخرى.
بنفس المنطق: تمويل شركة (company_funding) لا يُخلط مع دفعات موردين (supplier_payments)، وكل فئة من الفئات العشر تُحسب
بمعزل تام عن غيرها.

الخطوة الثانية - قارن كل فئة من الفئات الأربع عشرة (العشر أعلاه + opening_balance و closing_balance و bayan_balance و discrepancy):
- بالنسبة لـ opening_balance و closing_balance: استخدم القيم الجاهزة المُعطاة لك في known_facts.el_bayan (لا تُعد حسابها من القيود).
- بالنسبة لـ bayan_balance: قارن القيمة التي كتبتها الصيدلية (known_facts.image.bayan_balance_as_reported_by_pharmacy) مع القيمة الحقيقية known_facts.el_bayan.closing_balance.
- بالنسبة لـ discrepancy: قارن الفرق الذي حسبته الصيدلية بنفسها (known_facts.image.discrepancy_as_reported_by_pharmacy) مع الفرق الصحيح المُعطى known_facts.true_discrepancy_computed.
- تجاهل فروقات التقريب الصغيرة (أقل من 0.01) واعتبرها تطابقا.
- أي فرق رقمي حقيقي يجب ذكره دائما، لا تُسقطه بصمت.
- إن كانت القيمة غير موجودة في مصدر وموجودة بالآخر، استخدم status المناسب (missing_in_bayan أو missing_in_image).
- إن لم تكن الفئة قابلة للتطبيق إطلاقا استخدم not_applicable.

أعد النتيجة حصرا وفق مخطط الإخراج المطلوب (schema)، بدون أي نص خارج الحقول، واكتب summary_ar وكل note_ar باللغة العربية بإيجاز ووضوح."""


def _build_output_schema() -> dict:
    return {
        'type': 'object',
        'properties': {
            'categories': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'key': {'type': 'string', 'enum': ALL_CATEGORY_KEYS},
                        'bayan_value': {'type': ['number', 'null']},
                        'image_value': {'type': ['number', 'null']},
                        'difference': {'type': ['number', 'null']},
                        'status': {'type': 'string', 'enum': VALID_STATUSES},
                        'note_ar': {'type': ['string', 'null']},
                    },
                    'required': ['key', 'bayan_value', 'image_value', 'difference', 'status', 'note_ar'],
                    'additionalProperties': False,
                },
            },
            'totals_diff': {
                'type': 'object',
                'properties': {
                    'bayan_closing_balance': {'type': ['number', 'null']},
                    'image_closing_balance': {'type': ['number', 'null']},
                    'difference': {'type': ['number', 'null']},
                    'note_ar': {'type': ['string', 'null']},
                },
                'required': ['bayan_closing_balance', 'image_closing_balance', 'difference', 'note_ar'],
                'additionalProperties': False,
            },
            'summary_ar': {'type': 'string'},
        },
        'required': ['categories', 'totals_diff', 'summary_ar'],
        'additionalProperties': False,
    }


def _build_known_facts(el_bayan_table: dict, image_extraction: dict) -> dict:
    image_totals = aggregate_image_totals(image_extraction)

    true_discrepancy = None
    if image_totals.get('closing_balance') is not None and el_bayan_table.get('closing_balance') is not None:
        true_discrepancy = round(image_totals['closing_balance'] - el_bayan_table['closing_balance'], 2)

    return {
        'el_bayan': {
            'opening_balance': el_bayan_table.get('opening_balance'),
            'closing_balance': el_bayan_table.get('closing_balance'),
            'transactions': el_bayan_table.get('transactions', []),
        },
        'image': {
            'opening_balance': image_totals.get('opening_balance'),
            'closing_balance': image_totals.get('closing_balance'),
            'bayan_balance_as_reported_by_pharmacy': image_totals.get('bayan_balance'),
            'discrepancy_as_reported_by_pharmacy': image_totals.get('discrepancy'),
            'flow_totals': {key: image_totals.get(key) for key in FLOW_CATEGORY_KEYS},
        },
        'true_discrepancy_computed': true_discrepancy,
    }, image_totals


def _normalize_categories(raw_categories: list, row_notes: dict) -> list:
    """Guarantee exactly one row per known category key, in a stable order.

    `row_notes` is the per-category text read off the image's "الملاحظات"
    column (vision_service's third call) - attached here as plain lookup,
    not passed through the AI comparison call, since copying a string from
    one dict to another needs no judgment.
    """
    by_key = {row.get('key'): row for row in raw_categories if isinstance(row, dict)}
    normalized = []
    for key in ALL_CATEGORY_KEYS:
        row = by_key.get(key, {})
        normalized.append({
            'key': key,
            'label_ar': CATEGORY_LABELS_AR[key],
            'bayan_value': row.get('bayan_value'),
            'image_value': row.get('image_value'),
            'difference': row.get('difference'),
            'status': row.get('status', STATUS_NOT_APPLICABLE),
            'note_ar': row.get('note_ar'),
            'image_note': row_notes.get(key),
        })
    return normalized


def run_comparison(el_bayan_table: dict, image_extraction: dict) -> dict:
    """Run the AI comparison call and return the fully structured result dict.

    Raises AIServiceError on any failure.
    """
    known_facts, image_totals = _build_known_facts(el_bayan_table, image_extraction)
    client = get_client()

    try:
        response = client.responses.create(
            model=settings.OPENAI_TEXT_MODEL,
            input=[
                {'role': 'system', 'content': [{'type': 'input_text', 'text': SYSTEM_PROMPT}]},
                {
                    'role': 'user',
                    'content': [{
                        'type': 'input_text',
                        'text': 'known_facts:\n' + json.dumps(known_facts, ensure_ascii=False),
                    }],
                },
            ],
            text={
                'format': {
                    'type': 'json_schema',
                    'name': 'cashbox_comparison',
                    'schema': _build_output_schema(),
                    'strict': True,
                },
            },
        )
    except APITimeoutError as exc:
        raise AIServiceError('انتهت مهلة الاتصال بخدمة OpenAI أثناء المطابقة.') from exc
    except APIError as exc:
        raise AIServiceError(friendly_api_error(exc, 'المطابقة')) from exc
    except OpenAIError as exc:
        raise AIServiceError(f'تعذر الاتصال بخدمة OpenAI: {exc}') from exc

    output_text = getattr(response, 'output_text', None)
    if not output_text:
        raise AIServiceError('لم يُرجع نموذج الذكاء الاصطناعي نتيجة مطابقة.')

    try:
        result = json.loads(output_text)
    except json.JSONDecodeError as exc:
        raise AIServiceError('تعذر تفسير نتيجة المطابقة (JSON غير صالح).') from exc

    row_notes = image_extraction.get('row_notes', {}) or {}
    categories = _normalize_categories(result.get('categories', []), row_notes)
    matched = [c for c in categories if c['status'] == STATUS_MATCH]
    mismatched = [c for c in categories if c['status'] == STATUS_MISMATCH]
    missing_in_image = [c for c in categories if c['status'] == STATUS_MISSING_IN_IMAGE]
    missing_in_excel = [c for c in categories if c['status'] == STATUS_MISSING_IN_BAYAN]

    return {
        'categories': categories,
        'matched_rows_count': len(matched),
        'mismatched_rows_count': len(mismatched),
        'missing_in_image': missing_in_image,
        'missing_in_excel': missing_in_excel,
        'totals_diff': result.get('totals_diff', {}),
        'summary_ar': result.get('summary_ar', ''),
        'has_differences': bool(mismatched or missing_in_image or missing_in_excel),
        'image_totals': image_totals,
    }
