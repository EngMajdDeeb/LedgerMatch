"""
Step 2 of the pipeline: extract the pharmacy's photographed cash-box table
into structured JSON via OpenAI vision calls.

The image typically shows a "صندوق" (cash-box) reconciliation sheet split
into one or two shift columns (e.g. الكادر الصباحي / الكادر المسائي), each
with the same set of flow rows (sales, returns, purchases, expenses...),
plus four day-level balance rows (opening/closing balance, the pharmacy's
transcription of the El-Bayan balance, and its self-computed discrepancy)
that are sometimes printed outside any shift's column.

Asking a single vision call to extract both the per-shift flow numbers *and*
the day-level balances in one shot proved unreliable in testing: whenever
the prompt/schema mentioned the day-level fields, the model would
consistently return null for every per-shift flow value, even though the
same per-shift-only prompt/schema extracted them correctly on its own. So
this module makes small, single-purpose vision calls instead of one complex
one, then merges the results into the shape aggregation.py expects.

A third call reads the sheet's free-text "الملاحظات" (notes) column, which
is per-row (one note per category, not per-shift) - the pharmacy sometimes
writes an explanatory remark next to a specific line (e.g. a delayed
shipment note next to مشتريات آجل). That note is carried through to the
comparison result as-is (see comparison_service._normalize_categories) so
it appears next to that category in the results table, without the AI
comparison step having to reason about it at all.
"""
from __future__ import annotations

import base64
import io
import json
from concurrent.futures import ThreadPoolExecutor

from django.conf import settings
from openai import APIError, APITimeoutError, OpenAIError
from PIL import Image

from ..constants import ALL_CATEGORY_KEYS, BALANCE_CATEGORY_KEYS, FLOW_CATEGORY_KEYS
from .openai_client import AIServiceError, friendly_api_error, get_client

MAX_DIMENSION = 2000
JPEG_QUALITY = 85

FLOW_SYSTEM_PROMPT = """أنت مساعد متخصص باستخراج بيانات جداول "صندوق" الصيدليات من صورة فوتوغرافية.
سيتم إرسال صورة لجدول تسوية صندوق يومي لصيدلية، مقسّم عادة إلى عمود أو عمودين (وردية صباحية ووردية مسائية).
اقرأ الصورة بدقة واستخرج قيم البنود التالية فقط لكل وردية على حدة كما تظهر تحت عمود تلك الوردية تحديدا:
مبيعات، مرتجع مبيعات، مشتريات نقدا، مشتريات آجل، مرتجع مشتريات، دفعات موردين، تمويل شركة، مصاريف1، مصاريف2، مصاريف3.
- إن كانت الخانة فارغة أو تحتوي على "-" فأعد القيمة null.
- لا تُقرّب أو تُعدّل الأرقام، انسخها كما هي.
- إن وُجدت وردية واحدة فقط أعد عنصرًا واحدًا في shifts.
- أعد النتيجة وفق المخطط (schema) المطلوب حصرًا بدون أي نص إضافي."""

DAY_TOTALS_SYSTEM_PROMPT = """أنت مساعد متخصص باستخراج بيانات جداول "صندوق" الصيدليات من صورة فوتوغرافية.
سيتم إرسال صورة لجدول تسوية صندوق يومي لصيدلية. استخرج فقط المعلومات التالية على مستوى اليوم كاملا
(رقم واحد لكل بند، بغض النظر عن العمود أو التظليل اللوني الذي يظهر تحته في الصورة):
- اسم الصيدلية وتاريخ التقرير كما يظهران في أعلى الجدول.
- صندوق بداية (opening_balance)
- صندوق نهاية (closing_balance)
- صندوق البيان (bayan_balance)
- فروقات الصندوق (discrepancy)
- أي ملاحظات نصية مكتوبة في الجدول (notes)

تنبيه مهم: "صندوق البيان" و"فروقات الصندوق" صفّان منفصلان ومتجاوران في الجدول، ومن السهل الخلط بينهما.
انظر بدقة إلى اسم البند المكتوب في نفس السطر الأفقي للرقم قبل أن تُقرر إلى أي حقل ينتمي ذلك الرقم.
إن كان أحد الصفين فارغا (أو يحتوي "-") والآخر فيه رقم، فأعد القيمة الفارغة null لصفها بالتحديد ولا تنسخ
رقم الصف الآخر إليها.
- إن كانت الخانة فارغة أو تحتوي على "-" فأعد القيمة null.
- لا تُقرّب أو تُعدّل الأرقام، انسخها كما هي.
- أعد النتيجة وفق المخطط (schema) المطلوب حصرًا بدون أي نص إضافي."""

ROW_NOTES_SYSTEM_PROMPT = """أنت مساعد متخصص باستخراج بيانات جداول "صندوق" الصيدليات من صورة فوتوغرافية.
الجدول في الصورة يحتوي عمود "الملاحظات" (وقد يكون بلا عنوان واضح، لكنه عادة العمود الأخير أو منفصل عن أعمدة الورديات).
لكل بند من البنود التالية، انظر إلى الصف الأفقي الخاص به وانسخ ما هو مكتوب في عمود الملاحظات في نفس ذلك الصف تحديدا:
مبيعات، مرتجع مبيعات، مشتريات نقدا، مشتريات آجل، مرتجع مشتريات، دفعات موردين، تمويل شركة، مصاريف1، مصاريف2، مصاريف3،
صندوق بداية، صندوق نهاية، صندوق البيان، فروقات الصندوق.
- انسخ النص كما هو حرفيا، لا تُلخّص ولا تُترجم ولا تُعدّل.
- إن كان صف بند ما لا يحتوي على أي ملاحظة (فارغ أو "-") فأعد null لذلك البند تحديدا.
- لا تخلط ملاحظة بند مع بند آخر مجاور له - كل ملاحظة تخص صفها فقط.
- لا تخترع ملاحظات غير مكتوبة فعليا في الصورة.
- أعد النتيجة وفق المخطط (schema) المطلوب حصرًا بدون أي نص إضافي."""


def _build_flow_values_schema() -> dict:
    properties = {key: {'type': ['number', 'null']} for key in FLOW_CATEGORY_KEYS}
    return {
        'type': 'object',
        'properties': properties,
        'required': FLOW_CATEGORY_KEYS,
        'additionalProperties': False,
    }


def _build_flow_schema() -> dict:
    return {
        'type': 'object',
        'properties': {
            'shifts': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        'shift_label': {'type': ['string', 'null']},
                        'hours': {'type': ['string', 'null']},
                        'values': _build_flow_values_schema(),
                    },
                    'required': ['shift_label', 'hours', 'values'],
                    'additionalProperties': False,
                },
            },
        },
        'required': ['shifts'],
        'additionalProperties': False,
    }


def _build_day_totals_schema() -> dict:
    balance_properties = {key: {'type': ['number', 'null']} for key in BALANCE_CATEGORY_KEYS}
    return {
        'type': 'object',
        'properties': {
            'pharmacy_name': {'type': ['string', 'null']},
            'report_date': {'type': ['string', 'null']},
            'day_totals': {
                'type': 'object',
                'properties': balance_properties,
                'required': BALANCE_CATEGORY_KEYS,
                'additionalProperties': False,
            },
            'notes': {'type': ['string', 'null']},
        },
        'required': ['pharmacy_name', 'report_date', 'day_totals', 'notes'],
        'additionalProperties': False,
    }


def _build_row_notes_schema() -> dict:
    properties = {key: {'type': ['string', 'null']} for key in ALL_CATEGORY_KEYS}
    return {
        'type': 'object',
        'properties': {
            'row_notes': {
                'type': 'object',
                'properties': properties,
                'required': ALL_CATEGORY_KEYS,
                'additionalProperties': False,
            },
        },
        'required': ['row_notes'],
        'additionalProperties': False,
    }


def _prepare_image_data_url(image_file) -> str:
    """Downscale/compress the image and return it as a base64 data URL."""
    image_file.seek(0)
    with Image.open(image_file) as img:
        img = img.convert('RGB')
        if max(img.size) > MAX_DIMENSION:
            img.thumbnail((MAX_DIMENSION, MAX_DIMENSION), Image.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=JPEG_QUALITY)
        encoded = base64.b64encode(buffer.getvalue()).decode('utf-8')
    return f'data:image/jpeg;base64,{encoded}'


def _call_vision(client, data_url, system_prompt, user_prompt, schema_name, schema):
    try:
        response = client.responses.create(
            model=settings.OPENAI_VISION_MODEL,
            input=[
                {'role': 'system', 'content': [{'type': 'input_text', 'text': system_prompt}]},
                {
                    'role': 'user',
                    'content': [
                        {'type': 'input_text', 'text': user_prompt},
                        {'type': 'input_image', 'image_url': data_url},
                    ],
                },
            ],
            text={
                'format': {
                    'type': 'json_schema',
                    'name': schema_name,
                    'schema': schema,
                    'strict': True,
                },
            },
        )
    except APITimeoutError as exc:
        raise AIServiceError('انتهت مهلة الاتصال بخدمة OpenAI أثناء تحليل الصورة.') from exc
    except APIError as exc:
        raise AIServiceError(friendly_api_error(exc, 'تحليل الصورة')) from exc
    except OpenAIError as exc:
        raise AIServiceError(f'تعذر الاتصال بخدمة OpenAI: {exc}') from exc

    output_text = getattr(response, 'output_text', None)
    if not output_text:
        raise AIServiceError('لم يُرجع نموذج الذكاء الاصطناعي أي بيانات من الصورة.')

    try:
        return json.loads(output_text)
    except json.JSONDecodeError as exc:
        raise AIServiceError('تعذر تفسير البيانات المستخرجة من الصورة (JSON غير صالح).') from exc


def extract_image_table(image_file) -> dict:
    """Send the pharmacy's sales-image to OpenAI vision and return structured JSON.

    `image_file` is a Django file-like object (opened in binary mode).
    Raises AIServiceError on any failure. The three extraction calls are
    independent, so they run concurrently to keep the pipeline's total
    latency down (each call can take up to a minute or more on its own).
    """
    data_url = _prepare_image_data_url(image_file)
    client = get_client()

    with ThreadPoolExecutor(max_workers=3) as executor:
        flow_future = executor.submit(
            _call_vision, client, data_url, FLOW_SYSTEM_PROMPT,
            'استخرج قيم بنود الحركة (المبيعات والمرتجعات والمشتريات والمصاريف) لكل وردية من هذه الصورة.',
            'cashbox_flow_extraction', _build_flow_schema(),
        )
        day_future = executor.submit(
            _call_vision, client, data_url, DAY_TOTALS_SYSTEM_PROMPT,
            'استخرج اسم الصيدلية والتاريخ وأرصدة الصندوق على مستوى اليوم من هذه الصورة.',
            'cashbox_day_totals_extraction', _build_day_totals_schema(),
        )
        notes_future = executor.submit(
            _call_vision, client, data_url, ROW_NOTES_SYSTEM_PROMPT,
            'استخرج محتوى عمود الملاحظات لكل بند من هذه الصورة.',
            'cashbox_row_notes_extraction', _build_row_notes_schema(),
        )
        flow_result = flow_future.result()
        day_result = day_future.result()
        notes_result = notes_future.result()

    return {
        'pharmacy_name': day_result.get('pharmacy_name'),
        'report_date': day_result.get('report_date'),
        'day_totals': day_result.get('day_totals', {}),
        'shifts': flow_result.get('shifts', []),
        'notes': day_result.get('notes'),
        'row_notes': notes_result.get('row_notes', {}),
    }
