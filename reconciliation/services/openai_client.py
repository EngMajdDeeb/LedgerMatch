"""Shared OpenAI client factory + common error types for the AI services."""
from django.conf import settings
from openai import APIError, OpenAI


class AIServiceError(Exception):
    """Raised whenever an OpenAI call fails or returns something unusable."""


_client = None


def get_client() -> OpenAI:
    global _client
    if not settings.OPENAI_API_KEY:
        raise AIServiceError(
            'مفتاح OPENAI_API_KEY غير مُعرَّف. الرجاء إضافته إلى ملف .env'
        )
    if _client is None:
        _client = OpenAI(api_key=settings.OPENAI_API_KEY, timeout=settings.OPENAI_TIMEOUT_SECONDS)
    return _client


def friendly_api_error(exc: APIError, action_ar: str) -> str:
    """Turn an OpenAI APIError into a clear Arabic message instead of the
    raw JSON error body, recognising the failure modes an operator would
    actually need to act on (out of credits, bad key, rate limit).
    """
    status = getattr(exc, 'status_code', None)
    body = getattr(exc, 'body', None) or {}
    code = (body.get('code') if isinstance(body, dict) else None) or getattr(exc, 'code', None) or ''

    if status == 429 and code in ('insufficient_quota', 'credit_balance_exhausted'):
        return (
            'نفد رصيد حساب OpenAI، لذلك تعذّرت ' + action_ar + '. '
            'الرجاء إضافة رصيد من https://platform.openai.com/settings/organization/billing/ '
            'ثم إعادة المحاولة لهذه الصيدلية.'
        )
    if status == 429:
        return f'تم تجاوز الحد المسموح من الطلبات لخدمة OpenAI أثناء {action_ar}. الرجاء الانتظار قليلا ثم إعادة المحاولة.'
    if status == 401:
        return 'مفتاح OpenAI API غير صالح أو منتهي الصلاحية. تحقق من OPENAI_API_KEY في ملف .env.'
    return f'خطأ من خدمة OpenAI أثناء {action_ar}: {exc}'
