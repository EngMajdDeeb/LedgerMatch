# تسوية صناديق الصيدليات (Dawak Finance)

أداة داخلية بلغة Django لمطابقة صندوق كل صيدلية آخر اليوم: صورة الجدول التي ترسلها الصيدلية
مقابل ملف الإكسل المُصدَّر من برنامج المحاسبة "البيان". المطابقة نفسها تتم عبر نموذج ذكاء اصطناعي
(OpenAI) بعد تحضير كل مصدر إلى بيانات منظّمة (JSON).

## المتطلبات

- Python 3.11+ (تم التطوير والاختبار على Python 3.14)
- مفتاح OpenAI API صالح مع صلاحية الوصول لنموذج يدعم الرؤية (vision) والإخراج المنظم (structured outputs)

## الإعداد المحلي

```powershell
# 1) إنشاء بيئة افتراضية وتثبيت الحزم
python -m venv venv
.\venv\Scripts\pip install -r requirements.txt

# 2) إعداد متغيرات البيئة
copy .env.example .env
# ثم عدّل .env وأضف OPENAI_API_KEY (وأي إعدادات أخرى) - هذا الملف لا يُرفع لنظام التحكم بالإصدار (git-ignored)

# 3) تطبيق الترحيلات (migrations) - سيُنشئ قاعدة بيانات SQLite ويزرع صيدلية تجريبية واحدة
.\venv\Scripts\python manage.py migrate

# 4) (اختياري) إنشاء مستخدم مدير للوصول إلى /admin
.\venv\Scripts\python manage.py createsuperuser

# 5) تجميع الملفات الثابتة (مطلوب لأن WhiteNoise يخدمها حتى محليا)
.\venv\Scripts\python manage.py collectstatic --noinput

# 6) تشغيل الخادم المحلي
.\venv\Scripts\python manage.py runserver
```

افتح المتصفح على `http://127.0.0.1:8000/`.

## متغيرات البيئة (`.env`)

| المتغير | الوصف |
|---|---|
| `DJANGO_SECRET_KEY` | مفتاح Django السري (غيّره في الإنتاج) |
| `DJANGO_DEBUG` | `True` للتطوير المحلي |
| `DJANGO_ALLOWED_HOSTS` | قائمة مفصولة بفواصل |
| `OPENAI_API_KEY` | مفتاح OpenAI - **إلزامي** لتشغيل عمليات المطابقة |
| `OPENAI_VISION_MODEL` | اسم النموذج المستخدم لاستخراج بيانات الصورة (رؤية) |
| `OPENAI_TEXT_MODEL` | اسم النموذج المستخدم لخطوة المطابقة النصية |
| `OPENAI_TIMEOUT_SECONDS` | مهلة الاتصال بـ OpenAI (بالثواني) |

## كيف تعمل الأداة

1. لكل صيدلية، تُرفع **صورة أو ملف إكسل** لصندوق آخر اليوم (خانة "تقرير صندوق الصيدلية") بالإضافة إلى
   ملف إكسل "البيان" عبر HTMX (بدون إعادة تحميل الصفحة).
2. عند الضغط على "قارن هذه الصيدلية" أو "قارن كل الصيدليات":
   - `reconciliation/services/excel_service.py` يقرأ ملف البيان (كشف حساب/قيود محاسبية) ويحوّله إلى
     جدول منظّم (تنظيف بيانات فقط، بدون أي حكم أو مطابقة).
   - تقرير صندوق الصيدلية يُستخرج بإحدى طريقتين حسب نوع الملف المرفوع (`views._extract_pharmacy_submission`):
     - **صورة**: `reconciliation/services/vision_service.py` يستخرج البيانات عبر نموذج رؤية من OpenAI.
     - **ملف إكسل**: `reconciliation/services/pharmacy_excel_service.py` يقرأ نفس التخطيط مباشرة من
       الخلايا (بدون أي ذكاء اصطناعي) - أدق من قراءة الصورة عند توفره.
     كلا الطريقتين تُرجعان نفس البنية (JSON) بنفس التصنيفات (مبيعات، مرتجعات، مشتريات، مصاريف، أرصدة الصندوق).
   - `reconciliation/services/comparison_service.py` يرسل الجدولين إلى نموذج نصي من OpenAI ليقوم هو
     بتصنيف حركات كشف الحساب ومطابقتها مع بيانات الصورة، ويعيد نتيجة مُهيكلة (JSON) تتضمن كل بند:
     قيمته حسب البيان، قيمته حسب الصورة، الفرق، الحالة، وملخص عربي قصير.
   - `reconciliation/services/excel_export_service.py` يبني ملف إكسل حقيقي من بيانات الصورة المستخرجة
     لمراجعته يدويا جنب ملف البيان الأصلي (قابل للتنزيل من بطاقة كل صيدلية).
3. تُخزَّن كل الملفات المرفوعة والنتائج في قاعدة البيانات ومجلد `media/` كسجل تاريخي دائم.

## التشغيل عبر Docker (للنشر خلف Cloudflare Tunnel)

```powershell
# 1) تأكد من وجود .env (بنفس المتغيرات أعلاه) في مجلد المشروع
# 2) تأكد أن DJANGO_DEBUG=False و DJANGO_ALLOWED_HOSTS يتضمن نطاقك (أو * مؤقتا)
#    وأضف DJANGO_CSRF_TRUSTED_ORIGINS=https://your-domain.example (بدون هذا سترفض النماذج POST عبر النفق)

docker compose up -d --build
```

التطبيق سيعمل على **`http://localhost:8000`**. شغّل نفق Cloudflare مقابل هذا المنفذ:

```powershell
cloudflared tunnel --url http://localhost:8000
```

(أو نفق دائم مربوط بنطاقك عبر `cloudflared tunnel run` بعد `cloudflared tunnel login` و`cloudflared tunnel create`).

ملاحظات مهمة:
- `docker-compose.yml` يربط `db.sqlite3` و `media/` مباشرة بمجلد المشروع على القرص (bind mount)، فالبيانات تبقى محفوظة حتى بعد `docker compose down` أو إعادة البناء.
- الحاوية تُعاد تشغيلها تلقائيا عند إعادة تشغيل الجهاز (`restart: unless-stopped`) طالما Docker Desktop يعمل.
- **لا يوجد تسجيل دخول حاليا** - أي شخص يملك رابط النفق يمكنه رفع/مقارنة بيانات الصيدليات. هذا خيار مقصود بناء على الطلب، لكن يجب عدم مشاركة الرابط علنا.
- لمشاهدة السجلات: `docker compose logs -f web`. لإيقاف التطبيق: `docker compose down` (لا يحذف البيانات).

## اختبار الخدمات مباشرة (Django shell)

```powershell
.\venv\Scripts\python manage.py shell
```

```python
from reconciliation.services import excel_service, vision_service, comparison_service

excel_table = excel_service.parse_el_bayan_excel('path/to/بيان.xlsx')
with open('path/to/image.png', 'rb') as f:
    image_extraction = vision_service.extract_image_table(f)
result = comparison_service.run_comparison(excel_table, image_extraction)
print(result['summary_ar'])
```

## هيكل المشروع

```
dawak_finance/          إعدادات المشروع (settings, urls)
reconciliation/         التطبيق الرئيسي
  models.py              Pharmacy / DailyReconciliation / ComparisonResult
  constants.py            تعريف فئات صندوق اليوم (مبيعات، مرتجعات، مشتريات...)
  validators.py            فحص نوع/حجم الملفات المرفوعة
  services/
    excel_service.py          تحضير جدول البيان (pandas/openpyxl)
    vision_service.py         استخراج بيانات صورة الصندوق (OpenAI - رؤية)
    pharmacy_excel_service.py استخراج بيانات صندوق الصيدلية من ملف إكسل مباشرة (بدون AI)
    aggregation.py            جمع حسابي بسيط (غير ذكاء اصطناعي) لتجميع نتائج تقرير الصندوق
    comparison_service.py     محرك المطابقة (OpenAI - نصي)
    excel_export_service.py   بناء ملف إكسل تدقيقي من بيانات تقرير الصندوق
  views.py / urls.py         نقاط HTMX
  templates/reconciliation/  الواجهة (عربي/RTL بالكامل)
```
