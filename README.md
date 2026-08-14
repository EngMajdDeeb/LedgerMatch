# LedgerMatch 🧾

> An internal Django tool that reconciles daily pharmacy cash registers using AI-powered document understanding — matching pharmacy-submitted cash reports against accounting system exports automatically.

---

## Overview

LedgerMatch automates end-of-day cash reconciliation for pharmacy networks. Each pharmacy submits either a **photo** or an **Excel file** of their daily cash drawer. The tool compares it against the exported ledger from the **Al-Bayan** accounting system and produces a structured, itemized discrepancy report — powered by OpenAI vision and text models.

**The problem it solves:** Manual reconciliation across multiple pharmacy branches is slow, error-prone, and hard to audit. LedgerMatch reduces this to a single click.

---

## How It Works

```
Pharmacy Submission          Al-Bayan Export
(Image or Excel)             (Excel file)
       │                           │
       ▼                           ▼
 Vision Service /          Excel Service
 Excel Parser              (pandas/openpyxl)
       │                           │
       └──────────┬────────────────┘
                  ▼
         Comparison Service
         (OpenAI text model)
                  │
                  ▼
     Structured JSON Result
     (per-line: value, diff, status, summary)
                  │
                  ▼
     HTMX UI Update + Excel Export
```

1. **Upload** — The pharmacy submits their daily cash report (image or Excel) alongside the Al-Bayan export, via HTMX (no page reload).
2. **Extract** — If an image is submitted, the OpenAI vision model extracts structured data. If an Excel file is submitted, a dedicated parser reads the cells directly (faster and more accurate).
3. **Normalize** — Both sources are converted to the same JSON schema: sales, returns, purchases, expenses, and cash balances.
4. **Compare** — The OpenAI text model receives both structured tables, classifies each ledger line, matches it against the pharmacy report, and returns a full reconciliation result with per-item status and an Arabic summary.
5. **Export** — An auditable Excel file is generated from the extracted pharmacy data, downloadable from each pharmacy's card for side-by-side manual review.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Django 5.x, Python 3.11+ |
| AI / Vision | OpenAI API (vision + structured outputs) |
| Data Processing | pandas, openpyxl |
| Frontend | HTMX, Django Templates (RTL/Arabic) |
| Database | SQLite (with Docker bind mount for persistence) |
| Static Files | WhiteNoise |
| Deployment | Docker Compose + Cloudflare Tunnel |

---

## Requirements

- Python 3.11+
- A valid OpenAI API key with access to a vision-capable model and structured outputs support

---

## Local Setup

```powershell
# 1. Create virtual environment and install dependencies
python -m venv venv
.\venv\Scripts\pip install -r requirements.txt

# 2. Configure environment variables
copy .env.example .env
# Edit .env and set OPENAI_API_KEY (and other settings) — this file is git-ignored

# 3. Apply migrations (creates SQLite DB and seeds one demo pharmacy)
.\venv\Scripts\python manage.py migrate

# 4. (Optional) Create a superuser for /admin access
.\venv\Scripts\python manage.py createsuperuser

# 5. Collect static files (required — WhiteNoise serves them even locally)
.\venv\Scripts\python manage.py collectstatic --noinput

# 6. Start the development server
.\venv\Scripts\python manage.py runserver
```

Open your browser at `http://127.0.0.1:8000/`.

---

## Environment Variables

| Variable | Description |
|---|---|
| `DJANGO_SECRET_KEY` | Django secret key — change in production |
| `DJANGO_DEBUG` | Set to `True` for local development |
| `DJANGO_ALLOWED_HOSTS` | Comma-separated list of allowed hosts |
| `OPENAI_API_KEY` | **Required** — used for all AI operations |
| `OPENAI_VISION_MODEL` | Model name for image data extraction (vision) |
| `OPENAI_TEXT_MODEL` | Model name for the text-based comparison step |
| `OPENAI_TIMEOUT_SECONDS` | Timeout for OpenAI API calls (seconds) |

---

## Docker Deployment (with Cloudflare Tunnel)

```powershell
# 1. Ensure .env is present with the variables above
# 2. Set DJANGO_DEBUG=False and DJANGO_ALLOWED_HOSTS to your domain (or * temporarily)
#    Add DJANGO_CSRF_TRUSTED_ORIGINS=https://your-domain.example
#    (Without this, POST requests through the tunnel will be rejected)

docker compose up -d --build
```

The app runs on `http://localhost:8000`. Point a Cloudflare Tunnel at that port:

```powershell
cloudflared tunnel --url http://localhost:8000
```

**Important notes:**
- `docker-compose.yml` uses bind mounts for `db.sqlite3` and `media/` — data persists across rebuilds and `docker compose down`.
- The container restarts automatically on machine reboot (`restart: unless-stopped`) as long as Docker Desktop is running.
- **No authentication is implemented** — anyone with the tunnel URL can upload and compare pharmacy data. Do not share the URL publicly.
- View logs: `docker compose logs -f web` | Stop: `docker compose down` (data is preserved).

---

## Testing Services via Django Shell

```powershell
.\venv\Scripts\python manage.py shell
```

```python
from reconciliation.services import excel_service, vision_service, comparison_service

excel_table = excel_service.parse_el_bayan_excel('path/to/bayan.xlsx')

with open('path/to/image.png', 'rb') as f:
    image_extraction = vision_service.extract_image_table(f)

result = comparison_service.run_comparison(excel_table, image_extraction)
print(result['summary_ar'])
```

---

## Project Structure

```
ledgermatch/
├── dawak_finance/              # Project settings, URLs, WSGI
├── reconciliation/             # Main application
│   ├── models.py               # Pharmacy / DailyReconciliation / ComparisonResult
│   ├── constants.py            # Cash category definitions (sales, returns, purchases…)
│   ├── validators.py           # File type and size validation
│   ├── views.py                # HTMX endpoints
│   ├── urls.py
│   ├── services/
│   │   ├── excel_service.py            # Al-Bayan Excel parser (pandas/openpyxl)
│   │   ├── vision_service.py           # Image extraction via OpenAI vision
│   │   ├── pharmacy_excel_service.py   # Pharmacy Excel parser (no AI)
│   │   ├── aggregation.py              # Simple arithmetic aggregation
│   │   ├── comparison_service.py       # AI-powered reconciliation engine
│   │   └── excel_export_service.py     # Audit Excel file builder
│   └── templates/reconciliation/       # Arabic RTL UI templates
├── media/                      # Uploaded files and results (git-ignored)
├── .env.example
├── requirements.txt
├── Dockerfile
└── docker-compose.yml
```

---

## Key Design Decisions

- **Dual-input support:** Both image and Excel submissions produce the same normalized JSON structure, keeping the comparison logic format-agnostic.
- **AI for comparison, not just extraction:** The comparison step is deliberately AI-driven to handle the fuzzy mapping between Al-Bayan's accounting categories and pharmacy cash report line items.
- **HTMX over SPA:** Full-page-reload-free UX without the complexity of a separate frontend framework.
- **SQLite + bind mounts:** Keeps deployment simple for an internal single-server tool without sacrificing data persistence.

---

## License

Internal use only. Not licensed for public distribution.
