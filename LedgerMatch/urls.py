"""
URL configuration for dawak_finance project.
"""
import re

from django.conf import settings
from django.contrib import admin
from django.urls import include, path, re_path
from django.views.static import serve as serve_static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('reconciliation.urls')),
]

# Served unconditionally (not just in DEBUG): this is a single-container
# deployment with no separate nginx/CDN in front of it, so Django itself has
# to serve uploaded images/Excel files and the generated audit workbooks.
# Fine at this app's traffic scale (internal tool, ~10 pharmacies/day).
#
# Deliberately NOT using django.conf.urls.static.static() here - that
# helper silently no-ops (registers nothing) whenever DEBUG is False, which
# is exactly the production case this deployment needs it for. Wiring the
# underlying view directly bypasses that DEBUG gate.
urlpatterns += [
    re_path(
        r'^%s(?P<path>.*)$' % re.escape(settings.MEDIA_URL.lstrip('/')),
        serve_static,
        {'document_root': settings.MEDIA_ROOT},
    ),
]
