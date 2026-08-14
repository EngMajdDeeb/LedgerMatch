from django.contrib import admin

from .models import ComparisonResult, DailyReconciliation, Pharmacy


@admin.register(Pharmacy)
class PharmacyAdmin(admin.ModelAdmin):
    list_display = ('name', 'created_at')
    search_fields = ('name',)


@admin.register(DailyReconciliation)
class DailyReconciliationAdmin(admin.ModelAdmin):
    list_display = ('pharmacy', 'date', 'status', 'created_at')
    list_filter = ('status', 'date', 'pharmacy')
    search_fields = ('pharmacy__name',)


@admin.register(ComparisonResult)
class ComparisonResultAdmin(admin.ModelAdmin):
    list_display = ('reconciliation', 'has_differences', 'matched_rows_count', 'mismatched_rows_count', 'created_at')
    list_filter = ('has_differences',)
