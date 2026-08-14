from django.urls import path

from . import views

app_name = 'reconciliation'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('pharmacies/add/', views.add_pharmacy, name='add_pharmacy'),
    path('pharmacies/<int:pharmacy_id>/upload-image/<str:date_str>/', views.upload_image, name='upload_image'),
    path('pharmacies/<int:pharmacy_id>/upload-excel/<str:date_str>/', views.upload_excel, name='upload_excel'),
    path('reconciliations/<int:reconciliation_id>/compare/', views.run_comparison, name='run_comparison'),
    path('download-combined/', views.download_combined_excel, name='download_combined_excel'),
]
