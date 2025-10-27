from django.urls import path
from . import views

urlpatterns = [
    path('', views.document_list, name='document_list'),
    path('document/<int:pk>/', views.document_detail, name='document_detail'),
    path('document/<int:document_id>/status/', views.document_status, name='document_status'),
    path('process/', views.process_document_view, name='process_document'),
]