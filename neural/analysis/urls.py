from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('calculate/', views.calculate, name='calculate'),
    path('docxForms/', views.docxForms, name='docxForms'),
    path('test/', views.test, name='test'),
    path("create/", views.create),
    path("calculate/createRec/", views.createRec),
    path('docxForms/createDoc', views.create_all_word_document_view, name='create_all_word_document_view'),
]