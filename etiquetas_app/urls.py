from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('procesar/', views.procesar_archivos, name='procesar_archivos'),
    path('descargar/', views.descargar, name='descargar'),
    path('obtener-documento/', views.obtener_documento, name='obtener_documento'),
]