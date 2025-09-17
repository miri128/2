from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_file, name='upload'),
    path('summary/', views.summary_view, name='summary'),
]