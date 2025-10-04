from django.urls import path
from . import views

urlpatterns = [
    path('arrayup', views.upload_file, name='arrayup'),
    path('summary/', views.summary_view, name='summary'),
    path('', views.TopView.as_view(), name="top"),

]