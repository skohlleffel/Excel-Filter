from django.urls import path, include
from . import views
from django.conf import settings

urlpatterns = [
    path('files/', views.file_list, name='file_list'),
    path('', views.upload_file, name='upload_file'),
]