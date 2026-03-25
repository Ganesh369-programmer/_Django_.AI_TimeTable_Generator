from django.urls import path
from . import views

urlpatterns = [
    path('generate/', views.generate_page, name='generate'),
    path('result/', views.result_page, name='result'),
    path('get-subject-form/', views.get_subject_form, name='get_subject_form'),
    path("timetable/", views.timetable_form, name="timetable"),
    #don't Give the similar name 
    path("convert-json/" , views.json_2_excel_converter , name="json_to_excel"),
    path('download-excel/', views.download_excel, name='download_excel'),
    path('preview-excel/', views.preview_excel, name='preview_excel'), 
]