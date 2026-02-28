from django.urls import path
from . import views

urlpatterns = [
    path('generate/', views.generate_page, name='generate'),
    path('result/', views.result_page, name='result'),
    path('get-subject-form/', views.get_subject_form, name='get_subject_form'),
    path("timetable/", views.timetable_form, name="timetable"),
    # path("preview-excel/", views.preview_excel, name="preview_excel"),
]