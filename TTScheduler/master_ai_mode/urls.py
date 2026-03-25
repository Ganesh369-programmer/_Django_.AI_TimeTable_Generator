from django.urls import path
from . import views

urlpatterns = [
    path('master_generate/', views.master_generate_page, name='master_generate'),
    path('master-timetable/', views.master_timetable_view, name='master_timetable'),
    path('master-result/',         views.master_result_view,    name='master_result'),
    path('master-json-to-excel/',  views.master_json_to_excel,  name='master_json_to_excel'),
    path('master-preview-excel/',  views.master_preview_excel,  name='master_preview_excel'),
]