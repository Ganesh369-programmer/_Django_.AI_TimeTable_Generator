from django.shortcuts import render, redirect
from django.contrib import messages
from django.http import HttpResponse
import pandas as pd
import json
import traceback


from django.http import JsonResponse, HttpResponse
from django.template.loader import render_to_string
# ===============================
# MAIN PAGE
# ===============================
def generate_page(request):
    return render(request, 'ai_mode/generate.html')


def result_page(request):
    return render(request, 'ai_mode/result.html')



# ===============================
# AJAX SUBJECT FORM LOADER (optional)
# ===============================
def get_subject_form(request):
    if request.method == "GET":
        subject = request.GET.get("subject", "")
        if not subject:
            return JsonResponse({"error": "No subject"}, status=400)

        html = render_to_string(
            "ai_mode/form_template.html",
            {"subject": subject},
            request=request
        )
        return JsonResponse({"html": html})

    return JsonResponse({"error": "Method not allowed"}, status=405)





# timetable form submit
def timetable_form(request):

    if request.method=="POST":

        dept_names=request.POST.getlist("dept_name[]")
        semesters=request.POST.getlist("semester[]")
        working_days=request.POST.getlist("working_days[]")
        lecture_duration=request.POST.getlist("lecture_duration[]")
        lab_duration=request.POST.getlist("lab_duration[]")
        break_name=request.POST.getlist("break_name[]")
        classrooms=request.POST.getlist("classrooms[]")
        labs=request.POST.getlist("labs[]")

        subject_files=request.FILES.getlist("subject_excel[]")
        teacher_files=request.FILES.getlist("teacher_excel[]")

        total=len(dept_names)

        print("\n=========== MULTI AI PROMPT ===========")

        for i in range(total):

            # SUBJECT JSON
            subject_json=[]
            if i < len(subject_files):
                df=pd.read_excel(subject_files[i])
                subject_json=df.to_dict(orient="records")

            # TEACHER JSON
            teacher_json=[]
            if i < len(teacher_files):
                df2=pd.read_excel(teacher_files[i])
                teacher_json=df2.to_dict(orient="records")

            ai_prompt=f"""
==============================
DEPARTMENT {i+1}
==============================
Department: {dept_names[i]}
Semester: {semesters[i]}

Working Days: {working_days[i]}
Lecture Duration: {lecture_duration[i]}
Lab Duration: {lab_duration[i]}
Break: {break_name[i]}

Classrooms: {classrooms[i]}
Labs: {labs[i]}

SUBJECT JSON:
{json.dumps(subject_json,indent=2)}

TEACHER JSON:
{json.dumps(teacher_json,indent=2)}

Generate weekly timetable JSON.
"""

            print(ai_prompt)

        print("\n=======================================\n")

        return HttpResponse("✅ Multiple AI prompts generated. Check terminal.")

    return HttpResponse("Invalid")