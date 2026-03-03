from django.shortcuts import render
from django.http import JsonResponse, HttpResponse
from django.template.loader import render_to_string
import pandas as pd
import json
import traceback


# ===============================
# MAIN PAGES
# ===============================

def generate_page(request):
    return render(request, 'ai_mode/generate.html')


def result_page(request):
    return render(request, 'ai_mode/result.html')


# ===============================
# AJAX SUBJECT FORM LOADER
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


# ===============================
# SAFE HELPER FUNCTION
# ===============================

def safe_get(lst, index):
    """Safely get value from list"""
    return lst[index] if index < len(lst) else ""


# ===============================
# TIMETABLE FORM SUBMIT
# ===============================

def timetable_form(request):

    if request.method == "POST":

        try:

            # ===============================
            # GET ALL LIST DATA
            # ===============================

            dept_names = request.POST.getlist("dept_name[]")
            semesters = request.POST.getlist("semester[]")
            working_days = request.POST.getlist("working_days[]")
            lecture_duration = request.POST.getlist("lecture_duration[]")
            lab_duration = request.POST.getlist("lab_duration[]")
            break_name = request.POST.getlist("break_name[]")
            classrooms = request.POST.getlist("classrooms[]")
            labs = request.POST.getlist("labs[]")

            subject_files = request.FILES.getlist("subject_excel[]")
            teacher_files = request.FILES.getlist("teacher_excel[]")

            start_time = request.POST.getlist("start_time[]")
            end_time = request.POST.getlist("end_time[]")
            break_duration = request.POST.getlist("break_duration[]")
            break_position = request.POST.getlist("break_position[]")

            total_classrooms = request.POST.getlist("total_classrooms[]")
            total_labs = request.POST.getlist("total_labs[]")
            shared_rooms = request.POST.getlist("shared_rooms[]")

            mini_project = request.POST.getlist("mini_project[]")
            
            total_students = request.POST.getlist("total_students[]")
            students_per_batch = request.POST.getlist("students_per_batch[]")
            total_batches = request.POST.getlist("total_batches[]")
            batch_names = request.POST.getlist("batch_names[]")
            lab_mode = request.POST.getlist("lab_allocation_mode[]")

            total = len(dept_names)

            print("\n=========== MULTI AI PROMPT ===========")
            print("Total Departments:", total)
            print("ACT AS AN EXPERT ACADEMIC SCHEDULER. Generate a valid JSON timetable for the following department.")
            all_prompts = []

            # ===============================
            # LOOP THROUGH DEPARTMENTS
            # ===============================

            for i in range(total):

                # -------------------------------
                # SUBJECT EXCEL TO JSON
                # -------------------------------
                subject_json = []
                if i < len(subject_files):
                    df = pd.read_excel(subject_files[i])
                    subject_json = df.to_dict(orient="records")

                # -------------------------------
                # TEACHER EXCEL TO JSON
                # -------------------------------
                teacher_json = []
                if i < len(teacher_files):
                    df2 = pd.read_excel(teacher_files[i])
                    teacher_json = df2.to_dict(orient="records")

                # -------------------------------
                # AI PROMPT
                # -------------------------------

                ai_prompt = f"""

                        ==============================
                        DEPARTMENT {i+1}
                        ==============================
                        Department: {safe_get(dept_names,i)}
                        Semester: {safe_get(semesters,i)}

                        Working Days: {safe_get(working_days,i)}
                        Lecture Duration: {safe_get(lecture_duration,i)} Hour
                        Lab Duration: {safe_get(lab_duration,i)} Hour
                        Break: {safe_get(break_name,i)} Minutes

                        Classrooms: {safe_get(classrooms,i)}
                        Labs: {safe_get(labs,i)}

                        SUBJECT JSON:
                        {json.dumps(subject_json, indent=2)}

                        TEACHER JSON:
                        {json.dumps(teacher_json, indent=2)}

                        
                        Academic Structure:
                        Start Time: {safe_get(start_time,i)}
                        End Time: {safe_get(end_time,i)}
                        Break Duration: {safe_get(break_duration,i)}
                        Break Position: {safe_get(break_position,i)}

                        Infrastructure:
                        Total Classrooms: {safe_get(total_classrooms,i)}
                        Total Labs: {safe_get(total_labs,i)}
                        Shared Rooms: {safe_get(shared_rooms,i)}

                        Special Blocks:
                        Mini Project: {safe_get(mini_project,i)}

                        Batch System:
                        Total Students: {safe_get(total_students,i)}
                        Students Per Batch: {safe_get(students_per_batch,i)}
                        Total Batches: {safe_get(total_batches,i)}
                        Batch Names: {safe_get(batch_names,i)}
                        Lab Mode: {safe_get(lab_mode,i)}


                        Generate weekly timetable JSON.`
                        """
                 

                print(ai_prompt)
                all_prompts.append(ai_prompt)

            print("\n=======================================\n")
            end = """### STRICT JSON OUTPUT FORMAT REQUIRED:
                Follow this structure EXACTLY. 
                - For Theory: {{"time": "9.30-10.30", "subject": "Name", "class": "ShortCode", "room": "No"}}
                - For Labs: {{"time": "11.30-1.30", "type": "Lab/Practical", "details": ["SUB-TEACHER-ROOM-BATCH"]}}
                - For Breaks: {{"time": "1.30-2.00", "type": "Break"}}

                Example Object:
                {{
                "timetable": [
                    {{
                    "day": "MON",
                    "schedule": [
                        {{"time": "9.30-10.30", "subject": "IoT", "class": "PSS", "room": "323"}},
                        {{"time": "11.30-1.30", "type": "Lab/Practical", "details": ["CSSL-SRC-325A-T1", "MCL-SVT-321-T2" , AIL-DES-324-T3]}}
                    ]
                    }}
                ]
                }}

                DO NOT ADD ANY PROSE OR EXPLANATION. ONLY RETURN THE JSON."""
            print(end)
            # ===============================
            # RETURN SUCCESS RESPONSE
            # ===============================

            return JsonResponse({
                "status": "success",
                "message": "Multiple AI prompts generated successfully.",
                "total_departments": total,
                "prompts": all_prompts
            })

        except Exception as e:
            print("ERROR OCCURRED:")
            print(traceback.format_exc())

            return JsonResponse({
                "status": "error",
                "message": str(e)
            })

    return HttpResponse("Invalid Request Method")