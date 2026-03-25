import google.generativeai as genai
import json

# 🔑 Replace with your Gemini API key
GEMINI_API_KEY = "AIzaSyAdSd1BOCoLpYj8wToLywG5_eeCsjNobvQ"

genai.configure(api_key=GEMINI_API_KEY)

model = genai.GenerativeModel("gemini-3.1-flash-lite-preview")
#gemini-3-flash-preview
#gemini-3.1-flash-lite-preview

def generate_timetable_from_prompt(prompt_text):
    """
    Send prompt to Gemini and return JSON response
    """

    try:
        response = model.generate_content(prompt_text)

        raw_text = response.text.strip()

        # Sometimes Gemini wraps JSON in ```json ``` block
        if raw_text.startswith("```"):
            raw_text = raw_text.replace("```json", "").replace("```", "").strip()

        parsed_json = json.loads(raw_text)

        return {
            "status": "success",
            "data": parsed_json
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }