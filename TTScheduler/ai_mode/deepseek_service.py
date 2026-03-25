# deepseek_service.py
"""
DeepSeek V3.1 via NVIDIA API - Fixed Version
"""

from openai import OpenAI
import json
from typing import Dict, Any

client = OpenAI(
    base_url="https://integrate.api.nvidia.com/v1",
    api_key="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"
)

MODEL = "deepseek-ai/deepseek-v3.2"


def generate_timetable_from_prompt(prompt: str) -> Dict[str, Any]:
    """
    Fixed version with better stream handling
    """
    try:
        completion = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "You are a strict timetable generator. Return ONLY valid JSON. No extra text."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            top_p=0.7,
            max_tokens=8192,
            extra_body={"chat_template_kwargs": {"thinking": True}},
            stream=True
        )

        full_response = ""
        reasoning = ""

        print("\n🤖 DeepSeek is thinking...\n")

        for chunk in completion:
            if not chunk.choices:
                continue

            delta = chunk.choices[0].delta

            # Capture reasoning (thinking)
            if getattr(delta, "reasoning_content", None):
                reasoning += delta.reasoning_content
                print("🧠", delta.reasoning_content, end="", flush=True)

            # Capture main content (this is what we need)
            if getattr(delta, "content", None) is not None:
                full_response += delta.content
                print(delta.content, end="", flush=True)

        print("\n\n✅ DeepSeek finished\n")

        # Final check
        if not full_response.strip():
            print("⚠️ WARNING: DeepSeek returned empty response!")
            return {
                "status": "error",
                "message": "DeepSeek returned empty response. Try again."
            }

        return {
            "status": "success",
            "data": full_response.strip()
        }

    except Exception as e:
        print(f"❌ API Error: {e}")
        return {
            "status": "error",
            "message": f"DeepSeek API Error: {str(e)}"
        }


