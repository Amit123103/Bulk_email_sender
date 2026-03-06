"""
AI Generator — Optional OpenAI integration for email generation.
Gracefully handles missing openai package.
"""

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

import os
import json


class AIEmailGenerator:
    """
    Generates professional email content using OpenAI GPT.
    Falls back gracefully if openai is not installed or API key is missing.
    """

    TONES = {
        "Professional": "Write in a professional, polished business tone.",
        "Friendly": "Write in a warm, friendly, and approachable tone.",
        "Formal": "Write in a very formal and respectful tone.",
        "Casual": "Write in a casual, conversational tone.",
        "Persuasive": "Write in a persuasive, compelling tone focused on benefits.",
        "Urgent": "Write with urgency, encouraging immediate action.",
    }

    DEFAULT_MODEL = "gpt-3.5-turbo"

    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY", "")
        self.available = OPENAI_AVAILABLE and bool(self.api_key)
        if self.available:
            self.client = openai.OpenAI(api_key=self.api_key)
        else:
            self.client = None

    def is_available(self) -> bool:
        return self.available

    def generate_email(self, description: str, tone: str = "Professional",
                       variables: list = None, is_html: bool = True) -> dict:
        """
        Generate email subject and body from a plain text description.

        Returns: {"success": bool, "subject": str, "body": str, "error": str|None}
        """
        if not self.available:
            missing = []
            if not OPENAI_AVAILABLE:
                missing.append("Install openai package: pip install openai")
            if not self.api_key:
                missing.append("Set OPENAI_API_KEY environment variable or pass api_key")
            return {
                "success": False,
                "subject": "",
                "body": "",
                "error": "AI generation unavailable. " + ". ".join(missing)
            }

        tone_instruction = self.TONES.get(tone, self.TONES["Professional"])
        format_instruction = (
            "Format the email body as clean HTML with proper tags (p, br, strong, etc.). "
            "Do NOT include <!DOCTYPE>, <html>, <head>, or <body> tags — just the inner content."
            if is_html else
            "Format the email as plain text with proper line breaks."
        )

        vars_instruction = ""
        if variables:
            vars_instruction = (
                f"\n\nAvailable personalization variables that you SHOULD use in the email: "
                f"{', '.join(variables)}. Use them naturally where appropriate."
            )

        prompt = (
            f"Generate a professional email based on this description:\n\n"
            f"\"{description}\"\n\n"
            f"Instructions:\n"
            f"- {tone_instruction}\n"
            f"- {format_instruction}\n"
            f"- Generate both a compelling subject line and the email body.\n"
            f"- Keep the subject under 60 characters.\n"
            f"- Make the email concise but effective (150-300 words)."
            f"{vars_instruction}\n\n"
            f"Respond ONLY with valid JSON in this exact format:\n"
            f'{{"subject": "Your subject here", "body": "Your email body here"}}'
        )

        try:
            response = self.client.chat.completions.create(
                model=self.DEFAULT_MODEL,
                messages=[
                    {"role": "system", "content": "You are an expert email copywriter. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=1000,
            )

            content = response.choices[0].message.content.strip()

            # Parse JSON response
            # Handle potential markdown code block wrapping
            if content.startswith("```"):
                content = content.split("\n", 1)[1]
                if content.endswith("```"):
                    content = content[:-3]
                content = content.strip()

            result = json.loads(content)
            return {
                "success": True,
                "subject": result.get("subject", ""),
                "body": result.get("body", ""),
                "error": None
            }

        except json.JSONDecodeError:
            # If JSON parsing fails, try to extract subject and body manually
            return {
                "success": True,
                "subject": "Your Email Subject",
                "body": content if content else "Could not parse AI response.",
                "error": "AI responded with non-JSON format. Body set to raw response."
            }
        except Exception as e:
            return {
                "success": False,
                "subject": "",
                "body": "",
                "error": f"AI generation failed: {str(e)}"
            }

    def get_available_tones(self) -> list:
        return list(self.TONES.keys())
