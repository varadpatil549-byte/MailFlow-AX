import json
from ollama import chat


MODEL = "qwen3:4b-instruct"

CATEGORIES = [
    "Invoice",
    "Payment",
    "Vendor Query",
    "PO / Procurement",
    "GRIR",
    "Access / System",
    "Process Issue",
    "General",
    "Other",
]


def classify_email(subject: str, body: str) -> dict:
    """Classify an email using the local Ollama model."""

    prompt = f"""
You are an enterprise email classification engine.

Classify the email into EXACTLY ONE of these categories:

{", ".join(CATEGORIES)}

Rules:
1. category must be exactly one of the allowed categories.
2. confidence must be an integer from 0 to 100.
3. reason must be short.
4. Return ONLY JSON.
5. Never create a new category.

Email Subject:
{subject}

Email Body:
{body[:5000]}
"""

    response = chat(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": prompt,
            }
        ],
        format={
            "type": "object",
            "properties": {
                "category": {
                    "type": "string",
                    "enum": CATEGORIES,
                },
                "confidence": {
                    "type": "integer",
                    "minimum": 0,
                    "maximum": 100,
                },
                "reason": {
                    "type": "string",
                },
            },
            "required": [
                "category",
                "confidence",
                "reason",
            ],
        },
        options={
            "temperature": 0,
        },
    )

    result = json.loads(response.message.content)

    return result