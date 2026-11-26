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
    """
    Classify an email using the local Ollama model.
    """

    prompt = f"""
You are an enterprise email classification engine.

Classify the email into EXACTLY ONE of these categories:

{", ".join(CATEGORIES)}

Rules:
1. category must be exactly one of the allowed categories.
2. confidence must be an integer from 0 to 100.
3. reason must be a short explanation.
4. Return ONLY JSON.
5. Do not create a new category.

Email Subject:
{subject}

Email Body:
{body}
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

    import json

    result = json.loads(response.message.content)

    return result


if __name__ == "__main__":

    test_subject = "Invoice INV-12345 pending"

    test_body = """
    Hi Team,

    We have submitted invoice INV-12345.
    The invoice is still pending and has not been posted.
    Please check the invoice status and advise.

    Regards,
    Vendor Team
    """

    result = classify_email(
        test_subject,
        test_body,
    )

    print("\n" + "=" * 60)
    print("MAILFLOW AX - AI CLASSIFICATION TEST")
    print("=" * 60)

    print(f"Category   : {result['category']}")
    print(f"Confidence : {result['confidence']}%")
    print(f"Reason     : {result['reason']}")

    print("=" * 60)