import os

from google import genai

api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise SystemExit("Set GEMINI_API_KEY to run this optional Gemini experiment.")

client = genai.Client(api_key=api_key)
user_query = input("Enter your query: ")
response = client.models.generate_content(
    model="gemini-2.5-flash",
    contents=user_query,
)
print("Gemini Response:")
print(response.text)