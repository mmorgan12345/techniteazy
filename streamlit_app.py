import google.generativeai as genai

API_KEY="AQ.Ab8RN6JwKfYEBzQcR-cBSxv8vHeyNRiVeARk1u_T3ioihu8DhA"

# Automatically picks up GEMINI_API_KEY from your environment variables
client = genai.Clilent(api_key=API_KEY)

response = client.models.generate_content(
    model="gemini-3.7-flash",
    contents="Why is the sky blue?",
)

print(response.text)
