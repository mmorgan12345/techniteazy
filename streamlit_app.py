import google.generativeai as genai

import os
from dotenv import load_dotenv

#load the .env file
load_dotenv()

#access your api key

api_key = os.getenv("API_KEY")

# Automatically picks up GEMINI_API_KEY from your environment variables
client = genai.Clilent(api_key=API_KEY)

response = client.models.generate_content(
    model="gemini-3.7-flash",
    contents="Why is the sky blue?",
)

print(response.text)
