from fastapi import FastAPI
import requests
from openai import OpenAI
import os

# ✅ Initialize OpenAI client (reads from Render environment variable)
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

app = FastAPI()


@app.get("/")
def home():
    return {"message": "API is live"}


@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        # ✅ Step 1: Read inputs
        project_id = data.get("project_id")
        project_name = data.get("project_name")
        file_url = data.get("file_url")

        # ✅ Step 2: Fetch file content
        try:
            file_content = requests.get(file_url).text
        except Exception as e:
            return {"document": f"Error reading file: {str(e)}"}

        # ✅ Step 3: Limit content (IMPORTANT)
        file_content = file_content[:3000]

        # ✅ Step 4: Prepare AI prompt
        prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {project_id}
Project Name: {project_name}

Based on the following content:
{file_content}

Create structured sections:
1. Overview
2. Scope
3. Business Requirements
4. Functional Requirements
5. Use Cases
6. Assumptions

Return output in proper HTML format using:
<h1>, <h2>, <p>, <ul>, <li>.
"""

        # ✅ Step 5: Call OpenAI
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )

        fsd_output = response.choices[0].message.content

        # ✅ Step 6: ALWAYS return valid JSON
        return {"document": str(fsd_output)}

    except Exception as e:
        # ✅ Catch any unexpected errors (VERY IMPORTANT)
        return {"document": f"Error generating FSD: {str(e)}"}
