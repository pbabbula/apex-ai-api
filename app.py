from fastapi import FastAPI
import requests
import cohere
import os

# ✅ Cohere client
co = cohere.Client(os.getenv("COHERE_API_KEY"))

app = FastAPI()


@app.get("/")
def home():
    return {"message": "API is live"}


@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        project_id = data.get("project_id")
        project_name = data.get("project_name")
        file_url = data.get("file_url")

        # ✅ Read file
        try:
            file_content = requests.get(file_url).text
        except Exception as e:
            return {"document": f"Error reading file: {str(e)}"}

        file_content = file_content[:3000]

        prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {project_id}
Project Name: {project_name}

Based on:
{file_content}

Create sections:
- Overview
- Scope
- Business Requirements
- Functional Requirements
- Use Cases
- Assumptions

Return in HTML format.
"""

        response = co.generate(
            model='command',
            prompt=prompt,
            max_tokens=800
        )

        fsd_output = response.generations[0].text

        return {"document": fsd_output}

    except Exception as e:
        return {"document": f"Error generating FSD: {str(e)}"}
