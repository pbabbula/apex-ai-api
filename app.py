from fastapi import FastAPI
import requests
import os
import io
import cohere

# Document libraries
from docx import Document
from pptx import Presentation
import pandas as pd

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

        # ✅ Download file
        response = requests.get(file_url)
        file_content = response.content

        file_type = file_url.split('.')[-1].lower()

        text_content = ""

        # ✅ DOCX
        if file_type == "docx":
            doc = Document(io.BytesIO(file_content))
            for para in doc.paragraphs:
                text_content += para.text + "\n"

        # ✅ PPTX
        elif file_type == "pptx":
            prs = Presentation(io.BytesIO(file_content))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text_content += shape.text + "\n"

        # ✅ EXCEL
        elif file_type in ["xlsx", "xls"]:
            df = pd.read_excel(io.BytesIO(file_content))
            text_content = df.to_string()

        # ✅ TEXT / MD
        elif file_type in ["txt", "md"]:
            text_content = file_content.decode("utf-8")

        else:
            text_content = "Unsupported file type"

        # ✅ Limit size for AI
        text_content = text_content[:3000]

        # ✅ CLEAN PROMPT (NO HTML)
        prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {project_id}
Project Name: {project_name}

STRICT RULES:
- Do NOT include HTML or CSS
- Return clean plain text only
- Use clear headings and bullet points

Content:
{text_content}

Structure:
1. Overview
2. Scope
3. Business Requirements
4. Functional Requirements
5. Use Cases
6. Assumptions
"""

        # ✅ Cohere Chat API (LATEST MODEL)
        response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = response.text

        return {"document": fsd_output}

    except Exception as e:
        return {"document": f"Error generating FSD: {str(e)}"}
