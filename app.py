from fastapi import FastAPI
from fastapi.responses import StreamingResponse
import requests
import io
import os
import cohere

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Document libraries
from docx import Document
from pptx import Presentation
import pandas as pd

# -------------------------------------------------------------------
# Configuration
# -------------------------------------------------------------------

co = cohere.Client(os.getenv("COHERE_API_KEY"))
app = FastAPI()

# -------------------------------------------------------------------
# Home endpoint
# -------------------------------------------------------------------

@app.get("/")
def home():
    return {"message": "API is live"}


# -------------------------------------------------------------------
# Create PDF in memory (NO FILE SAVING)
# -------------------------------------------------------------------

def create_pdf_buffer(text: str, project_id=None, project_name=None):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    width, height = A4

    x = 50
    y = height - 50
    line_height = 16

    c.setFont("Helvetica-Bold", 14)
    c.drawString(x, y, "Functional Specification Document")
    y -= 30

    c.setFont("Helvetica", 10)

    if project_id:
        c.drawString(x, y, f"Project ID: {project_id}")
        y -= line_height

    if project_name:
        c.drawString(x, y, f"Project Name: {project_name}")
        y -= line_height

    y -= 10

    for line in text.splitlines():
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - 50

        clean_line = line.strip()

        if not clean_line:
            y -= line_height
            continue

        if clean_line.endswith(":") or clean_line[:2].isdigit():
            c.setFont("Helvetica-Bold", 11)
        else:
            c.setFont("Helvetica", 10)

        max_chars = 100
        while len(clean_line) > max_chars:
            c.drawString(x, y, clean_line[:max_chars])
            clean_line = clean_line[max_chars:]
            y -= line_height

            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - 50

        c.drawString(x, y, clean_line)
        y -= line_height

    c.save()
    buffer.seek(0)

    return buffer


# -------------------------------------------------------------------
# Generate FSD endpoint
# -------------------------------------------------------------------

@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        project_id = data.get("project_id", "NA")
        project_name = data.get("project_name", "NA")
        file_url = data.get("file_url")

        if not file_url:
            return {"status": "ERROR", "message": "file_url is required"}

        # -------------------------
        # Download input file
        # -------------------------
        response = requests.get(file_url, timeout=30)
        response.raise_for_status()

        file_content = response.content
        file_type = file_url.split(".")[-1].lower().split("?")[0]

        text_content = ""

        # -------------------------
        # Parse file
        # -------------------------

        if file_type == "docx":
            doc = Document(io.BytesIO(file_content))

            for para in doc.paragraphs:
                if para.text.strip():
                    text_content += para.text + "\n"

            for table in doc.tables:
                for row in table.rows:
                    row_text = [
                        cell.text.strip()
                        for cell in row.cells
                        if cell.text.strip()
                    ]
                    if row_text:
                        text_content += " | ".join(row_text) + "\n"

        elif file_type == "pptx":
            prs = Presentation(io.BytesIO(file_content))

            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text_content += shape.text + "\n"

        elif file_type in ["xlsx", "xls"]:
            df = pd.read_excel(io.BytesIO(file_content))
            text_content = df.to_string(index=False)

        elif file_type in ["txt", "md"]:
            text_content = file_content.decode("utf-8", errors="ignore")

        else:
            return {"status": "ERROR", "message": f"Unsupported file type: {file_type}"}

        # Limit size
        text_content = text_content[:3000]

        # -------------------------
        # AI Prompt
        # -------------------------

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

        # -------------------------
        # Cohere API
        # -------------------------

        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text

        # -------------------------
        # Generate PDF (in memory)
        # -------------------------

        pdf_buffer = create_pdf_buffer(
            fsd_output,
            project_id,
            project_name
        )

        # -------------------------
        # Direct download response
        # -------------------------

        return StreamingResponse(
            pdf_buffer,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=FSD_{project_id}.pdf"
            }
        )

    except Exception as e:
        return {"status": "ERROR", "message": str(e)}
 
