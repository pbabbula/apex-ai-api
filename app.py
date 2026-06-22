from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
import requests
import io
import os
import cohere
import base64
import binascii
from uuid import uuid4
from pathlib import Path

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

from docx import Document
from pptx import Presentation
import pandas as pd
from PyPDF2 import PdfReader

# -------------------------------------------------------------------
# Config
# -------------------------------------------------------------------

co = cohere.Client(os.getenv("COHERE_API_KEY"))

BASE_URL = os.getenv(
    "BASE_URL",
    "https://apex-ai-api.onrender.com"
)

app = FastAPI()
TEMP_STORE = {}

# -------------------------------------------------------------------
# Extract content (IMPROVED from 1st code)
# -------------------------------------------------------------------

def extract_file_content(file_bytes: bytes, file_type: str) -> str:
    file_type = file_type.lower()
    text = ""

    try:
        if file_type == "docx":
            doc = Document(io.BytesIO(file_bytes))

            for p in doc.paragraphs:
                if p.text.strip():
                    text += p.text + "\n"

            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_text:
                        text += " | ".join(row_text) + "\n"

        elif file_type == "pptx":
            prs = Presentation(io.BytesIO(file_bytes))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text + "\n"

        elif file_type in ["xlsx", "xls"]:
            df = pd.read_excel(io.BytesIO(file_bytes))
            text = df.to_string(index=False)

        elif file_type == "pdf":
            reader = PdfReader(io.BytesIO(file_bytes))
            for i, page in enumerate(reader.pages):
                page_text = page.extract_text() or ""
                text += f"\n--- Page {i+1} ---\n{page_text}"

        elif file_type in ["txt", "csv", "json", "xml", "html"]:
            text = file_bytes.decode("utf-8", errors="ignore")

        else:
            text = "Unsupported file type"

    except Exception as e:
        text = f"Error processing file: {str(e)}"

    return text


# -------------------------------------------------------------------
# PDF Creator (same as yours)
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

        c.drawString(x, y, clean_line)
        y -= line_height

    c.save()
    buffer.seek(0)
    return buffer


# -------------------------------------------------------------------
# MAIN ENDPOINT (ENHANCED)
# -------------------------------------------------------------------

@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        project_id = data.get("project_id", "NA")
        project_name = data.get("project_name", "NA")

        file_url = data.get("file_url")
        base64_content = data.get("base64_content")
        file_name = data.get("file_name", "file.bin")

        file_bytes = None
        file_type = None

        # -------------------------
        # OPTION 1: URL
        # -------------------------
        if file_url:
            response = requests.get(file_url, timeout=30)
            response.raise_for_status()

            file_bytes = response.content
            file_type = file_url.split(".")[-1].split("?")[0]

        # -------------------------
        # OPTION 2: BASE64
        # -------------------------
        elif base64_content:
            file_bytes = base64.b64decode(base64_content)

            suffix = Path(file_name).suffix
            file_type = suffix.replace(".", "")

        else:
            return JSONResponse(
                status_code=400,
                content={"status": "ERROR", "message": "Provide file_url or base64_content"}
            )

        # -------------------------
        # Extract content
        # -------------------------
        text_content = extract_file_content(file_bytes, file_type)

        text_content = text_content[:3000]

        # -------------------------
        # AI Prompt
        # -------------------------
        prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {project_id}
Project Name: {project_name}

STRICT RULES:
- No HTML
- Clean formatting
- Use headings

Content:
{text_content}

Sections:
1. Overview
2. Scope
3. Business Requirements
4. Functional Requirements
5. Use Cases
6. Assumptions
"""

        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text

        # -------------------------
        # Generate PDF
        # -------------------------
        pdf_buffer = create_pdf_buffer(fsd_output, project_id, project_name)
        pdf_bytes = pdf_buffer.getvalue()

        file_id = str(uuid4())
        TEMP_STORE[file_id] = pdf_bytes

        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        return {
            "status": "SUCCESS",
            "document": fsd_output,
            "pdf_base64": pdf_base64,
            "download_link": f"{BASE_URL}/download-inline/{file_id}",
            "detected_file_type": file_type
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "ERROR", "message": str(e)}
        )


# -------------------------------------------------------------------
# DOWNLOAD
# -------------------------------------------------------------------

@app.get("/download-inline/{file_id}")
def download_inline(file_id: str):

    if file_id not in TEMP_STORE:
        return JSONResponse(
            status_code=404,
            content={"status": "ERROR", "message": "File expired"}
        )

    return StreamingResponse(
        io.BytesIO(TEMP_STORE[file_id]),
        media_type="application/pdf",
        headers={
            "Content-Disposition": "attachment; filename=FSD.pdf",
            "Cache-Control": "no-store"
        }
    )
