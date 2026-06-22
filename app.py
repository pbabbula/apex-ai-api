from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
import requests
import os
import io
import cohere
import mimetypes
from pathlib import Path
from uuid import uuid4
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Document libraries
from docx import Document
from pptx import Presentation
import pandas as pd

# -------------------------------------------------------------------
# Configuration
# -------------------------------------------------------------------

BASE_URL = os.getenv(
    "BASE_URL",
    "https://python-document-new-2.onrender.com"
)

FILES_DIR = Path(os.getcwd()) / "generated_files"
FILES_DIR.mkdir(parents=True, exist_ok=True)

# Cohere client
co = cohere.Client(os.getenv("COHERE_API_KEY"))

app = FastAPI()


# -------------------------------------------------------------------
# Home endpoint
# -------------------------------------------------------------------

@app.get("/")
def home():
    return {"message": "API is live"}


# -------------------------------------------------------------------
# Helper: Create PDF from generated FSD text
# -------------------------------------------------------------------

def create_pdf_from_text(text: str, pdf_path: Path, project_id=None, project_name=None):
    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4

    c.setTitle("Functional Specification Document")

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

        # Basic heading handling
        clean_line = line.strip()

        if not clean_line:
            y -= line_height
            continue

        if clean_line.endswith(":") or clean_line[:2].isdigit():
            c.setFont("Helvetica-Bold", 11)
        else:
            c.setFont("Helvetica", 10)

        # Wrap long lines manually
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


# -------------------------------------------------------------------
# Generate FSD endpoint
# -------------------------------------------------------------------

@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        project_id = data.get("project_id")
        project_name = data.get("project_name")
        file_url = data.get("file_url")

        if not file_url:
            return {
                "status": "ERROR",
                "message": "file_url is required"
            }

        # Download file
        response = requests.get(file_url, timeout=30)
        response.raise_for_status()

        file_content = response.content
        file_type = file_url.split(".")[-1].lower().split("?")[0]

        text_content = ""

        # DOCX
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

        # PPTX
        elif file_type == "pptx":
            prs = Presentation(io.BytesIO(file_content))

            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text_content += shape.text + "\n"

        # EXCEL
        elif file_type in ["xlsx", "xls"]:
            df = pd.read_excel(io.BytesIO(file_content))
            text_content = df.to_string(index=False)

        # TEXT / MD
        elif file_type in ["txt", "md"]:
            text_content = file_content.decode("utf-8", errors="ignore")

        else:
            return {
                "status": "ERROR",
                "message": f"Unsupported file type: {file_type}"
            }

        # Limit input size for AI
        text_content = text_content[:3000]

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

        # Cohere Chat API
        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text

        # ------------------------------------------------------------
        # Save generated FSD files
        # ------------------------------------------------------------

        unique_id = f"{project_id}_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{uuid4().hex[:8]}"

        txt_file_name = f"fsd_{unique_id}.txt"
        txt_file_path = FILES_DIR / txt_file_name

        pdf_file_name = f"fsd_{unique_id}.pdf"
        pdf_file_path = FILES_DIR / pdf_file_name

        # Save TXT
        with open(txt_file_path, "w", encoding="utf-8") as f:
            f.write(fsd_output)

        # Save PDF
        create_pdf_from_text(
            text=fsd_output,
            pdf_path=pdf_file_path,
            project_id=project_id,
            project_name=project_name
        )

        return {
            "status": "SUCCESS",
            "message": "FSD generated successfully",
            "document": fsd_output,

            "txt_file_name": txt_file_name,
            "txt_download_link": f"{BASE_URL}/download/{txt_file_name}",

            "pdf_file_name": pdf_file_name,
            "pdf_download_link": f"{BASE_URL}/download/{pdf_file_name}"
        }

    except Exception as e:
        return {
            "status": "ERROR",
            "message": f"Error generating FSD: {str(e)}"
        }


# -------------------------------------------------------------------
# Download endpoint
# -------------------------------------------------------------------

@app.get("/download/{file_name}")
def download_file(file_name: str):
    safe_file_name = Path(file_name).name
    file_path = FILES_DIR / safe_file_name

    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found")

    media_type, _ = mimetypes.guess_type(str(file_path))

    if not media_type:
        media_type = "application/octet-stream"

    return FileResponse(
        path=str(file_path),
        media_type=media_type,
        filename=safe_file_name,
        headers={
            "Content-Disposition": f'attachment; filename="{safe_file_name}"',
            "Cache-Control": "no-store"
        }
    )
