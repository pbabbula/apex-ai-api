from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
import requests
import io
import os
import cohere
import base64
from uuid import uuid4
from datetime import datetime, timedelta
import json

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

from docx import Document
from pptx import Presentation
import pandas as pd

# -------------------------------------------------------------------
# Config
# -------------------------------------------------------------------

co = cohere.Client(os.getenv("COHERE_API_KEY"))

BASE_URL = os.getenv(
    "BASE_URL",
    "https://apex-ai-api.onrender.com"
)

app = FastAPI()
FILE_METADATA = {}

# -------------------------------------------------------------------
# Storage helpers
# -------------------------------------------------------------------

def save_metadata(file_id, metadata):
    os.makedirs("file_store", exist_ok=True)
    with open(f"file_store/{file_id}.json", "w") as f:
        json.dump(metadata, f)
    FILE_METADATA[file_id] = metadata

def load_metadata(file_id):
    try:
        with open(f"file_store/{file_id}.json", "r") as f:
            return json.load(f)
    except:
        return None

# -------------------------------------------------------------------
# PDF creator
# -------------------------------------------------------------------

def create_pdf_buffer(text, project_id=None, project_name=None):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    width, height = A4
    x, y = 50, height - 50
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
            y = height - 50
            c.setFont("Helvetica", 10)

        if not line.strip():
            y -= line_height
            continue

        if line.endswith(":") or line[:2].isdigit():
            c.setFont("Helvetica-Bold", 11)
        else:
            c.setFont("Helvetica", 10)

        c.drawString(x, y, line[:100])
        y -= line_height

    c.save()
    buffer.seek(0)
    return buffer

# -------------------------------------------------------------------
# HOME
# -------------------------------------------------------------------

@app.get("/")
def home():
    return {"status": "API running"}

# -------------------------------------------------------------------
# GENERATE DOC (ONLY PLACE WHERE AI IS CALLED)
# -------------------------------------------------------------------

@app.post("/generate-doc")
async def generate_doc(data: dict):

    try:
        project_id = data.get("project_id", "NA")
        project_name = data.get("project_name", "NA")
        file_url = data.get("file_url")

        if not file_url:
            return JSONResponse(
                status_code=400,
                content={"status": "ERROR", "message": "file_url required"}
            )

        # Download file
        response = requests.get(file_url, timeout=30)
        response.raise_for_status()
        file_content = response.content

        file_type = file_url.split(".")[-1].lower().split("?")[0]
        text_content = ""

        # Parse file
        if file_type == "docx":
            doc = Document(io.BytesIO(file_content))
            text_content = "\n".join(p.text for p in doc.paragraphs if p.text.strip())

        elif file_type == "pptx":
            prs = Presentation(io.BytesIO(file_content))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text_content += shape.text + "\n"

        elif file_type in ["xlsx", "xls"]:
            df = pd.read_excel(io.BytesIO(file_content))
            text_content = df.to_string(index=False)

        elif file_type in ["txt", "md"]:
            text_content = file_content.decode("utf-8", errors="ignore")

        else:
            return {"status": "ERROR", "message": "Unsupported file"}

        text_content = text_content[:3000]

        # AI CALL ✅ ONLY HERE
        prompt = f"""
Generate a professional Functional Specification Document.

Project: {project_name}
Content:
{text_content}
"""

        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text

        # Create PDF
        pdf_buffer = create_pdf_buffer(fsd_output, project_id, project_name)
        pdf_bytes = pdf_buffer.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode()

        file_id = str(uuid4())

        metadata = {
            "project_id": project_id,
            "project_name": project_name,
            "file_url": file_url,
            "file_type": file_type,
            "fsd_output": fsd_output,
            "pdf_base64": pdf_base64,
            "created_at": datetime.now().isoformat()
        }

        save_metadata(file_id, metadata)

        return {
            "status": "SUCCESS",
            "document": fsd_output,
            "download_link": f"{BASE_URL}/download-inline/{file_id}"
        }

    except Exception as e:
        return {"status": "ERROR", "message": str(e)}

# -------------------------------------------------------------------
# DOWNLOAD (NO AI CALL ✅)
# -------------------------------------------------------------------

@app.get("/download-inline/{file_id}")
def download_inline(file_id: str):

    metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

    if not metadata:
        return JSONResponse(status_code=404, content={"status": "ERROR", "message": "Not found"})

    # ✅ NO AI CALL HERE
    fsd_output = metadata["fsd_output"]

    pdf_buffer = create_pdf_buffer(
        fsd_output,
        metadata["project_id"],
        metadata["project_name"]
    )

    return StreamingResponse(
        pdf_buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=FSD.pdf"}
    )

# -------------------------------------------------------------------
# REFRESH (NO AI CALL ✅ just rebuild PDF)
# -------------------------------------------------------------------

@app.post("/refresh-doc/{file_id}")
async def refresh_doc(file_id: str):

    metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

    if not metadata:
        return JSONResponse(status_code=404, content={"status": "ERROR", "message": "Not found"})

    # ✅ REUSE EXISTING DOCUMENT
    fsd_output = metadata["fsd_output"]

    pdf_buffer = create_pdf_buffer(
        fsd_output,
        metadata["project_id"],
        metadata["project_name"]
    )

    pdf_bytes = pdf_buffer.getvalue()
    pdf_base64 = base64.b64encode(pdf_bytes).decode()

    metadata["pdf_base64"] = pdf_base64
    save_metadata(file_id, metadata)

    return {
        "status": "SUCCESS",
        "message": "PDF regenerated (no AI call)"
    }
