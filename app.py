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
# Configuration
# -------------------------------------------------------------------

co = cohere.Client(os.getenv("COHERE_API_KEY"))

BASE_URL = os.getenv(
    "BASE_URL",
    "https://apex-ai-api.onrender.com"
)

app = FastAPI()

# Store metadata (instead of PDF bytes)
FILE_METADATA = {}

# -------------------------------------------------------------------
# Helper: Save metadata to persistent storage
# -------------------------------------------------------------------

def save_metadata(file_id, metadata):
    """Save metadata to a JSON file for persistence"""
    os.makedirs("file_store", exist_ok=True)
    with open(f"file_store/{file_id}.json", "w") as f:
        json.dump(metadata, f)
    FILE_METADATA[file_id] = metadata

def load_metadata(file_id):
    """Load metadata from persistent storage"""
    try:
        with open(f"file_store/{file_id}.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return None

def cleanup_expired_files():
    """Remove expired file metadata"""
    now = datetime.now().isoformat()
    for file_id in list(FILE_METADATA.keys()):
        if FILE_METADATA[file_id].get("expiry") < now:
            del FILE_METADATA[file_id]
            try:
                os.remove(f"file_store/{file_id}.json")
            except:
                pass

# -------------------------------------------------------------------
# Home endpoint
# -------------------------------------------------------------------

@app.get("/")
def home():
    return {"message": "API is live"}


# -------------------------------------------------------------------
# Create PDF in memory
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

        # New page if needed
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - 50

        clean_line = line.strip()

        if not clean_line:
            y -= line_height
            continue

        # Heading detection
        if clean_line.endswith(":") or clean_line[:2].isdigit():
            c.setFont("Helvetica-Bold", 11)
        else:
            c.setFont("Helvetica", 10)

        # Line wrap
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
        expiry_hours = data.get("expiry_hours", 99999)  # ~11 years (essentially permanent)

        if not file_url:
            return JSONResponse(
                status_code=400,
                content={"status": "ERROR", "message": "file_url is required"}
            )

        # -------------------------
        # Download file
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
            return JSONResponse(
                status_code=400,
                content={"status": "ERROR", "message": f"Unsupported file type: {file_type}"}
            )

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
        # Generate PDF
        # -------------------------
        pdf_buffer = create_pdf_buffer(
            fsd_output,
            project_id,
            project_name
        )

        pdf_bytes = pdf_buffer.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        file_id = str(uuid4())
        
        # Calculate expiry time
        expiry_time = (datetime.now() + timedelta(hours=expiry_hours)).isoformat()

        # Store metadata (not the PDF itself)
        metadata = {
            "project_id": project_id,
            "project_name": project_name,
            "file_url": file_url,
            "file_type": file_type,
            "fsd_output": fsd_output,
            "pdf_base64": pdf_base64,
            "created_at": datetime.now().isoformat(),
            "expiry": expiry_time,
            "expiry_hours": expiry_hours
        }

        save_metadata(file_id, metadata)

        return {
            "status": "SUCCESS",
            "document": fsd_output,
            "pdf_base64": pdf_base64,
            "download_link": f"{BASE_URL}/download-inline/{file_id}",
            "expires_in_hours": expiry_hours,
            "file_id": file_id
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "ERROR", "message": str(e)}
        )


# -------------------------------------------------------------------
# Download endpoint (regenerates PDF dynamically)
# -------------------------------------------------------------------

@app.get("/download-inline/{file_id}")
def download_inline(file_id: str):

    # Try to load from memory first, then from disk
    metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

    if not metadata:
        return JSONResponse(
            status_code=404,
            content={"status": "ERROR", "message": "File not found"}
        )

    # Check expiry
    if metadata.get("expiry") < datetime.now().isoformat():
        return JSONResponse(
            status_code=410,
            content={"status": "ERROR", "message": "File has expired"}
        )

    # -------------------------
    # DYNAMIC REGENERATION
    # -------------------------
    try:
        # Fetch fresh file content from URL
        response = requests.get(metadata["file_url"], timeout=30)
        response.raise_for_status()

        file_content = response.content
        file_type = metadata["file_type"]
        text_content = ""

        # Parse the fresh file
        if file_type == "docx":
            doc = Document(io.BytesIO(file_content))
            for para in doc.paragraphs:
                if para.text.strip():
                    text_content += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
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

        text_content = text_content[:3000]

        # Check if content changed
        old_content_hash = hash(metadata.get("fsd_output", ""))
        new_content_hash = hash(text_content)

        if old_content_hash != new_content_hash:
            # Content changed - regenerate FSD via AI
            prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {metadata["project_id"]}
Project Name: {metadata["project_name"]}

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
            ai_response = co.chat(
                model="command-a-03-2025",
                message=prompt,
                temperature=0.3
            )
            fsd_output = ai_response.text
            metadata["fsd_output"] = fsd_output
            save_metadata(file_id, metadata)
        else:
            fsd_output = metadata["fsd_output"]

        # Generate PDF
        pdf_buffer = create_pdf_buffer(
            fsd_output,
            metadata["project_id"],
            metadata["project_name"]
        )

        return StreamingResponse(
            pdf_buffer,
            media_type="application/pdf",
            headers={
                "Content-Disposition": "attachment; filename=FSD.pdf",
                "Cache-Control": "no-store"
            }
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "ERROR", "message": f"Failed to regenerate PDF: {str(e)}"}
        )


# -------------------------------------------------------------------
# Refresh/Update endpoint
# -------------------------------------------------------------------

@app.post("/refresh-doc/{file_id}")
async def refresh_doc(file_id: str):
    """Manually trigger regeneration of FSD for a file_id"""
    
    metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

    if not metadata:
        return JSONResponse(
            status_code=404,
            content={"status": "ERROR", "message": "File not found"}
        )

    if metadata.get("expiry") < datetime.now().isoformat():
        return JSONResponse(
            status_code=410,
            content={"status": "ERROR", "message": "File has expired"}
        )

    try:
        # Fetch fresh content and regenerate
        response = requests.get(metadata["file_url"], timeout=30)
        response.raise_for_status()

        file_content = response.content
        file_type = metadata["file_type"]
        text_content = ""

        if file_type == "docx":
            doc = Document(io.BytesIO(file_content))
            for para in doc.paragraphs:
                if para.text.strip():
                    text_content += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
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

        text_content = text_content[:3000]

        prompt = f"""
Generate a professional Functional Specification Document (FSD).

Project ID: {metadata["project_id"]}
Project Name: {metadata["project_name"]}

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

        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text
        pdf_buffer = create_pdf_buffer(fsd_output, metadata["project_id"], metadata["project_name"])
        pdf_bytes = pdf_buffer.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        metadata["fsd_output"] = fsd_output
        metadata["pdf_base64"] = pdf_base64
        save_metadata(file_id, metadata)

        return {
            "status": "SUCCESS",
            "message": "Document refreshed successfully",
            "document": fsd_output,
            "pdf_base64": pdf_base64
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"status": "ERROR", "message": str(e)}
        )
