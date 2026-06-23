from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
import requests
import io
import os
import cohere
import base64
from uuid import uuid4
from datetime import datetime
import json
import textwrap
from urllib.parse import urlparse

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
from pptx import Presentation
import pandas as pd


# -------------------------------------------------------------------
# Config
# -------------------------------------------------------------------

COHERE_API_KEY = os.getenv("COHERE_API_KEY")

if not COHERE_API_KEY:
    raise RuntimeError("COHERE_API_KEY environment variable is missing")

co = cohere.Client(COHERE_API_KEY)

BASE_URL = os.getenv(
    "BASE_URL",
    "https://apex-ai-api.onrender.com"
)

# If you later use Render Persistent Disk, set FILE_STORE_DIR to mounted path.
# Example: /var/data/file_store
FILE_STORE_DIR = os.getenv("FILE_STORE_DIR", "file_store")

app = FastAPI(title="APEX AI FSD Generator")

FILE_METADATA = {}


# -------------------------------------------------------------------
# Storage helpers
# -------------------------------------------------------------------

def ensure_store_dir():
    os.makedirs(FILE_STORE_DIR, exist_ok=True)


def get_metadata_path(file_id):
    ensure_store_dir()
    return os.path.join(FILE_STORE_DIR, f"{file_id}.json")


def save_metadata(file_id, metadata):
    ensure_store_dir()

    file_path = get_metadata_path(file_id)

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False)

    FILE_METADATA[file_id] = metadata


def load_metadata(file_id):
    try:
        file_path = get_metadata_path(file_id)

        if not os.path.exists(file_path):
            return None

        with open(file_path, "r", encoding="utf-8") as f:
            metadata = json.load(f)

        FILE_METADATA[file_id] = metadata
        return metadata

    except Exception:
        return None


# -------------------------------------------------------------------
# File helpers
# -------------------------------------------------------------------

def get_file_extension(file_url):
    parsed = urlparse(file_url)
    path = parsed.path

    if "." not in path:
        return ""

    return path.split(".")[-1].lower()


def extract_text_from_file(file_content, file_type):
    text_content = ""

    if file_type == "docx":
        doc = Document(io.BytesIO(file_content))
        text_content = "\n".join(
            p.text for p in doc.paragraphs if p.text.strip()
        )

    elif file_type == "pptx":
        prs = Presentation(io.BytesIO(file_content))

        slide_texts = []

        for slide_index, slide in enumerate(prs.slides, start=1):
            slide_texts.append(f"\nSlide {slide_index}:")

            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_texts.append(shape.text.strip())

        text_content = "\n".join(slide_texts)

    elif file_type in ["xlsx", "xls"]:
        engine = "openpyxl" if file_type == "xlsx" else "xlrd"

        excel_data = pd.read_excel(
            io.BytesIO(file_content),
            sheet_name=None,
            engine=engine
        )

        sheet_texts = []

        for sheet_name, df in excel_data.items():
            sheet_texts.append(f"\nSheet: {sheet_name}")
            sheet_texts.append(df.to_string(index=False))

        text_content = "\n".join(sheet_texts)

    elif file_type in ["txt", "md"]:
        text_content = file_content.decode("utf-8", errors="ignore")

    else:
        raise ValueError("Unsupported file type")

    return text_content.strip()


# -------------------------------------------------------------------
# PDF creator
# -------------------------------------------------------------------

def create_pdf_buffer(text, project_id=None, project_name=None):
    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    x = 50
    y = height - 50
    line_height = 15
    max_chars_per_line = 95

    def new_page():
        nonlocal y
        c.showPage()
        y = height - 50
        c.setFont("Helvetica", 10)

    def draw_line(line, font_name="Helvetica", font_size=10):
        nonlocal y

        if y < 50:
            new_page()

        c.setFont(font_name, font_size)
        c.drawString(x, y, line)
        y -= line_height

    c.setFont("Helvetica-Bold", 15)
    c.drawString(x, y, "Functional Specification Document")
    y -= 30

    if project_id:
        draw_line(f"Project ID: {project_id}", "Helvetica", 10)

    if project_name:
        draw_line(f"Project Name: {project_name}", "Helvetica", 10)

    y -= 10

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            y -= line_height
            continue

        is_heading = (
            line.endswith(":")
            or line.startswith("#")
            or line[:2].isdigit()
            or line.lower().startswith((
                "section",
                "module",
                "requirement",
                "overview",
                "scope",
                "assumption",
                "dependency",
                "acceptance",
                "conclusion"
            ))
        )

        font_name = "Helvetica-Bold" if is_heading else "Helvetica"
        font_size = 11 if is_heading else 10

        wrapped_lines = textwrap.wrap(
            line,
            width=max_chars_per_line,
            replace_whitespace=False,
            drop_whitespace=True
        )

        for wrapped_line in wrapped_lines:
            draw_line(wrapped_line, font_name, font_size)

    c.save()
    buffer.seek(0)

    return buffer


# -------------------------------------------------------------------
# AI prompt
# -------------------------------------------------------------------

def build_fsd_prompt(project_id, project_name, text_content):
    return f"""
Generate a professional Functional Specification Document.

Use the following structure:

1. Document Overview
2. Project Details
3. Business Objective
4. Scope
5. Functional Requirements
6. User Roles and Responsibilities
7. Process Flow
8. Input Requirements
9. Output Requirements
10. Validation Rules
11. Error Handling
12. Assumptions
13. Dependencies
14. Acceptance Criteria
15. Conclusion

Project ID: {project_id}
Project Name: {project_name}

Source Content:
{text_content}
"""


# -------------------------------------------------------------------
# HOME
# -------------------------------------------------------------------

@app.get("/")
def home():
    return {
        "status": "API running",
        "storage": FILE_STORE_DIR
    }


# -------------------------------------------------------------------
# DEBUG FILES
# -------------------------------------------------------------------

@app.get("/debug-files")
def debug_files():
    ensure_store_dir()

    return {
        "status": "SUCCESS",
        "store_dir": FILE_STORE_DIR,
        "memory_ids": list(FILE_METADATA.keys()),
        "disk_files": os.listdir(FILE_STORE_DIR)
    }


# -------------------------------------------------------------------
# GENERATE DOC
# -------------------------------------------------------------------

@app.post("/generate-doc")
async def generate_doc(data: dict):
    try:
        project_id = str(data.get("project_id") or "NA").strip()
        project_name = str(data.get("project_name") or "NA").strip()

        file_url = data.get("file_url")
        manual_content = data.get("manual_content")

        text_content = ""
        file_type = "manual"

        # ------------------------------------------------------------
        # Option 1: Manual content from APEX
        # ------------------------------------------------------------
        if manual_content and str(manual_content).strip():
            text_content = str(manual_content).strip()
            file_type = "manual"

        # ------------------------------------------------------------
        # Option 2: File URL
        # ------------------------------------------------------------
        elif file_url and str(file_url).strip():
            file_url = str(file_url).strip()
            file_type = get_file_extension(file_url)

            supported_types = ["docx", "pptx", "xlsx", "xls", "txt", "md"]

            if file_type not in supported_types:
                return JSONResponse(
                    status_code=400,
                    content={
                        "status": "ERROR",
                        "message": "Unsupported file type. Supported types: docx, pptx, xlsx, xls, txt, md"
                    }
                )

            response = requests.get(file_url, timeout=30)
            response.raise_for_status()

            file_content = response.content
            text_content = extract_text_from_file(file_content, file_type)

        else:
            return JSONResponse(
                status_code=400,
                content={
                    "status": "ERROR",
                    "message": "Either manual_content or file_url is required"
                }
            )

        text_content = text_content[:6000]

        if not text_content.strip():
            return JSONResponse(
                status_code=400,
                content={
                    "status": "ERROR",
                    "message": "No readable content found"
                }
            )

        # ------------------------------------------------------------
        # AI CALL - ONLY HERE
        # ------------------------------------------------------------
        prompt = build_fsd_prompt(
            project_id=project_id,
            project_name=project_name,
            text_content=text_content
        )

        ai_response = co.chat(
            model="command-a-03-2025",
            message=prompt,
            temperature=0.3
        )

        fsd_output = ai_response.text.strip()

        if not fsd_output:
            return JSONResponse(
                status_code=500,
                content={
                    "status": "ERROR",
                    "message": "AI returned empty response"
                }
            )

        # ------------------------------------------------------------
        # Create PDF
        # ------------------------------------------------------------
        pdf_buffer = create_pdf_buffer(
            fsd_output,
            project_id,
            project_name
        )

        pdf_bytes = pdf_buffer.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

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
            "file_id": file_id,
            "download_link": f"{BASE_URL}/download-inline/{file_id}"
        }

    except requests.exceptions.RequestException as e:
        return JSONResponse(
            status_code=400,
            content={
                "status": "ERROR",
                "message": f"File download failed: {str(e)}"
            }
        )

    except ValueError as e:
        return JSONResponse(
            status_code=400,
            content={
                "status": "ERROR",
                "message": str(e)
            }
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "ERROR",
                "message": str(e)
            }
        )


# -------------------------------------------------------------------
# DOWNLOAD - NO AI CALL
# -------------------------------------------------------------------

@app.get("/download-inline/{file_id}")
def download_inline(file_id: str):
    try:
        metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

        if not metadata:
            return JSONResponse(
                status_code=404,
                content={
                    "status": "ERROR",
                    "message": "Not found"
                }
            )

        # Prefer saved PDF base64
        if metadata.get("pdf_base64"):
            pdf_bytes = base64.b64decode(metadata["pdf_base64"])
            pdf_buffer = io.BytesIO(pdf_bytes)
            pdf_buffer.seek(0)
        else:
            # Fallback: regenerate PDF from saved FSD text
            pdf_buffer = create_pdf_buffer(
                metadata["fsd_output"],
                metadata["project_id"],
                metadata["project_name"]
            )

        project_id = metadata.get("project_id") or "NA"
        filename = f"FSD_{project_id}.pdf"

        return StreamingResponse(
            pdf_buffer,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"'
            }
        )

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "ERROR",
                "message": str(e)
            }
        )


# -------------------------------------------------------------------
# REFRESH - NO AI CALL
# -------------------------------------------------------------------

@app.post("/refresh-doc/{file_id}")
async def refresh_doc(file_id: str):
    try:
        metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

        if not metadata:
            return JSONResponse(
                status_code=404,
                content={
                    "status": "ERROR",
                    "message": "Not found"
                }
            )

        pdf_buffer = create_pdf_buffer(
            metadata["fsd_output"],
            metadata["project_id"],
            metadata["project_name"]
        )

        pdf_bytes = pdf_buffer.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

        metadata["pdf_base64"] = pdf_base64
        metadata["refreshed_at"] = datetime.now().isoformat()

        save_metadata(file_id, metadata)

        return {
            "status": "SUCCESS",
            "message": "PDF regenerated without AI call",
            "download_link": f"{BASE_URL}/download-inline/{file_id}"
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "ERROR",
                "message": str(e)
            }
        )


# -------------------------------------------------------------------
# GET DOCUMENT TEXT - NO AI CALL
# -------------------------------------------------------------------

@app.get("/document/{file_id}")
def get_document(file_id: str):
    try:
        metadata = FILE_METADATA.get(file_id) or load_metadata(file_id)

        if not metadata:
            return JSONResponse(
                status_code=404,
                content={
                    "status": "ERROR",
                    "message": "Not found"
                }
            )

        return {
            "status": "SUCCESS",
            "file_id": file_id,
            "project_id": metadata.get("project_id"),
            "project_name": metadata.get("project_name"),
            "file_url": metadata.get("file_url"),
            "file_type": metadata.get("file_type"),
            "document": metadata.get("fsd_output"),
            "download_link": f"{BASE_URL}/download-inline/{file_id}",
            "created_at": metadata.get("created_at")
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "ERROR",
                "message": str(e)
            }
        )
