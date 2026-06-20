from fastapi import FastAPI
import requests

# ✅ NEW: Add OpenAI
from openai import OpenAI

# ✅ Initialize AI client (will use API key from environment later)
client = OpenAI()

app = FastAPI(
    title="APEX AI API",
    version="1.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)


@app.get("/")
def home():
    return {"message": "API is live"}


@app.post("/generate-doc")
async def generate_doc(data: dict):

    project_id = data.get("project_id")
    project_name = data.get("project_name")
    file_url = data.get("file_url")

    try:
        file_content = requests.get(file_url).text
    except:
        file_content = "File not accessible"

    # ❌ OLD APPROACH (MANUAL HTML - NOT REAL FSD)
    '''
    html = f"""
    <h1>Functional Specification Document</h1>

    <h2>Project Info</h2>
    <p><b>ID:</b> {project_id}</p>
    <p><b>Name:</b> {project_name}</p>

    <h2>Details</h2>
    <p>{file_content[:1000]}</p>
    """
    return {"document": html}
    '''

    # ✅ ✅ NEW APPROACH — AI GENERATED FSD

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

    Return output in HTML format (use <h1>, <h2>, <p>, <ul>, <li>).
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "user", "content": prompt}
        ],
        temperature=0.3
    )

    fsd_output = response.choices[0].message.content

    return {"document": fsd_output}
