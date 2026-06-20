from fastapi import FastAPI
import requests

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

    html = f"""
    <h1>Functional Specification Document</h1>

    <h2>Project Info</h2>
    <p><b>ID:</b> {project_id}</p>
    <p><b>Name:</b> {project_name}</p>

    <h2>Details</h2>
    <p>{file_content[:1000]}</p>
    """

    return {"document": html}
