
from fastapi import FastAPI, File, UploadFile, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import shutil, os, uuid, datetime
from backend.logic import process_all_files

app = FastAPI()

UPLOAD_DIR = "backend/uploads"
OUTPUT_DIR = "backend/output"

templates = Jinja2Templates(directory="frontend")
app.mount("/static", StaticFiles(directory="frontend/static"), name="static")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate/")
async def generate_doc(
    word_template: UploadFile = File(...),
    ppt_files: list[UploadFile] = File(...)
):
    try:
        word_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{word_template.filename}")
        with open(word_path, "wb") as f:
            shutil.copyfileobj(word_template.file, f)

        ppt_paths = []
        for ppt in ppt_files:
            ppt_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{ppt.filename}")
            with open(ppt_path, "wb") as f:
                shutil.copyfileobj(ppt.file, f)
            ppt_paths.append(ppt_path)

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(OUTPUT_DIR, f"PMO_Report_{timestamp}.docx")

        process_all_files(word_path, ppt_paths, output_path)

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(output_path)
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
