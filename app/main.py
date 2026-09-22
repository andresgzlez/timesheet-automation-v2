"""
App web de demostracion: subir fuente + plantilla, procesar con el motor
central (eligiendo automaticamente el pipeline correcto: estandar,
Project Daisy, o multi-codigo como HCA -- ver app/router.py) y mostrar la
pantalla de revision con el Excel resultante para descargar.
"""

import sys
import uuid
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse
from jinja2 import Environment, FileSystemLoader

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.router import process_upload

UPLOAD_DIR = ROOT / "app" / "uploads"
OUTPUT_DIR = ROOT / "app" / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Timesheet Automation")
jinja_env = Environment(loader=FileSystemLoader(str(ROOT / "app" / "templates")))


def render(template_name: str, **context) -> HTMLResponse:
    template = jinja_env.get_template(template_name)
    return HTMLResponse(template.render(**context))


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return render("index.html")


@app.post("/process", response_class=HTMLResponse)
async def process_files(
    request: Request,
    project: str = Form(...),
    source_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    job_id = uuid.uuid4().hex[:10]
    source_ext = ".xls" if source_file.filename.lower().endswith(".xls") else ".xlsx"
    source_path = UPLOAD_DIR / f"{job_id}_source{source_ext}"
    template_path = UPLOAD_DIR / f"{job_id}_template.xlsx"
    source_path.write_bytes(await source_file.read())
    template_path.write_bytes(await template_file.read())

    output_name = f"{project} - RESULTADO.xlsx".replace("/", "-")
    output_token = f"{job_id}__{output_name}"
    output_path = OUTPUT_DIR / output_token

    outcome = process_upload(project, source_path, template_path, output_path)

    if outcome.error and outcome.matched_count == 0:
        return render("error.html", message=outcome.error)

    return render(
        "result.html",
        profile_name=project,
        pipeline=outcome.pipeline,
        matched_count=outcome.matched_count,
        missing_people=outcome.missing_people,
        ambiguous=outcome.ambiguous,
        breakdown=outcome.breakdown,
        notes=outcome.totals.get("notes", []),
        cobrar_warning=outcome.error,
        download_token=output_token,
        output_name=output_name,
    )


@app.get("/download/{token}")
def download(token: str):
    path = OUTPUT_DIR / token
    filename = token.split("__", 1)[1] if "__" in token else token
    return FileResponse(path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
