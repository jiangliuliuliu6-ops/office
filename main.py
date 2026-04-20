import os
import uuid
import tempfile
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, UploadFile, Form, BackgroundTasks, Request
from fastapi.responses import JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from utils.processor import WordImageProcessor

app = FastAPI()

# Static files and templates
if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Temp directory for processed files
TEMP_DIR = Path(tempfile.gettempdir()) / "word_fixer_pro"
TEMP_DIR.mkdir(parents=True, exist_ok=True)


def cleanup(path: str):
    """Background task: delete the processed file after download."""
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# ─────────────────────────────────────────────────────────────
#  POST /api/process
#  Accepts the docx + images directly (no pre-upload step needed)
# ─────────────────────────────────────────────────────────────
@app.post("/api/process")
async def api_process(
    background_tasks: BackgroundTasks,
    docx_file: UploadFile = File(...),
    mode: str = Form("fix"),
    page_width: float = Form(16.0),
    scale: str = Form("0.9"),
    align_center: str = Form("true"),
    auto_width: str = Form("true"),
    images: List[UploadFile] = File(default=[]),
):
    try:
        # Convert form string booleans
        do_align_center = align_center.lower() not in ("false", "0", "")
        do_auto_width   = auto_width.lower()   not in ("false", "0", "")
        scale_f         = float(scale) if scale else 0.9

        session_id = str(uuid.uuid4())

        # Save uploaded docx to temp
        doc_suffix = Path(docx_file.filename or "document.docx").suffix or ".docx"
        doc_path   = TEMP_DIR / f"{session_id}_input{doc_suffix}"
        doc_path.write_bytes(await docx_file.read())

        # Save uploaded images to temp
        img_list: List[tuple] = []   # [(filename, bytes), ...]
        for img in images:
            if img and img.filename:
                img_bytes = await img.read()
                if img_bytes:
                    img_list.append((img.filename, img_bytes))

        # Process
        processor = WordImageProcessor(str(doc_path))
        result = processor.process(
            mode          = mode,
            images        = img_list,
            scale         = scale_f,
            align_center  = do_align_center,
            auto_width    = do_auto_width,
            page_width_cm = page_width,
        )

        # Save output
        original_stem = Path(docx_file.filename or "document").stem
        output_filename = f"fixed_{original_stem}.docx"
        output_path     = TEMP_DIR / f"{session_id}_output.docx"
        processor.save(str(output_path))

        # Clean up temp input
        background_tasks.add_task(cleanup, str(doc_path))

        return {
            "success":    True,
            "session_id": session_id,
            "filename":   output_filename,
            "matched":    result.get("matched", []),
            "unmatched":  result.get("unmatched", []),
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"success": False, "detail": str(e)},
        )


# ─────────────────────────────────────────────────────────────
#  GET /api/download/{session_id}
# ─────────────────────────────────────────────────────────────
@app.get("/api/download/{session_id}")
async def api_download(session_id: str, background_tasks: BackgroundTasks):
    # Find the output file for this session
    matches = list(TEMP_DIR.glob(f"{session_id}_output*.docx"))
    if not matches:
        return JSONResponse(status_code=404, content={"error": "File not found or already downloaded"})

    file_path = matches[0]
    background_tasks.add_task(cleanup, str(file_path))

    return FileResponse(
        path         = str(file_path),
        filename     = "fixed_document.docx",
        media_type   = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
