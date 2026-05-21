"""
Turnitin Assistant v5.8-server-ready - FastAPI Server

REST API to trigger automation workflows.
Runs alongside or instead of the desktop pywebview app.
"""
import os
import sys
import threading
import logging
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional
from contextlib import asynccontextmanager

from fastapi import FastAPI, Header, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Ensure project root is in sys.path
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

import config

# ── In-memory log buffer ──
LOG_BUFFER = []


def log_to_buffer(message: str):
    """Log a message both to stdout and to the in-memory buffer."""
    timestamp = datetime.now().isoformat()
    entry = {"timestamp": timestamp, "message": message}
    LOG_BUFFER.append(entry)
    # Keep last 500 entries
    if len(LOG_BUFFER) > 500:
        LOG_BUFFER.pop(0)
    print(f"[{timestamp}] {message}")


# ── Workflow state ──
WORKFLOW_LOCKS = {
    "ojs_download": {"status": "idle", "message": "", "started_at": None},
    "turnitin_upload": {"status": "idle", "message": "", "started_at": None},
    "template_screening": {"status": "idle", "message": "", "started_at": None},
    "ojs_report_upload": {"status": "idle", "message": "", "started_at": None},
}


def get_workflow_runner(engine_name: str):
    """Import and return the workflow function for the given engine name."""
    if engine_name == "ojs_download":
        from engines import ojs_download_engine
        return ojs_download_engine.run_ojs_download_workflow
    elif engine_name == "turnitin_upload":
        from engines import turnitin_engine_api
        return turnitin_engine_api.run_turnitin_upload_workflow
    elif engine_name == "template_screening":
        from engines import template_screening_engine
        return template_screening_engine.run_template_screening_workflow
    elif engine_name == "ojs_report_upload":
        from engines import ojs_report_upload_engine
        return ojs_report_upload_engine.run_ojs_report_upload_workflow
    else:
        raise ValueError(f"Unknown engine: {engine_name}")


def run_workflow_in_thread(engine_name: str, lock_key: str):
    """Run a workflow in a background thread and update its status."""
    try:
        WORKFLOW_LOCKS[lock_key]["status"] = "running"
        WORKFLOW_LOCKS[lock_key]["started_at"] = datetime.now().isoformat()
        log_to_buffer(f"[SERVER] Starting {engine_name} workflow...")

        workflow_fn = get_workflow_runner(engine_name)

        def log_callback(message: str):
            log_to_buffer(message)

        result = workflow_fn(log_callback=log_callback)

        if result.get("status") == "success":
            WORKFLOW_LOCKS[lock_key]["status"] = "completed"
            WORKFLOW_LOCKS[lock_key]["message"] = result.get("message", "Completed successfully")
            log_to_buffer(f"[SERVER] ✅ {engine_name} workflow completed: {result.get('message', '')}")
        else:
            WORKFLOW_LOCKS[lock_key]["status"] = "failed"
            WORKFLOW_LOCKS[lock_key]["message"] = result.get("message", "Unknown failure")
            log_to_buffer(f"[SERVER] ❌ {engine_name} workflow failed: {result.get('message', '')}")

    except Exception as e:
        WORKFLOW_LOCKS[lock_key]["status"] = "failed"
        WORKFLOW_LOCKS[lock_key]["message"] = str(e)
        log_to_buffer(f"[SERVER] ❌ {engine_name} workflow error: {e}")
        print(f"[SERVER] Error in {engine_name}: {e}", file=sys.stderr)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Lifespan handler to initialise config on startup."""
    log_to_buffer("[SERVER] Turnitin Assistant v5.8-server-ready starting...")
    log_to_buffer(f"[SERVER] APP_MODE={os.environ.get('APP_MODE', 'server')}")
    config.ensure_directories()
    log_to_buffer("[SERVER] Config directories ensured.")
    yield
    log_to_buffer("[SERVER] Shutting down...")


app = FastAPI(
    title="Turnitin Assistant API",
    version="5.8-server-ready",
    lifespan=lifespan,
)

# CORS — allow all origins for Vercel frontend compatibility
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Auth helper ──
def verify_api_token(authorization: Optional[str] = None):
    """Check Bearer token if API_TOKEN is set in environment."""
    api_token = os.environ.get("API_TOKEN", "")
    if not api_token or api_token == "change-me":
        return  # No token configured = open access
    if not authorization:
        raise HTTPException(status_code=401, detail="Missing Authorization header")
    scheme, _, token = authorization.partition(" ")
    if scheme.lower() != "bearer" or token != api_token:
        raise HTTPException(status_code=403, detail="Invalid API token")


# ── Public endpoints ──

@app.get("/health")
def health():
    """Public health check."""
    return {
        "status": "ok",
        "version": "5.8-server-ready",
        "timestamp": datetime.now().isoformat(),
    }


# ── Protected API endpoints ──

@app.get("/api/status")
def api_status(authorization: Optional[str] = Header(None)):
    """Return application status."""
    verify_api_token(authorization)

    master_excel_path = Path(config.MASTER_EXCEL)
    master_excel_exists = master_excel_path.exists()

    return {
        "version": "5.8-server-ready",
        "mode": os.environ.get("APP_MODE", "server"),
        "master_excel_exists": master_excel_exists,
        "master_excel_path": str(config.MASTER_EXCEL),
        "storage": {
            "ojs_downloads": str(config.OJS_DOWNLOADS_DIR),
            "turnitin_reports": str(config.TURNITIN_REPORTS_DIR),
            "screening_reports": str(config.TEMPLATE_SCREENING_CONFIG["reports_dir"]),
        },
        "workflows": WORKFLOW_LOCKS,
    }


@app.post("/api/ojs-download/start")
def start_ojs_download(authorization: Optional[str] = Header(None)):
    """Start OJS download workflow in background."""
    verify_api_token(authorization)
    if WORKFLOW_LOCKS["ojs_download"]["status"] == "running":
        return JSONResponse(
            status_code=409,
            content={"error": "OJS Download workflow is already running", "status": "running"}
        )
    WORKFLOW_LOCKS["ojs_download"] = {"status": "queued", "message": "", "started_at": None}
    thread = threading.Thread(
        target=run_workflow_in_thread,
        args=("ojs_download", "ojs_download"),
        daemon=True,
    )
    thread.start()
    return {"status": "started", "workflow": "ojs_download"}


@app.post("/api/turnitin-upload/start")
def start_turnitin_upload(authorization: Optional[str] = Header(None)):
    """Start Turnitin upload workflow in background."""
    verify_api_token(authorization)
    if WORKFLOW_LOCKS["turnitin_upload"]["status"] == "running":
        return JSONResponse(
            status_code=409,
            content={"error": "Turnitin Upload workflow is already running", "status": "running"}
        )
    WORKFLOW_LOCKS["turnitin_upload"] = {"status": "queued", "message": "", "started_at": None}
    thread = threading.Thread(
        target=run_workflow_in_thread,
        args=("turnitin_upload", "turnitin_upload"),
        daemon=True,
    )
    thread.start()
    return {"status": "started", "workflow": "turnitin_upload"}


@app.post("/api/template-screening/start")
def start_template_screening(authorization: Optional[str] = Header(None)):
    """Start Template Screening workflow in background."""
    verify_api_token(authorization)
    if WORKFLOW_LOCKS["template_screening"]["status"] == "running":
        return JSONResponse(
            status_code=409,
            content={"error": "Template Screening workflow is already running", "status": "running"}
        )
    WORKFLOW_LOCKS["template_screening"] = {"status": "queued", "message": "", "started_at": None}
    thread = threading.Thread(
        target=run_workflow_in_thread,
        args=("template_screening", "template_screening"),
        daemon=True,
    )
    thread.start()
    return {"status": "started", "workflow": "template_screening"}


@app.post("/api/ojs-report-upload/start")
def start_ojs_report_upload(authorization: Optional[str] = Header(None)):
    """Start OJS Report Upload workflow in background."""
    verify_api_token(authorization)
    if WORKFLOW_LOCKS["ojs_report_upload"]["status"] == "running":
        return JSONResponse(
            status_code=409,
            content={"error": "OJS Report Upload workflow is already running", "status": "running"}
        )
    WORKFLOW_LOCKS["ojs_report_upload"] = {"status": "queued", "message": "", "started_at": None}
    thread = threading.Thread(
        target=run_workflow_in_thread,
        args=("ojs_report_upload", "ojs_report_upload"),
        daemon=True,
    )
    thread.start()
    return {"status": "started", "workflow": "ojs_report_upload"}


@app.get("/api/logs")
def get_logs(authorization: Optional[str] = Header(None)):
    """Return recent log entries."""
    verify_api_token(authorization)
    return {"logs": LOG_BUFFER[-200:]}


@app.get("/api/files/master-excel")
def download_master_excel(authorization: Optional[str] = Header(None)):
    """Download the master Excel file if it exists."""
    verify_api_token(authorization)
    excel_path = Path(config.MASTER_EXCEL)
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Master Excel file not found")
    return FileResponse(
        path=str(excel_path),
        filename=excel_path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Run directly for testing ──
if __name__ == "__main__":
    import uvicorn
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("server.main:app", host=host, port=port, reload=True)