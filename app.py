#!/usr/bin/env python3
"""
FastAPI BOM Comparison Service
- Upload two BOM Excel files to compare and generate outputs.
- Supports dynamic column comparison selection via mapped pairs.
- Exposes a headers endpoint so the frontend can render selectable columns.
- Runs comparator and hierarchy generator in isolated per-request session dirs.
- MODIFIED: Unified 0-100% progress logic across both scripts.
- MODIFIED: Added support for Split Mappings (Old vs New) for Cross-Over mode.
- MODIFIED: Replaced simple dynamic list with mapped dynamic_pairs for optional comparison.
- FIX: Added favicon.ico route.
- FIX: Captures subprocess stderr/stdout to report actual crash reasons to frontend.
"""

import json
import os
import re
import shutil
import subprocess
import uuid
from pathlib import Path
from typing import Optional, List
import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse, Response
from fastapi.staticfiles import StaticFiles

# --- Configuration ---
BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
SESSION_FILES_DIR = BASE_DIR / "session_files"
SESSION_FILES_DIR.mkdir(exist_ok=True)

import sys
import inspect
import types

_orig_getsource = inspect.getsource
def _safe_getsource(obj):
    try:
        return _orig_getsource(obj)
    except OSError:
        if isinstance(obj, types.ModuleType):
            return ""
        raise
inspect.getsource = _safe_getsource

DEFAULT_SUMMARY_NAME = "summary.json"
CONFIG_JSON_PATH = str(BASE_DIR / "config91112.json")


app = FastAPI(
    title="BOM Comparison Service",
    description="Upload two BOM Excel files to compare them and generate a visual hierarchy.",
)


# Serve static assets
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# ------------- Utility helpers -------------
def _ensure_file(path: Path, what: str):
    if not path.is_file():
        raise HTTPException(status_code=404, detail=f"{what} not found: {path.name}")

def _save_upload(file: UploadFile, dest: Path):
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)

def _safe_report_name(old_filename: str, new_filename: str) -> str:
    def _slug(name: str) -> str:
        stem = Path(name).stem
        cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("_")
        return cleaned[:24] or "input"

    token = uuid.uuid4().hex[:8]
    return f"BOM_comparison_{_slug(old_filename)}_vs_{_slug(new_filename)}_{token}.xlsx"

def _merged_headers(old_path: Path, new_path: Path) -> tuple [List[str], List[str], List[str]]:
    try:
        # Optimization: Read only 0 rows just to get columns. Much faster.
        # Engine='openpyxl' ensures .xlsx compatibility
        df_old = pd.read_excel(old_path, engine='openpyxl', nrows=0)
        df_new = pd.read_excel(new_path, engine='openpyxl', nrows=0)
    except ImportError:
        raise HTTPException(status_code=500, detail="Server missing 'openpyxl' library. Please install it.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed reading Excel files: {str(e)}")
    
    cols_old = set(df_old.columns)
    cols_new = set(df_new.columns)
    
    # Return INTERSECTION, plus individuals
    common_cols = cols_old.intersection(cols_new)
    
    return sorted([str(c) for c in common_cols]), sorted([str(c) for c in cols_old]), sorted([str(c) for c in cols_new])

# ------------- Pages -------------
@app.get("/", include_in_schema=False)
async def read_root():
    index_path = STATIC_DIR / "index.html"
    _ensure_file(index_path, "index.html")
    return FileResponse(str(index_path))

@app.get("/visualise", include_in_schema=False)
def visualise_page():
    vis_file = STATIC_DIR / "visualise.html"
    _ensure_file(vis_file, "visualise.html")
    return FileResponse(str(vis_file))

# Fix for 404 Favicon error
@app.get('/favicon.ico', include_in_schema=False)
async def favicon():
    return Response(status_code=204)

# ------------- API: List headers before compare -------------
@app.post("/api/headers", tags=["BOM Comparison"])
async def list_bom_headers(
    old_bom_file: UploadFile = File(...),
    new_bom_file: UploadFile = File(...)
):
    session_id = str(uuid.uuid4())
    session_dir = SESSION_FILES_DIR / session_id
    session_dir.mkdir()

    tmp_old = session_dir / old_bom_file.filename
    tmp_new = session_dir / new_bom_file.filename
    try:
        _save_upload(old_bom_file, tmp_old)
        _save_upload(new_bom_file, tmp_new)
        common_headers, old_headers, new_headers = _merged_headers(tmp_old, tmp_new)

        return JSONResponse(
            status_code=200,
            content={
                "message": "Headers fetched.",
                "session_id": session_id,
                "headers_common": common_headers,
                "headers_old": old_headers,
                "headers_new": new_headers
            }
        )
    except Exception as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Failed to fetch headers: {e}")

# ------------- API: Compare and generate outputs (STREAMING) -------------
@app.post("/api/compare", tags=["BOM Comparison"])
async def compare_boms_and_generate_json(
    old_bom_file: UploadFile = File(...),
    new_bom_file: UploadFile = File(...),
    dynamic_pairs: Optional[str] = Form(default="[]"),
    print_columns: Optional[str] = Form(default=""),
    skip_choice: str = Form(default="Y"),
    config_path: Optional[str] = Form(default=""),
    map_old: Optional[str] = Form(default=None), 
    map_new: Optional[str] = Form(default=None),
):
    if skip_choice not in {"Y", "X"}:
        raise HTTPException(status_code=400, detail="skip_choice must be 'Y' or 'X'")

    session_id = str(uuid.uuid4())
    session_dir = SESSION_FILES_DIR / session_id
    session_dir.mkdir()

    # File Paths
    dynamic_excel_name = _safe_report_name(old_bom_file.filename, new_bom_file.filename)
    internal_excel_name = "comparison_report.xlsx"

    old_ext = Path(old_bom_file.filename).suffix or ".xlsx"
    new_ext = Path(new_bom_file.filename).suffix or ".xlsx"
    tmp_old_bom_path = session_dir / f"old_input{old_ext}"
    tmp_new_bom_path = session_dir / f"new_input{new_ext}"
    excel_output_path = session_dir / internal_excel_name
    json_output_path = session_dir / "bom_hierarchy.json"
    summary_path = session_dir / DEFAULT_SUMMARY_NAME

    # Save files
    _save_upload(old_bom_file, tmp_old_bom_path)
    _save_upload(new_bom_file, tmp_new_bom_path)

    # Config Path
    comparator_config_file = config_path if config_path else CONFIG_JSON_PATH
    if comparator_config_file:
        if not Path(comparator_config_file).is_file():
             if (BASE_DIR / comparator_config_file).is_file():
                 comparator_config_file = str(BASE_DIR / comparator_config_file)

    # 1. Comparator Command (-u for unbuffered)
    # Using sys.executable to call ourself in worker mode
    worker_base = [sys.executable] if getattr(sys, 'frozen', False) else [sys.executable, str(BASE_DIR / "app.py")]
    
    cmd_compare = worker_base + [
        "--worker", "comparator",
        "--old", str(tmp_old_bom_path),
        "--new", str(tmp_new_bom_path),
        "--output", str(excel_output_path),
        "--skip_choice", skip_choice,
    ]
    
    if dynamic_pairs: cmd_compare += ["--dynamic_pairs", dynamic_pairs]
    if print_columns: cmd_compare += ["--print_columns", print_columns]
    if comparator_config_file: cmd_compare += ["--config", comparator_config_file]
    if map_old: cmd_compare += ["--map_old", map_old]
    if map_new: cmd_compare += ["--map_new", map_new]

    def _resolve_generated_report_path(default_report_path: Path, comparator_logs: List[str]) -> Path:
        if default_report_path.is_file():
            return default_report_path

        # Prefer explicit comparator log output path when available.
        marker = "Exporting results to:"
        for line in reversed(comparator_logs):
            if marker in line:
                raw_path = line.split(marker, 1)[1].strip()
                if raw_path.endswith("..."):
                    raw_path = raw_path[:-3].strip()
                candidate = Path(raw_path)
                if not candidate.is_absolute():
                    candidate = session_dir / candidate
                if candidate.is_file():
                    return candidate
                break

        # Fallback: pick the most recently written report in the session folder.
        candidates = [
            p for p in session_dir.glob("*.xlsx")
            if p.name not in {tmp_old_bom_path.name, tmp_new_bom_path.name}
        ]
        if candidates:
            return max(candidates, key=lambda p: p.stat().st_mtime)

        return default_report_path


    MILESTONES = {
        "Reading OLD BOM": (5, "Reading Base BOM..."),
        "Reading NEW BOM": (10, "Reading New BOM..."),
        "Mapping & normalizing": (15, "Normalizing Data..."),
        "Step 1": (25, "Identifying Exact Matches..."),
        "Step 3": (40, "Running Fuzzy Analysis..."),
        "Classifying": (70, "Classifying Changes..."),
        "Updating node statuses": (75, "Updating Hierarchy Status..."),
        "Building output": (80, "Generating Excel Report..."),
        "Generating submodule summary": (85, "Analyzing Submodules..."),
        "Successfully updated": (95, "Finalizing Reports..."),
    }

    # --- Generator Function ---
    def process_generator():
        yield json.dumps({"progress": 1, "message": "Initializing..."}) + "\n"

        log_buffer = [] # Buffer to capture errors

        # --- Phase 1: Comparator ---
        try:
            process = subprocess.Popen(
                cmd_compare, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=session_dir, bufsize=1
            )
            
            for line in process.stdout:
                line = line.strip()
                if not line: continue
                log_buffer.append(line)
                if len(log_buffer) > 50: log_buffer.pop(0) # Keep last 50 lines
                
                # Check for milestones
                for key, (pct, msg) in MILESTONES.items():
                    if key in line:
                        yield json.dumps({"progress": pct, "message": msg}) + "\n"
                        break
                
            process.wait()
            if process.returncode != 0:
                # Send the buffered log as the error message
                error_details = "\n".join(log_buffer[-15:])
                yield json.dumps({"error": f"Comparator script failed (Code {process.returncode}):\n{error_details}"}) + "\n"
                return

        except Exception as e:
            yield json.dumps({"error": str(e)}) + "\n"
            return

        report_path_for_hierarchy = _resolve_generated_report_path(excel_output_path, log_buffer)
        report_path_for_download = report_path_for_hierarchy

        # --- Phase 2: Hierarchy ---
        yield json.dumps({"progress": 82, "message": "Building Hierarchy Tree..."}) + "\n"
        log_buffer = [] # Reset buffer

        # 2. Hierarchy Command (-u for unbuffered)
        cmd_hierarchy = worker_base + [
            "--worker", "hierarchy",
            "--input", str(report_path_for_hierarchy),
            "--output", str(json_output_path),
            "--submodules", str(BASE_DIR / "submodule_structure.json"),
            "--config_path", str(BASE_DIR / "config91112.json"),
            "--summary_json_path", str(summary_path),
            "--excel_report_path", str(report_path_for_hierarchy),
            "--manual_overrides_path", str(tmp_new_bom_path)
        ]
        
        try:
            process_h = subprocess.Popen(
                cmd_hierarchy, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=session_dir, bufsize=1
            )
            for line in process_h.stdout:
                line = line.strip()
                if not line: continue
                log_buffer.append(line)
                if len(log_buffer) > 50: log_buffer.pop(0)

                for key, (pct, msg) in MILESTONES.items():
                    if key in line:
                        yield json.dumps({"progress": pct, "message": msg}) + "\n"
                        break
                
            process_h.wait()
            if process_h.returncode != 0:
                error_details = "\n".join(log_buffer[-15:])
                yield json.dumps({"error": f"Hierarchy script failed (Code {process_h.returncode}):\n{error_details}"}) + "\n"
                return

        except Exception as e:
            yield json.dumps({"error": str(e)}) + "\n"
            return

        # --- Phase 3: Completion ---
        if not report_path_for_hierarchy.is_file() or not json_output_path.is_file():
             yield json.dumps({"error": "Critical: Output files missing."}) + "\n"
             return

        # Create alias after hierarchy updates so downloaded file includes final summary sheets.
        if dynamic_excel_name != report_path_for_hierarchy.name:
            aliased_download_path = session_dir / dynamic_excel_name
            try:
                if aliased_download_path.exists():
                    aliased_download_path.unlink()
                shutil.copy2(report_path_for_hierarchy, aliased_download_path)
                report_path_for_download = aliased_download_path
            except Exception:
                report_path_for_download = report_path_for_hierarchy

        result_payload = {
                "progress": 100,
                "message": "Processing Complete",
                "data": {
                    "session_id": session_id,
                    "excel_report_url": f"/download/{session_id}/{report_path_for_download.name}",
                    "json_data_url": f"/download/{session_id}/bom_hierarchy.json",
                    "summary_url": f"/download/{session_id}/{DEFAULT_SUMMARY_NAME}"
                }
        }
        yield json.dumps(result_payload) + "\n"

    return StreamingResponse(process_generator(), media_type="application/x-ndjson")

# ------------- Download route -------------
@app.get("/download/{session_id}/{filename}", tags=["Downloads"])
async def download_session_file(session_id: str, filename: str):
    file_path = SESSION_FILES_DIR / session_id / filename
    if not file_path.is_file():
        raise HTTPException(status_code=404, detail=f"File '{filename}' not found.")

    media_type = "application/octet-stream"
    if filename.endswith(".xlsx"):
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif filename.endswith(".json"):
        media_type = "application/json"

    return FileResponse(
        path=str(file_path),
        media_type=media_type,
        filename=filename
    )

if __name__ == "__main__":
    import uvicorn
    import sys
    import os
    
    # Check for worker mode
    if len(sys.argv) > 1 and sys.argv[1] == "--worker":
        worker_type = sys.argv[2]
        # Remove --worker and worker_type from argv so scripts don't get confused
        sys.argv = [sys.argv[0]] + sys.argv[3:] 
        if worker_type == "comparator":
            import bom_comp_levelcs
            bom_comp_levelcs.main()
        elif worker_type == "hierarchy":
            import generate_json_hierarchy
            generate_json_hierarchy.main()
        sys.exit(0)

    # Normal FastAPI startup
    port = int(os.environ.get("PORT", 9112))
    
    import socket
    def get_local_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('10.255.255.255', 1))
            ip = s.getsockname()[0]
        except Exception:
            ip = '127.0.0.1'
        finally:
            s.close()
        return ip
        
    local_ip = get_local_ip()
    print("="*60, flush=True)
    print(f"Share this link with others on your network: -> http://{local_ip}:{port} <-", flush=True)
    print("="*60, flush=True)
    
    if getattr(sys, 'frozen', False):
        uvicorn.run(app, host="0.0.0.0", port=port, log_config=None, access_log=False)
    else:
        uvicorn.run("app:app", host="0.0.0.0", port=port, reload=True)


