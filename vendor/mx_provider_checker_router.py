from fastapi import APIRouter, Request, UploadFile, File, Form, Query
from fastapi.responses import HTMLResponse, StreamingResponse
import io
import time
from typing import List

from lib.file_handler import get_dataframe, store_dataframe, generate_csv_bytes, get_download_filename
from lib.mx_checker import start_task, get_task, cancel_task, auto_detect_column
from ._base import (
    get_templates, tool_context, handle_upload, handle_buffer_load,
    make_preview_html, error_html,
)

router = APIRouter(prefix="/csv/mx_provider_checker")
TOOL_ID = "mx_provider_checker"


@router.get("", response_class=HTMLResponse)
async def page(request: Request):
    ctx = tool_context(TOOL_ID)
    ctx["request"] = request
    return get_templates().TemplateResponse("pages/tool_page.html", ctx)


@router.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    file_id, df, filename, validation, err = await handle_upload(file, TOOL_ID)
    if err:
        return HTMLResponse(err)

    columns = df.columns.tolist()
    detected = auto_detect_column(columns)

    ctx = tool_context(TOOL_ID, file_id=file_id, df=df, filename=filename)
    ctx.update(
        request=request,
        validation=validation,
        preview_rows=make_preview_html(df),
        columns=columns,
        detected_column=detected,
    )
    return get_templates().TemplateResponse("pages/_mx_checker_workspace.html", ctx)


@router.post("/load-buffer", response_class=HTMLResponse)
async def load_buffer(request: Request):
    file_id, df, filename, validation, err = handle_buffer_load(TOOL_ID)
    if err:
        return HTMLResponse(err)

    columns = df.columns.tolist()
    detected = auto_detect_column(columns)

    ctx = tool_context(TOOL_ID, file_id=file_id, df=df, filename=filename)
    ctx.update(
        request=request,
        validation=validation,
        preview_rows=make_preview_html(df),
        columns=columns,
        detected_column=detected,
        loaded_from_buffer=True,
    )
    return get_templates().TemplateResponse("pages/_mx_checker_workspace.html", ctx)


@router.post("/process", response_class=HTMLResponse)
async def process(
    request: Request,
    file_id: str = Form(...),
    column: str = Form(...),
    workers: int = Form(30),
    timeout: int = Form(5),
):
    df, filename = get_dataframe(file_id)
    if df is None:
        return error_html("Session expired. Please re-upload.")

    if column not in df.columns:
        return error_html(f"Column '{column}' not found.")

    task_id = start_task(file_id, column, workers, timeout)

    ctx = {
        "request": request,
        "task_id": task_id,
        "tool_id": TOOL_ID,
        "completed": 0,
        "total": 0,
        "pct": 0,
        "elapsed": "0.0",
        "status": "starting",
    }
    return get_templates().TemplateResponse("pages/_mx_checker_progress.html", ctx)


@router.get("/progress/{task_id}", response_class=HTMLResponse)
async def progress(request: Request, task_id: str):
    task = get_task(task_id)
    if task is None:
        return error_html("Task not found.")

    elapsed = time.time() - task.start_time

    if task.status == "completed":
        # Return the result partial — polling stops (no hx-trigger)
        pct_map = {}
        for provider, count in task.provider_stats.items():
            pct_map[provider] = round(count / task.total_records * 100, 1) if task.total_records else 0

        sorted_stats = sorted(task.provider_stats.items(), key=lambda x: x[1], reverse=True)

        # Get result preview
        df, _ = get_dataframe(task.result_file_id)
        result_preview = make_preview_html(df) if df is not None else ""

        output_filename = get_download_filename(_ or 'data.csv', suffix='mx_checked')

        ctx = {
            "request": request,
            "tool_id": TOOL_ID,
            "task_id": task_id,
            "result_file_id": task.result_file_id,
            "result_preview": result_preview,
            "output_filename": output_filename,
            "total_records": task.total_records,
            "valid_domains": task.valid_domains,
            "elapsed": f"{elapsed:.1f}",
            "sorted_stats": sorted_stats,
            "pct_map": pct_map,
        }
        return get_templates().TemplateResponse("pages/_mx_checker_result.html", ctx)

    if task.status == "error":
        return error_html(task.error or "Unknown error during processing.")

    if task.status == "cancelled":
        return HTMLResponse(
            '<div class="bg-yellow-900/30 border border-yellow-700 text-yellow-300 px-4 py-3 rounded-lg text-sm">'
            'Processing cancelled.</div>'
        )

    # Still running — return progress partial (polling continues)
    pct = round(task.completed / task.total * 100, 1) if task.total else 0

    ctx = {
        "request": request,
        "task_id": task_id,
        "tool_id": TOOL_ID,
        "completed": task.completed,
        "total": task.total,
        "pct": pct,
        "elapsed": f"{elapsed:.1f}",
        "status": task.status,
    }
    return get_templates().TemplateResponse("pages/_mx_checker_progress.html", ctx)


@router.post("/cancel/{task_id}", response_class=HTMLResponse)
async def cancel(task_id: str):
    cancel_task(task_id)
    return HTMLResponse("")


@router.get("/download/{file_id}")
async def download(file_id: str):
    df, filename = get_dataframe(file_id)
    if df is None:
        return HTMLResponse("File not found", status_code=404)
    output_filename = get_download_filename(filename or 'data.csv', suffix='mx_checked')
    return StreamingResponse(
        io.BytesIO(generate_csv_bytes(df)),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{output_filename}"'},
    )


@router.get("/download/{file_id}/filtered")
async def download_filtered(file_id: str, providers: List[str] = Query(default=[])):
    df, filename = get_dataframe(file_id)
    if df is None:
        return HTMLResponse("File not found", status_code=404)

    if providers and 'mail_provider' in df.columns:
        df = df[df['mail_provider'].isin(providers)]

    output_filename = get_download_filename(filename or 'data.csv', suffix='mx_filtered')
    return StreamingResponse(
        io.BytesIO(generate_csv_bytes(df)),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{output_filename}"'},
    )

