"""FastAPI backend for Пивот."""

from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles

from consolidator import PROFILES, consolidate

app = FastAPI(title="Пивот", docs_url=None, redoc_url=None)

STATIC_DIR = Path(__file__).parent / "static"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/profiles")
async def profiles_list() -> dict:
    return {
        k: {
            "name": v.name,
            "sheet": v.sheet_name,
            "pattern": v.episode_pattern,
        }
        for k, v in PROFILES.items()
    }


@app.post("/consolidate")
async def do_consolidate(
    profile: str = Form("default"),
    files: list[UploadFile] = File(...),
) -> Response:
    if profile not in PROFILES:
        raise HTTPException(400, f"Неизвестный профиль: {profile}")
    if not files:
        raise HTTPException(400, "Нет файлов")

    pairs: list[tuple[str, bytes]] = []
    for f in files:
        content = await f.read()
        pairs.append((f.filename or "unknown.xlsx", content))

    try:
        xlsx_bytes, info = consolidate(pairs, profile)
    except ValueError as e:
        raise HTTPException(400, str(e))

    fname = "Сводная статистика.xlsx"
    info_header = json.dumps(info, ensure_ascii=True)  # ASCII-safe for HTTP header

    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MIME,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}",
            "X-Pivot-Info": info_header,
            "Access-Control-Expose-Headers": "X-Pivot-Info, Content-Disposition",
        },
    )


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
