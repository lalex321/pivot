"""FastAPI backend for Пивот."""

from __future__ import annotations

import json
import os
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles

from consolidator import PROFILES, consolidate

app = FastAPI(title="Пивот", docs_url=None, redoc_url=None)

STATIC_DIR = Path(__file__).parent / "static"
DATA_DIR = Path(os.environ.get("DATA_DIR", Path(__file__).parent / "_data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
USAGE_LOG = DATA_DIR / "usage.jsonl"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------- usage log ----------
def _client_ip(request: Request) -> str:
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    return request.client.host if request.client else "?"


def _log_usage(event: dict) -> None:
    try:
        with USAGE_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(event, ensure_ascii=False) + "\n")
    except OSError:
        pass  # лог не должен ронять основной запрос


def _read_usage(limit: int = 500) -> list[dict]:
    if not USAGE_LOG.exists():
        return []
    try:
        lines = USAGE_LOG.read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    events: list[dict] = []
    for line in reversed(lines):
        line = line.strip()
        if not line:
            continue
        try:
            events.append(json.loads(line))
        except json.JSONDecodeError:
            continue
        if len(events) >= limit:
            break
    return events


# ---------- routes ----------
@app.get("/")
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/admin")
async def admin_page() -> FileResponse:
    return FileResponse(STATIC_DIR / "admin.html")


@app.get("/admin/data")
async def admin_data(limit: int = 50) -> dict:
    events = _read_usage(limit=1000)
    now = datetime.now(timezone.utc)
    today_cutoff = now.replace(hour=0, minute=0, second=0, microsecond=0)
    week_cutoff = now - timedelta(days=7)

    today = week = success = failure = 0
    total_duration = total_files = total_episodes = total_characters = 0

    for e in events:
        ts_raw = e.get("ts", "")
        try:
            ts = datetime.fromisoformat(ts_raw.replace("Z", "+00:00"))
        except ValueError:
            continue
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        if ts >= today_cutoff:
            today += 1
        if ts >= week_cutoff:
            week += 1
        if e.get("ok"):
            success += 1
        else:
            failure += 1
        total_duration += int(e.get("duration_ms") or 0)
        total_files += int(e.get("files") or 0)
        total_episodes += int(e.get("episodes") or 0)
        total_characters += int(e.get("characters") or 0)

    runs = len(events)
    avg_duration = (total_duration // runs) if runs else 0

    return {
        "totals": {
            "runs": runs,
            "success": success,
            "failure": failure,
            "today": today,
            "week": week,
            "avg_duration_ms": avg_duration,
            "total_files": total_files,
            "total_episodes": total_episodes,
            "total_characters": total_characters,
        },
        "recent": events[: max(1, min(limit, 200))],
        "server_time": now.isoformat(),
    }


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
    request: Request,
    profile: str = Form("default"),
    files: list[UploadFile] = File(...),
) -> Response:
    started = time.time()
    ip = _client_ip(request)
    filenames = [(f.filename or "unknown.xlsx") for f in files]

    def _duration_ms() -> int:
        return int((time.time() - started) * 1000)

    if profile not in PROFILES:
        _log_usage({
            "ts": datetime.now(timezone.utc).isoformat(),
            "ok": False,
            "error": f"Неизвестный профиль: {profile}",
            "duration_ms": _duration_ms(),
            "files": len(files),
            "filenames": filenames,
            "ip": ip,
        })
        raise HTTPException(400, f"Неизвестный профиль: {profile}")
    if not files:
        _log_usage({
            "ts": datetime.now(timezone.utc).isoformat(),
            "ok": False,
            "error": "Нет файлов",
            "duration_ms": _duration_ms(),
            "files": 0,
            "ip": ip,
        })
        raise HTTPException(400, "Нет файлов")

    pairs: list[tuple[str, bytes]] = []
    for f in files:
        content = await f.read()
        pairs.append((f.filename or "unknown.xlsx", content))

    try:
        xlsx_bytes, info = consolidate(pairs, profile)
    except ValueError as e:
        _log_usage({
            "ts": datetime.now(timezone.utc).isoformat(),
            "ok": False,
            "error": str(e),
            "duration_ms": _duration_ms(),
            "files": len(files),
            "filenames": filenames,
            "profile": profile,
            "ip": ip,
        })
        raise HTTPException(400, str(e))

    common_name = (info.get("common_name") or "").strip()
    fname = f"{common_name} - Word Count Summary.xlsx" if common_name else "Word Count Summary.xlsx"
    info_header = json.dumps(info, ensure_ascii=True)

    _log_usage({
        "ts": datetime.now(timezone.utc).isoformat(),
        "ok": True,
        "duration_ms": _duration_ms(),
        "files": len(files),
        "filenames": filenames,
        "profile": profile,
        "episodes": len(info.get("episodes") or []),
        "characters": int(info.get("characters") or 0),
        "common_name": common_name,
        "warnings": len(info.get("warnings") or []),
        "ip": ip,
    })

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
