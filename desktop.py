"""
Standalone launcher for the Windows .exe build.

Запускает FastAPI-сервер Pivot на свободном порту и открывает UI
в браузере по умолчанию. Предназначен как entrypoint для PyInstaller.

Пользовательские данные (usage-лог) кладутся в %LOCALAPPDATA%\\Pivot,
чтобы пережить перезапуск .exe (MEIPASS — временная папка).
"""
from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


def _persistent_data_dir() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "Pivot"
    return Path.home() / ".pivot"


# Должно быть установлено ДО импорта app — app.py читает DATA_DIR из env.
os.environ.setdefault("DATA_DIR", str(_persistent_data_dir()))


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main() -> None:
    import uvicorn
    from app import app as fastapi_app

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"

    config = uvicorn.Config(fastapi_app, host="127.0.0.1", port=port, log_level="warning")
    server = uvicorn.Server(config)

    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()

    # Ждём старта, чтобы браузер не открылся до слушателя.
    for _ in range(50):
        if server.started:
            break
        time.sleep(0.1)

    print()
    print(f"  Pivot работает на {url}")
    print(f"  Данные: {os.environ['DATA_DIR']}")
    print("  Закрой это окно, чтобы остановить сервер.")
    print()
    webbrowser.open(url)

    try:
        while thread.is_alive():
            thread.join(timeout=1.0)
    except KeyboardInterrupt:
        server.should_exit = True


if __name__ == "__main__":
    main()
