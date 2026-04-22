# Сборка Pivot.exe для Windows

Один бинарник для коллег без Python.

## Что нужно

- Windows VM (10/11) с установленным Python 3.11+ из [python.org](https://python.org)
- Этот репозиторий, склонированный внутри VM

## Шаги (внутри VM)

```cmd
cd path\to\pivot

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
pip install pyinstaller

pyinstaller pivot.spec --clean --noconfirm
```

Готовый `.exe` — в `dist\Pivot.exe` (~40-60 МБ).

## Как пользоваться

Коллега двойным кликом запускает `Pivot.exe`. Откроется:

1. Консоль с адресом `http://127.0.0.1:<порт>` и сообщением «не закрывай окно»
2. Браузер по умолчанию со страницей Pivot

Данные (usage-лог) сохраняются в `%LOCALAPPDATA%\Pivot`. Закрытие консоли
останавливает сервер.

## Что НЕ попадает в .exe

- HTTP Basic Auth (`BASIC_AUTH_USER/PASSWORD`) — по умолчанию выключен, в
  десктоп-сборке не нужен
- Render-конфиг (`render.yaml`) — это для веб-деплоя

## Автосборка через GitHub Actions (опционально)

Когда надоест каждый раз заходить в VM — можно добавить workflow
`.github/workflows/windows-exe.yml` с runner-ом `windows-latest`, и `.exe`
будет артефактом на каждый push/tag. Скажи — настрою.
