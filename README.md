# Пивот

Сводная статистика слов по персонажам и сериям для студий дубляжа/локализации.
Загружаешь несколько xlsx-файлов серий → получаешь единый xlsx со сводными пивотами,
плоской таблицей и детализацией по каждой серии.

Стек: **FastAPI + vanilla JS**, тёмная тема, кастомный UI.

## Что на входе

По умолчанию — профиль `default`:

- формат файлов: `.xlsx`
- внутри каждого файла — лист `Word Count Summary`
- номер серии берётся из имени файла по паттерну `(\d+)\s*СЕРИЯ`
  (подходит `1 СЕРИЯ FINAL.xlsx`, `10 СЕРИЯ…` и т.п.)
- колонки листа (в порядке): Персонаж, Dialog WC, Transcription WC, Foreign Dialogue,
  Music And Song, Burnedin Subtitle, Onscreen Text, Total WC

Для других сериалов с другим форматом — добавь запись в `PROFILES`
в [consolidator.py](consolidator.py).

## Что на выходе

Один xlsx с листами:

1. **Сводная по сериям** — пивот персонаж × серия по Total Word Count
2. **Диалоги по сериям** — тот же пивот по Dialog Word Count
3. **Все данные** — длинная плоская таблица
4. **Итоги по сериям** — агрегаты по каждой серии
5. **Серия N** — по одному листу на серию с разбивкой по персонажам

## Запуск локально

```bash
cd pivot
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

uvicorn app:app --reload --port 8000
```

Открой `http://localhost:8000`.

## CLI (без веба)

```bash
python consolidator.py -i /path/to/folder -o out.xlsx
python consolidator.py -i /path/to/folder --profile default
```

## Эндпоинты

| Метод | Путь           | Назначение                                             |
| ----- | -------------- | ------------------------------------------------------ |
| GET   | `/`            | главная (index.html)                                   |
| GET   | `/profiles`    | список доступных профилей (JSON)                       |
| POST  | `/consolidate` | multipart: `profile`, `files[]` → xlsx в ответе        |

Ответ `POST /consolidate` — бинарный xlsx. Мета (серии, персонажи, предупреждения)
передаётся в header `X-Pivot-Info` (JSON), имя файла — в `Content-Disposition`
(RFC 5987 UTF-8).

## Структура

```
app.py                 FastAPI: /, /profiles, /consolidate
consolidator.py        ядро: профили, чтение, пивоты, запись xlsx; CLI
static/index.html      frontend (inline CSS + JS, dark+violet)
requirements.txt
```

## Планы

- [ ] Деплой на Render
- [ ] Добавить профили для других сериалов по мере появления
- [ ] Опционально — UI для редактирования профиля на лету
