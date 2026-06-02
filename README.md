# ExcelEngine

Детерминированный HTTP-сервис для заполнения Excel-шаблонов данными.
Используется как внешний сервис из Dify.ai (встроенная песочница Dify не содержит openpyxl).

## Принцип защиты шаблона

Ячейки, которые **нельзя изменять** (заголовки, подписи, формулы), закрашиваются заливкой
`#00A0F0` в Excel. После заполнения данными сервис автоматически убирает эту заливку.

> **Ограничение:** маркером `#00A0F0` следует метить только те ячейки, чей итоговый вид —
> без цветного фона. Маркер кладётся «поверх» любого существующего фона и при удалении
> заменяется на «без заливки» — исходный цвет фона не восстанавливается.

## Структура проекта

```
ExcelEngine/
├── engine/
│   └── core.py          # analyze(), fill() — чистые функции, без сети
├── app/
│   └── main.py          # FastAPI HTTP-слой
├── tests/
│   ├── conftest.py      # фикстура шаблона
│   └── test_core.py     # 10 тестов ядра
├── requirements.txt
└── README.md
```

## Локальный запуск

```bash
# Установить зависимости
pip install -r requirements.txt

# Задать API-ключ
export API_KEY=mysecretkey          # Linux/macOS
set API_KEY=mysecretkey             # Windows CMD
$env:API_KEY="mysecretkey"          # Windows PowerShell

# Запустить сервер
uvicorn app.main:app --reload --port 8000
```

## Запуск тестов

```bash
cd ExcelEngine
python -m pytest tests/ -v
```

## API

### POST /analyze

Принимает файл шаблона, возвращает JSON-схему заполняемых полей.

```bash
curl -X POST http://localhost:8000/analyze \
  -H "X-API-Key: mysecretkey" \
  -F "template=@template.xlsx" \
  -F "header_row=1"
```

Ответ:

```json
{
  "sheet": "Sheet1",
  "header_row": 1,
  "fillable_fields": [
    {"name": "Name", "col_letter": "B", "col_index": 2},
    {"name": "Amount", "col_letter": "C", "col_index": 3}
  ],
  "protected_cells": [
    {"coordinate": "A1", "col_letter": "A", "row": 1},
    {"coordinate": "D1", "col_letter": "D", "row": 1}
  ]
}
```

### POST /fill

Принимает шаблон + JSON с данными, возвращает заполненный `.xlsx`.

Одна запись:

```bash
curl -X POST http://localhost:8000/fill \
  -H "X-API-Key: mysecretkey" \
  -F "template=@template.xlsx" \
  -F 'data={"Name": "Alice", "Amount": 1500}' \
  --output result.xlsx
```

Несколько записей:

```bash
curl -X POST http://localhost:8000/fill \
  -H "X-API-Key: mysecretkey" \
  -F "template=@template.xlsx" \
  -F 'data=[{"Name": "Alice", "Amount": 1500}, {"Name": "Bob", "Amount": 2200}]' \
  --output result.xlsx
```

Необязательный параметр `header_row` (по умолчанию `1`):

```bash
curl -X POST http://localhost:8000/fill \
  -H "X-API-Key: mysecretkey" \
  -F "template=@template.xlsx" \
  -F "header_row=2" \
  -F 'data={"Name": "Carol"}' \
  --output result.xlsx
```

## Деплой

### Railway

1. Создать новый проект, указать этот репозиторий.
2. В разделе **Variables** добавить `API_KEY=<ваш ключ>`.
3. В **Start command** указать:
   ```
   uvicorn app.main:app --host 0.0.0.0 --port $PORT
   ```

### Render

1. Создать **Web Service**, выбрать репозиторий.
2. **Build command:** `pip install -r requirements.txt`
3. **Start command:** `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
4. Добавить переменную окружения `API_KEY`.
