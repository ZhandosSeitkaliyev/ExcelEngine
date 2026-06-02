# ExcelEngine

Детерминированный HTTP-сервис для заполнения Excel-шаблонов данными.

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



