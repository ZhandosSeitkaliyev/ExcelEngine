import json
import os
import urllib.parse

from fastapi import Depends, FastAPI, File, Form, HTTPException, Security, UploadFile
from fastapi.responses import Response
from fastapi.security.api_key import APIKeyHeader

from engine.core import (
    analyze,
    fill,
    schema_to_dict,
    wb_from_bytes,
    wb_to_bytes,
)

app = FastAPI(title="ExcelEngine", version="1.0.0")

_API_KEY_NAME = "X-API-Key"
_api_key_header = APIKeyHeader(name=_API_KEY_NAME, auto_error=True)


def _require_api_key(key: str = Security(_api_key_header)) -> str:
    expected = os.environ.get("API_KEY", "")
    if not expected:
        raise HTTPException(status_code=500, detail="API_KEY не задан на сервере")
    if key != expected:
        raise HTTPException(status_code=403, detail="Неверный API-ключ")
    return key


@app.post("/analyze")
async def analyze_endpoint(
    template: UploadFile = File(...),
    header_row: int = Form(1),
    _: str = Depends(_require_api_key),
):
    raw = await template.read()
    try:
        wb = wb_from_bytes(raw)
        schema = analyze(wb, header_row=header_row)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Ошибка чтения файла: {exc}")
    return schema_to_dict(schema)


@app.post("/fill")
async def fill_endpoint(
    template: UploadFile = File(...),
    data: str = Form(...),
    header_row: int = Form(1),
    _: str = Depends(_require_api_key),
):
    raw = await template.read()
    try:
        records = json.loads(data)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=422, detail=f"Невалидный JSON в поле data: {exc}")

    if isinstance(records, dict):
        records = [records]
    if not isinstance(records, list):
        raise HTTPException(status_code=422, detail="data должен быть объектом или массивом объектов")

    try:
        wb = wb_from_bytes(raw)
        schema = analyze(wb, header_row=header_row)
        wb = fill(wb, schema, records)
        result = wb_to_bytes(wb)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Ошибка обработки: {exc}")

    filename = template.filename or "result.xlsx"
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"

    encoded = urllib.parse.quote(filename, safe="")
    return Response(
        content=result,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"},
    )
