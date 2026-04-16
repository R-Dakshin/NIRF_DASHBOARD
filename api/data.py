import os
from fastapi import FastAPI
from fastapi.responses import JSONResponse

from sync_dashboard import build_data

app = FastAPI()


@app.get("/api/data")
def get_data():
    base_dir = os.path.dirname(os.path.dirname(__file__))
    excel_file = os.path.join(base_dir, "data.xlsx")
    try:
        payload = build_data(excel_file)
        return JSONResponse(payload)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
