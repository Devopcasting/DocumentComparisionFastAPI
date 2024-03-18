from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import uvicorn

from app.v1.endpoints.excel_endpoint import router as v1_excel_endpoint
from app.v1.endpoints.doc_endpoint import router as v1_docx_endpoint

app = FastAPI()
app.mount("/static", StaticFiles(directory=r"app\v1\static"), name="static")
app.include_router(v1_excel_endpoint, prefix="/v1")
app.include_router(v1_docx_endpoint, prefix="/v1")

if __name__ == '__main__':
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
