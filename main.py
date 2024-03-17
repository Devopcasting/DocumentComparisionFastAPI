from fastapi import FastAPI
import uvicorn

"""import version: 1"""
from app.v1.endpoints.excel_endpoint import router as v1_excel_endpoint


app = FastAPI()

app.include_router(v1_excel_endpoint, prefix="/v1")

if __name__ == '__main__':
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)