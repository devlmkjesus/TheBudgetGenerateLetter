import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import API_HOST, API_PORT, API_RELOAD, CORS_ORIGINS
from app.services.docx_generator.router import router as docx_generator_router

app = FastAPI(
    title="DOCX Generator API",
    description="API for generating DOCX documents",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def read_root():
    return {
        "message": "Welcome to the DOCX Generator API",
        "services": [{"name": "DOCX Generator", "endpoint": "/docx-generator"}],
        "documentation": "/docs",
    }


@app.get("/health")
def health():
    return {"ok": True}


app.include_router(docx_generator_router, prefix="/docx-generator", tags=["DOCX Generator"])


if __name__ == "__main__":
    uvicorn.run("main:app", host=API_HOST, port=API_PORT, reload=API_RELOAD)
