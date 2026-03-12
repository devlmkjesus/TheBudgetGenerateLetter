from fastapi import APIRouter

from app.services.docx_generator.router import router as docx_generator_router
from app.services.published_letter_docx.router import router as published_letter_docx_router

api_router = APIRouter()

api_router.include_router(docx_generator_router, prefix="/docx-generator", tags=["DOCX Generator"])
api_router.include_router(published_letter_docx_router, prefix="/published-letter-docx", tags=["Published Letter DOCX"])
