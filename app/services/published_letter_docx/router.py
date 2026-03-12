from typing import Any, Optional

from fastapi import APIRouter
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from .service import DOCX_MIME_TYPE, generate_published_letter_docx_bytes, to_base64


router = APIRouter()


class PublishedLetterDocxRequest(BaseModel):
    letterHeading: Optional[str] = None
    letterSubheading: Optional[str] = None
    date: Optional[str] = None
    body: Optional[Any] = None
    authorName: Optional[str] = None
    documentName: Optional[str] = None
    returnBase64: Optional[bool] = False


@router.post("/")
def create_published_letter_docx(payload: PublishedLetterDocxRequest):
    base_name = (payload.documentName or "published-letter").strip() if payload.documentName else "published-letter"
    if base_name.lower().endswith(".docx"):
        base_name = base_name[:-5]

    file_name = f"{base_name}.docx"

    try:
        docx_bytes = generate_published_letter_docx_bytes(
            letter_heading=payload.letterHeading,
            letter_subheading=payload.letterSubheading,
            date=payload.date,
            body=payload.body,
            author_name=payload.authorName,
        )
    except ValueError as e:
        return JSONResponse(status_code=400, content={"success": False, "error": str(e)})

    if payload.returnBase64:
        return {
            "success": True,
            "fileName": file_name,
            "mimeType": DOCX_MIME_TYPE,
            "data": to_base64(docx_bytes),
        }

    return Response(
        content=docx_bytes,
        media_type=DOCX_MIME_TYPE,
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )
