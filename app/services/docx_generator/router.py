from typing import Any, Optional

from fastapi import APIRouter
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from .service import DOCX_MIME_TYPE, generate_docx_bytes, parse_openai_json_content, to_base64


router = APIRouter()


class DocxRequest(BaseModel):
    city: Optional[str] = None
    authorName: Optional[str] = None
    date: Optional[str] = None
    body: Optional[Any] = None
    rawContent: Optional[str] = None
    documentName: Optional[str] = None
    returnBase64: Optional[bool] = False


@router.post("/")
def create_docx(payload: DocxRequest):
    resolved_body: Any = payload.body

    if resolved_body is None and payload.rawContent:
        try:
            resolved_body = parse_openai_json_content(payload.rawContent)
        except ValueError as e:
            return JSONResponse(status_code=400, content={"success": False, "error": str(e)})

    base_name = (payload.documentName or "document").strip() if payload.documentName else "document"
    if base_name.lower().endswith(".docx"):
        base_name = base_name[:-5]

    file_name = f"{base_name}.docx"

    docx_bytes = generate_docx_bytes(
        city=payload.city,
        author_name=payload.authorName,
        date=payload.date,
        body=resolved_body,
    )

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
