from typing import Any, Optional

from fastapi import APIRouter
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from .service import DOCX_MIME_TYPE, generate_diebotschaft_docx_bytes, to_base64


router = APIRouter()


class DieBotschaftDocxRequest(BaseModel):
    batchNumber: Optional[str] = None
    state: Optional[str] = None
    churchDistrict: Optional[str] = None
    author: Optional[str] = None
    date: Optional[str] = None
    body: Optional[Any] = None
    documentName: Optional[str] = None
    returnBase64: Optional[bool] = False


@router.post("/")
def create_diebotschaft_docx(payload: DieBotschaftDocxRequest):
    base_name = (payload.documentName or "diebotschaft").strip() if payload.documentName else "diebotschaft"
    if base_name.lower().endswith(".docx"):
        base_name = base_name[:-5]

    file_name = f"{base_name}.docx"

    try:
        docx_bytes = generate_diebotschaft_docx_bytes(
            batchNumber=payload.batchNumber,
            state=payload.state,
            churchDistrict=payload.churchDistrict,
            author=payload.author,
            date=payload.date,
            body=payload.body,
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
