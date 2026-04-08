from typing import Any, Optional

from fastapi import APIRouter
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

from .service import DOCX_MIME_TYPE, generate_iae_docx_bytes, to_base64


router = APIRouter()


class IaEDocxRequest(BaseModel):
    plural: Optional[str] = None
    singular: Optional[str] = None
    date: Optional[str] = None
    batchNumber: Optional[str] = None
    body: Optional[Any] = None
    documentName: Optional[str] = None
    returnBase64: Optional[bool] = False


@router.post("/")
def create_iae_docx(payload: IaEDocxRequest):
    base_name = (payload.documentName or "iae").strip() if payload.documentName else "iae"
    if base_name.lower().endswith(".docx"):
        base_name = base_name[:-5]

    file_name = f"{base_name}.docx"

    try:
        docx_bytes = generate_iae_docx_bytes(
            plural=payload.plural,
            singular=payload.singular,
            date=payload.date,
            batchNumber=payload.batchNumber,
            body=payload.body,
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

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
