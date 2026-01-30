import base64
import io
import json
from typing import Any, Dict, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def parse_openai_json_content(raw_content: str) -> Any:
    if not isinstance(raw_content, str):
        raise ValueError("rawContent must be a string")

    cleaned = raw_content

    if "```" in cleaned:
        cleaned = cleaned.replace("```json", "").replace("```", "").strip()

    cleaned = cleaned.lstrip()
    if cleaned.lower().startswith("json"):
        cleaned = cleaned[4:].lstrip()

    try:
        return json.loads(cleaned)
    except Exception as e:
        raise ValueError(f"Failed to parse JSON: {str(e)}")


def normalize_body_to_string(body: Any) -> str:
    if body is None:
        return ""
    if isinstance(body, str):
        return body
    if isinstance(body, list):
        parts = []
        for item in body:
            if isinstance(item, str):
                parts.append(item)
            else:
                parts.append(json.dumps(item, ensure_ascii=False))
        return "\n\n".join(parts)
    if isinstance(body, dict):
        try:
            return json.dumps(body, ensure_ascii=False, indent=2)
        except Exception:
            return str(body)
    return str(body)


def _set_times_new_roman_12pt(run):
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)


def generate_docx_bytes(*, city: Optional[str], author_name: Optional[str], date: Optional[str], body: Any) -> bytes:
    safe_city = city or "City"
    safe_author = author_name or "Name"
    safe_date = date or "No Date"

    title = f"{safe_city} – {safe_author}"
    body_text = normalize_body_to_string(body)

    doc = Document()

    # Title paragraph
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(title)
    r_title.bold = True
    _set_times_new_roman_12pt(r_title)

    # Body paragraphs (split by line breaks)
    full_text = f"{safe_date} – {body_text}" if body_text else f"{safe_date} – "
    lines = str(full_text).splitlines() or [""]

    for line in lines:
        p = doc.add_paragraph()
        # python-docx justification support varies; this is the closest setting.
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(line)
        _set_times_new_roman_12pt(r)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def to_base64(data: bytes) -> str:
    return base64.b64encode(data).decode("utf-8")
