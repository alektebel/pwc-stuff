"""
AWS Lambda handler — emulated locally.

Event shape (mirrors API Gateway proxy integration):
  {
    "body": "{\"template\": \"<base64>\", \"markdown\": \"...\", \"templateName\": \"...\"}",
    "isBase64Encoded": false,
    "httpMethod": "POST"
  }

Response shape:
  {
    "statusCode": 200,
    "headers": {...},
    "body": "{\"document\": \"<base64>\", \"filename\": \"...\", \"documentId\": \"...\"}"
  }
"""

import base64
import json
import traceback
import uuid
from datetime import datetime, timezone

from document_processor import DocumentProcessor

# ---------------------------------------------------------------------------
# CORS headers reused in every response
# ---------------------------------------------------------------------------
_CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
    "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
}


def _ok(body: dict) -> dict:
    return {
        "statusCode": 200,
        "headers": {"Content-Type": "application/json", **_CORS},
        "body": json.dumps(body),
    }


def _err(status: int, message: str) -> dict:
    return {
        "statusCode": status,
        "headers": {"Content-Type": "application/json", **_CORS},
        "body": json.dumps({"error": message}),
    }


# ---------------------------------------------------------------------------
# Handler
# ---------------------------------------------------------------------------

def lambda_handler(event: dict, context=None) -> dict:
    """
    Entry point.  *context* is optional so the function can be called
    directly in tests without a Lambda context object.
    """
    http_method = (event.get("httpMethod") or "POST").upper()

    # Pre-flight (OPTIONS)
    if http_method == "OPTIONS":
        return {"statusCode": 204, "headers": _CORS, "body": ""}

    if http_method != "POST":
        return _err(405, "Method not allowed")

    # ----------------------------------------------------------------
    # Parse body
    # ----------------------------------------------------------------
    raw_body = event.get("body") or "{}"
    if event.get("isBase64Encoded"):
        raw_body = base64.b64decode(raw_body).decode("utf-8")

    try:
        body = json.loads(raw_body) if isinstance(raw_body, str) else raw_body
    except json.JSONDecodeError as exc:
        return _err(400, f"Invalid JSON body: {exc}")

    # ----------------------------------------------------------------
    # Validate inputs
    # ----------------------------------------------------------------
    template_b64: str = body.get("template", "").strip()
    markdown_text: str = body.get("markdown", "").strip()
    template_name: str = body.get("templateName", "template.docx")

    if not template_b64:
        return _err(400, "Missing field: 'template' (base64-encoded .docx)")
    if not markdown_text:
        return _err(400, "Missing field: 'markdown'")

    # ----------------------------------------------------------------
    # Decode template
    # ----------------------------------------------------------------
    try:
        template_bytes = base64.b64decode(template_b64)
    except Exception as exc:
        return _err(400, f"Could not base64-decode template: {exc}")

    if len(template_bytes) < 4 or template_bytes[:4] != b"PK\x03\x04":
        return _err(400, "Template does not appear to be a valid .docx (ZIP) file")

    # ----------------------------------------------------------------
    # Process
    # ----------------------------------------------------------------
    try:
        processor = DocumentProcessor()
        output_bytes = processor.process(template_bytes, markdown_text)
    except Exception as exc:
        traceback.print_exc()
        return _err(500, f"Document processing failed: {exc}")

    # ----------------------------------------------------------------
    # Build response
    # ----------------------------------------------------------------
    document_id = str(uuid.uuid4())
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output_filename = f"output_{timestamp}.docx"

    return _ok(
        {
            "documentId": document_id,
            "filename": output_filename,
            "templateName": template_name,
            "createdAt": datetime.now(timezone.utc).isoformat(),
            "document": base64.b64encode(output_bytes).decode("utf-8"),
            # Byte sizes for diagnostics
            "templateSizeBytes": len(template_bytes),
            "outputSizeBytes": len(output_bytes),
        }
    )
