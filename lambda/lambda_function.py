"""
AWS Lambda handler — entry point is lambda_handler(event, context).

Receives a POST from the React app via API Gateway:
  { "markdown": "..." }

Fetches the .docx template from S3, processes the markdown,
and returns the generated document as base64.
"""

import base64
import json
import os
import traceback
import uuid
from datetime import datetime, timezone

import boto3
from document_processor import DocumentProcessor

_CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
    "Access-Control-Allow-Methods": "OPTIONS,POST",
}

S3_BUCKET = os.environ["S3_BUCKET_NAME"]
TEMPLATE_KEY = os.environ.get("TEMPLATE_S3_KEY", "templates/corporate_template.docx")


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


def _fetch_template() -> bytes:
    s3 = boto3.client("s3")
    response = s3.get_object(Bucket=S3_BUCKET, Key=TEMPLATE_KEY)
    return response["Body"].read()


def lambda_handler(event: dict, context=None) -> dict:
    http_method = (event.get("httpMethod") or "POST").upper()

    if http_method == "OPTIONS":
        return {"statusCode": 204, "headers": _CORS, "body": ""}

    if http_method != "POST":
        return _err(405, "Method not allowed")

    raw_body = event.get("body") or "{}"
    if event.get("isBase64Encoded"):
        raw_body = base64.b64decode(raw_body).decode("utf-8")

    try:
        body = json.loads(raw_body) if isinstance(raw_body, str) else raw_body
    except json.JSONDecodeError as exc:
        return _err(400, f"Invalid JSON body: {exc}")

    markdown_text: str = body.get("markdown", "").strip()
    if not markdown_text:
        return _err(400, "Missing field: 'markdown'")

    try:
        template_bytes = _fetch_template()
    except Exception as exc:
        traceback.print_exc()
        return _err(500, f"Could not fetch template from S3: {exc}")

    try:
        processor = DocumentProcessor()
        output_bytes = processor.process(template_bytes, markdown_text)
    except Exception as exc:
        traceback.print_exc()
        return _err(500, f"Document processing failed: {exc}")

    document_id = str(uuid.uuid4())
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output_filename = f"output_{timestamp}.docx"

    try:
        s3 = boto3.client("s3")
        s3.put_object(
            Bucket=S3_BUCKET,
            Key=f"outputs/{output_filename}",
            Body=output_bytes,
            ContentType=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
    except Exception:
        traceback.print_exc()

    return _ok({
        "documentId": document_id,
        "filename": output_filename,
        "createdAt": datetime.now(timezone.utc).isoformat(),
        "document": base64.b64encode(output_bytes).decode("utf-8"),
    })
