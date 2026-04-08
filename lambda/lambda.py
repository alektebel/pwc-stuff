import base64
import json
import os
import traceback

import boto3
from botocore.exceptions import ClientError

from document import DocumentProcessor

# ── Config (override via Lambda environment variables) ────────────────────────
AWS_REGION      = os.environ.get("AWS_REGION",      "eu-west-1")
DYNAMODB_TABLE  = os.environ.get("DYNAMODB_TABLE",  "auditoria-generacion-dev")
S3_BUCKET       = os.environ.get("S3_BUCKET",       "auditoria-context-dev")
TEMPLATE_S3_KEY = os.environ.get("TEMPLATE_S3_KEY", "templates/plantilla.docx")

PK_ATTR      = "report_id"
SK_ATTR      = "report_sort"
CONTENT_ATTR = "validated_content"

REPORT_SECTIONS = [
    "1. Alcance",
    "2. Valoracion",
    "3. Conclusiones",
    "4. Propuestas",
]

CORS = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
    "Access-Control-Allow-Methods": "OPTIONS,POST",
}

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ok(report_id: str, doc_bytes: bytes) -> dict:
    return {
        "statusCode": 200,
        "headers": {
            "Content-Type": DOCX_MIME,
            "Content-Disposition": f'attachment; filename="report_{report_id}.docx"',
            **CORS,
        },
        "body": base64.b64encode(doc_bytes).decode("utf-8"),
        "isBase64Encoded": True,
    }


def _err(status: int, message: str) -> dict:
    return {
        "statusCode": status,
        "headers": {"Content-Type": "application/json", **CORS},
        "body": json.dumps({"error": message}),
        "isBase64Encoded": False,
    }


# ── DynamoDB ──────────────────────────────────────────────────────────────────

def _fetch_sections(report_id: str) -> list:
    """
    Fetch the 4 report sections for *report_id* in REPORT_SECTIONS order.
    Raises ValueError if any section is missing or has empty validated_content.
    """
    table = boto3.resource("dynamodb", region_name=AWS_REGION).Table(DYNAMODB_TABLE)

    sections, empty = [], []
    for sort_key in REPORT_SECTIONS:
        try:
            resp = table.get_item(Key={PK_ATTR: report_id, SK_ATTR: sort_key})
        except ClientError as exc:
            raise RuntimeError(
                f"DynamoDB error on '{sort_key}': {exc.response['Error']['Message']}"
            ) from exc

        item = resp.get("Item")
        if item is None:
            raise ValueError(
                f"Section not found — {PK_ATTR}='{report_id}', {SK_ATTR}='{sort_key}'"
            )

        content = (item.get(CONTENT_ATTR) or "").strip()
        if not content:
            empty.append(sort_key)
        else:
            sections.append(content)

    if empty:
        raise ValueError(
            f"Sections with empty '{CONTENT_ATTR}': {', '.join(empty)}"
        )
    return sections


# ── S3 ────────────────────────────────────────────────────────────────────────

def _fetch_template() -> bytes:
    try:
        resp = boto3.client("s3", region_name=AWS_REGION).get_object(
            Bucket=S3_BUCKET, Key=TEMPLATE_S3_KEY
        )
        return resp["Body"].read()
    except ClientError as exc:
        raise RuntimeError(
            f"Cannot fetch template s3://{S3_BUCKET}/{TEMPLATE_S3_KEY}: "
            f"{exc.response['Error']['Message']}"
        ) from exc


# ── Handler ───────────────────────────────────────────────────────────────────

def lambda_handler(event: dict, context=None) -> dict:
    method = (event.get("httpMethod") or "POST").upper()

    if method == "OPTIONS":
        return {"statusCode": 204, "headers": CORS, "body": ""}
    if method != "POST":
        return _err(405, "Method not allowed")

    raw = event.get("body") or "{}"
    if event.get("isBase64Encoded"):
        raw = base64.b64decode(raw).decode("utf-8")

    try:
        body = json.loads(raw) if isinstance(raw, str) else raw
    except json.JSONDecodeError as exc:
        return _err(400, f"Invalid JSON: {exc}")

    report_id = (body.get("report_id") or "").strip()
    if not report_id:
        return _err(400, "Missing field: 'report_id'")

    # Cover-page metadata (required)
    for field in ("audit_code", "audit_title", "uai"):
        if not (body.get(field) or "").strip():
            return _err(400, f"Missing field: '{field}'")

    cover_fields = {
        "audit_code":   body["audit_code"].strip(),
        "audit_title":  body["audit_title"].strip(),
        "uai":          body["uai"].strip(),
        "date":         (body.get("date") or "").strip(),
        "recipients":   body.get("recipients") or [],
        "audit_status": (body.get("audit_status") or "BORRADOR").strip().upper(),
    }

    try:
        sections = _fetch_sections(report_id)
    except ValueError as exc:
        return _err(422, str(exc))
    except RuntimeError as exc:
        traceback.print_exc()
        return _err(502, str(exc))

    try:
        template_bytes = _fetch_template()
    except RuntimeError as exc:
        traceback.print_exc()
        return _err(502, str(exc))

    try:
        doc_bytes = DocumentProcessor().process(template_bytes, sections, fields=cover_fields)
    except Exception as exc:
        traceback.print_exc()
        return _err(500, f"Document generation failed: {exc}")

    return _ok(report_id, doc_bytes)
