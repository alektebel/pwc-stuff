"""
Local server that emulates:
  - AWS API Gateway (routing, CORS, binary responses)
  - AWS Lambda invocation (calls lambda_handler directly)
  - AWS DynamoDB (via db.py — JSON file on disk)
  - AWS S3 (stores files in uploads/ directory)

Endpoints
---------
POST   /api/process                 Convert markdown + template → .docx
GET    /api/documents               List generated documents
GET    /api/documents/<id>          Get document metadata
GET    /api/documents/<id>/download Download generated .docx
DELETE /api/documents/<id>          Delete a document record + file
"""

import base64
import json
import os
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, request, send_file, abort
from flask_cors import CORS

# Make sure local modules are importable when running from any directory
sys.path.insert(0, str(Path(__file__).parent))

import db
import lambda_handler as lh

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
app = Flask(__name__)
CORS(app)

BASE_DIR = Path(__file__).parent
OUTPUTS_DIR = BASE_DIR / "uploads" / "outputs"
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

MAX_TEMPLATE_MB = 20
MAX_MARKDOWN_MB = 5


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _json_error(status: int, message: str):
    return jsonify({"error": message}), status


def _save_output(doc_bytes: bytes, filename: str) -> Path:
    path = OUTPUTS_DIR / filename
    path.write_bytes(doc_bytes)
    return path


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/api/process", methods=["POST", "OPTIONS"])
def process():
    if request.method == "OPTIONS":
        return "", 204

    # ---- Validate content-type -----------------------------------------
    if not request.content_type or "multipart/form-data" not in request.content_type:
        return _json_error(415, "Content-Type must be multipart/form-data")

    # ---- Read template file --------------------------------------------
    template_file = request.files.get("template")
    if not template_file:
        return _json_error(400, "Missing file field: 'template'")

    template_bytes = template_file.read()
    if len(template_bytes) > MAX_TEMPLATE_MB * 1024 * 1024:
        return _json_error(413, f"Template exceeds {MAX_TEMPLATE_MB} MB")

    template_name = template_file.filename or "template.docx"

    # ---- Read markdown text --------------------------------------------
    markdown_text = request.form.get("markdown", "").strip()
    if not markdown_text:
        return _json_error(400, "Missing form field: 'markdown'")

    if len(markdown_text.encode()) > MAX_MARKDOWN_MB * 1024 * 1024:
        return _json_error(413, f"Markdown exceeds {MAX_MARKDOWN_MB} MB")

    # ---- Build Lambda event (API Gateway proxy format) -----------------
    event = {
        "httpMethod": "POST",
        "body": json.dumps(
            {
                "template": base64.b64encode(template_bytes).decode(),
                "markdown": markdown_text,
                "templateName": template_name,
            }
        ),
        "isBase64Encoded": False,
    }

    # ---- Invoke Lambda -------------------------------------------------
    response = lh.lambda_handler(event, context=None)
    status_code = response.get("statusCode", 500)
    body = json.loads(response.get("body", "{}"))

    if status_code != 200:
        return _json_error(status_code, body.get("error", "Unknown error"))

    # ---- Save output file ("S3") ---------------------------------------
    doc_b64 = body.get("document", "")
    output_bytes = base64.b64decode(doc_b64)
    filename = body.get("filename", f"output_{uuid.uuid4()}.docx")
    _save_output(output_bytes, filename)

    # ---- Persist metadata ("DynamoDB") ---------------------------------
    document_id = body.get("documentId", str(uuid.uuid4()))
    record = {
        "document_id": document_id,
        "template_name": template_name,
        "created_at": body.get("createdAt", datetime.now(timezone.utc).isoformat()),
        "output_path": filename,
        "markdown_preview": markdown_text[:300],
        "markdown_size": len(markdown_text.encode()),
        "output_size": body.get("outputSizeBytes", len(output_bytes)),
        "template_size": body.get("templateSizeBytes", len(template_bytes)),
        "status": "completed",
    }
    db.put_item(record)

    return jsonify(
        {
            "documentId": document_id,
            "filename": filename,
            "createdAt": record["created_at"],
            "outputSize": record["output_size"],
        }
    ), 201


@app.route("/api/documents", methods=["GET"])
def list_documents():
    items = db.list_items(limit=int(request.args.get("limit", 50)))
    # Strip the heavy fields before returning the list
    slim = [
        {k: v for k, v in item.items() if k != "output_path"}
        for item in items
    ]
    return jsonify({"documents": slim, "count": len(slim)})


@app.route("/api/documents/<document_id>", methods=["GET"])
def get_document(document_id: str):
    item = db.get_item(document_id)
    if not item:
        return _json_error(404, "Document not found")
    return jsonify({k: v for k, v in item.items() if k != "output_path"})


@app.route("/api/documents/<document_id>/download", methods=["GET"])
def download_document(document_id: str):
    item = db.get_item(document_id)
    if not item:
        return _json_error(404, "Document not found")

    output_path = OUTPUTS_DIR / item["output_path"]
    if not output_path.exists():
        return _json_error(404, "Output file missing on server")

    return send_file(
        output_path,
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
        as_attachment=True,
        download_name=item["output_path"],
    )


@app.route("/api/documents/<document_id>", methods=["DELETE"])
def delete_document(document_id: str):
    item = db.get_item(document_id)
    if not item:
        return _json_error(404, "Document not found")

    # Remove file
    output_path = OUTPUTS_DIR / item["output_path"]
    if output_path.exists():
        output_path.unlink()

    db.delete_item(document_id)
    return jsonify({"deleted": document_id}), 200


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "doc-transformer"})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"  Doc-Transformer API  →  http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=True)
