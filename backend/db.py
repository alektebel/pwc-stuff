"""
DynamoDB emulation using a local JSON file.

Table schema (single table):
  PK  = document_id  (str, uuid)
  SK  = "DOCUMENT"   (constant, mirrors a single-table DynamoDB pattern)

  Attributes:
    document_id    str
    template_name  str
    created_at     str  (ISO-8601 UTC)
    output_path    str  (relative path inside uploads/outputs/)
    markdown_size  int  (bytes)
    output_size    int  (bytes)
    status         str  ("completed" | "failed")
    error          str  (only on failed)
"""

import json
import os
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

_DB_PATH = Path(__file__).parent / "data" / "db.json"
_lock = threading.Lock()


def _load() -> dict:
    if not _DB_PATH.exists():
        return {"items": []}
    with open(_DB_PATH, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _save(data: dict) -> None:
    _DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(_DB_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def put_item(item: dict) -> None:
    """Insert or replace an item (upsert by document_id)."""
    with _lock:
        data = _load()
        # Replace if exists
        data["items"] = [
            i for i in data["items"] if i.get("document_id") != item["document_id"]
        ]
        item.setdefault(
            "created_at", datetime.now(timezone.utc).isoformat()
        )
        data["items"].append(item)
        _save(data)


def get_item(document_id: str) -> Optional[dict]:
    """Return a single item by document_id, or None."""
    with _lock:
        data = _load()
    for item in data["items"]:
        if item.get("document_id") == document_id:
            return item
    return None


def list_items(limit: int = 50) -> list:
    """Return up to *limit* items, newest first."""
    with _lock:
        data = _load()
    items = sorted(
        data["items"],
        key=lambda i: i.get("created_at", ""),
        reverse=True,
    )
    return items[:limit]


def delete_item(document_id: str) -> bool:
    """Delete an item.  Returns True if it existed."""
    with _lock:
        data = _load()
        before = len(data["items"])
        data["items"] = [
            i for i in data["items"] if i.get("document_id") != document_id
        ]
        existed = len(data["items"]) < before
        if existed:
            _save(data)
    return existed
