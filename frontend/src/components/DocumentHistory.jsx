import React, { useEffect, useState, useCallback } from "react";
import { listDocuments, downloadDocument, deleteDocument } from "../services/api.js";

function formatBytes(bytes) {
  if (!bytes) return "—";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatDate(iso) {
  if (!iso) return "—";
  try {
    return new Intl.DateTimeFormat(undefined, {
      dateStyle: "medium",
      timeStyle: "short",
    }).format(new Date(iso));
  } catch {
    return iso;
  }
}

export default function DocumentHistory({ refreshTrigger }) {
  const [documents, setDocuments] = useState([]);
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState(null);
  const [deleting, setDeleting] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await listDocuments();
      setDocuments(data.documents || []);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { load(); }, [load, refreshTrigger]);

  async function handleDownload(doc) {
    try {
      await downloadDocument(doc.document_id, doc.output_path || `${doc.document_id}.docx`);
    } catch (err) {
      alert(`Download failed: ${err.message}`);
    }
  }

  async function handleDelete(doc) {
    if (!confirm(`Delete document "${doc.template_name}"?`)) return;
    setDeleting(doc.document_id);
    try {
      await deleteDocument(doc.document_id);
      setDocuments((prev) => prev.filter((d) => d.document_id !== doc.document_id));
    } catch (err) {
      alert(`Delete failed: ${err.message}`);
    } finally {
      setDeleting(null);
    }
  }

  return (
    <aside className="history-panel">
      <div className="history-header">
        <h2 className="history-title">Generated Documents</h2>
        <button
          className="btn-ghost btn-icon"
          onClick={load}
          disabled={loading}
          title="Refresh"
        >
          {loading ? "⟳" : "↺"}
        </button>
      </div>

      {error && (
        <div className="history-error">
          <span>⚠ {error}</span>
          <button className="btn-ghost" onClick={load}>Retry</button>
        </div>
      )}

      {!loading && !error && documents.length === 0 && (
        <div className="history-empty">
          <p>No documents yet.</p>
          <p className="history-hint">Generate one to see it here.</p>
        </div>
      )}

      <ul className="history-list">
        {documents.map((doc) => (
          <li key={doc.document_id} className="history-item">
            <div className="history-item-top">
              <span className={`status-dot status-${doc.status}`} title={doc.status} />
              <span className="history-template">{doc.template_name}</span>
              <span className="history-size">{formatBytes(doc.output_size)}</span>
            </div>

            <div className="history-date">{formatDate(doc.created_at)}</div>

            {doc.markdown_preview && (
              <div className="history-preview">
                {doc.markdown_preview.slice(0, 120)}
                {doc.markdown_preview.length > 120 ? "…" : ""}
              </div>
            )}

            <div className="history-actions">
              <button
                className="btn-primary btn-sm"
                onClick={() => handleDownload(doc)}
                title="Download .docx"
              >
                ⬇ Download
              </button>
              <button
                className="btn-danger btn-sm"
                onClick={() => handleDelete(doc)}
                disabled={deleting === doc.document_id}
                title="Delete"
              >
                {deleting === doc.document_id ? "…" : "✕"}
              </button>
            </div>
          </li>
        ))}
      </ul>
    </aside>
  );
}
