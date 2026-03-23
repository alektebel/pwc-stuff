import React, { useState, useCallback } from "react";
import TemplateUploader from "./components/TemplateUploader.jsx";
import MarkdownEditor from "./components/MarkdownEditor.jsx";
import DocumentHistory from "./components/DocumentHistory.jsx";
import { processDocument, downloadDocument } from "./services/api.js";

// ---------------------------------------------------------------------------
// Progress bar
// ---------------------------------------------------------------------------
function ProgressBar({ value }) {
  return (
    <div className="progress-bar-track" role="progressbar" aria-valuenow={value}>
      <div className="progress-bar-fill" style={{ width: `${value}%` }} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------
export default function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [markdown, setMarkdown]         = useState("");
  const [status, setStatus]             = useState("idle"); // idle | processing | success | error
  const [progress, setProgress]         = useState(0);
  const [errorMsg, setErrorMsg]         = useState("");
  const [lastResult, setLastResult]     = useState(null);
  const [historyKey, setHistoryKey]     = useState(0);

  const canSubmit =
    templateFile !== null && markdown.trim().length > 0 && status !== "processing";

  const handleProcess = useCallback(async () => {
    if (!canSubmit) return;

    setStatus("processing");
    setProgress(0);
    setErrorMsg("");
    setLastResult(null);

    try {
      const result = await processDocument(templateFile, markdown, setProgress);
      setLastResult(result);
      setStatus("success");
      setHistoryKey((k) => k + 1); // refresh history panel
    } catch (err) {
      setStatus("error");
      setErrorMsg(err.message);
    }
  }, [canSubmit, templateFile, markdown]);

  async function handleDownloadLast() {
    if (!lastResult) return;
    try {
      await downloadDocument(lastResult.documentId, lastResult.filename);
    } catch (err) {
      alert(`Download failed: ${err.message}`);
    }
  }

  return (
    <div className="app-shell">
      {/* ---- Top bar ---- */}
      <header className="top-bar">
        <div className="top-bar-brand">
          <span className="brand-icon">📝</span>
          <span className="brand-name">Doc Transformer</span>
        </div>
        <span className="top-bar-sub">
          Markdown → Word · AWS Lambda (emulated)
        </span>
      </header>

      <div className="app-body">
        {/* ================================================================
            LEFT PANEL — Inputs
        ================================================================= */}
        <main className="input-panel">
          <section className="card">
            <h2 className="card-title">1. Upload Word Template</h2>
            <p className="card-hint">
              The processor extracts styles, page layout, header/footer, and
              images from this template, then applies them to your markdown.
            </p>
            <TemplateUploader file={templateFile} onChange={setTemplateFile} />
          </section>

          <section className="card">
            <h2 className="card-title">2. Write or Paste Markdown</h2>
            <MarkdownEditor value={markdown} onChange={setMarkdown} />
          </section>

          {/* ---- Generate button ---- */}
          <div className="generate-row">
            <button
              className="btn-primary btn-generate"
              disabled={!canSubmit}
              onClick={handleProcess}
            >
              {status === "processing" ? "Generating…" : "Generate Document"}
            </button>

            {status === "processing" && (
              <ProgressBar value={progress} />
            )}

            {status === "success" && lastResult && (
              <div className="result-banner result-success">
                <span>Document ready — {(lastResult.outputSize / 1024).toFixed(1)} KB</span>
                <button className="btn-primary" onClick={handleDownloadLast}>
                  ⬇ Download
                </button>
              </div>
            )}

            {status === "error" && (
              <div className="result-banner result-error">
                <span>⚠ {errorMsg}</span>
                <button className="btn-ghost" onClick={() => setStatus("idle")}>
                  Dismiss
                </button>
              </div>
            )}
          </div>
        </main>

        {/* ================================================================
            RIGHT PANEL — History
        ================================================================= */}
        <DocumentHistory refreshTrigger={historyKey} />
      </div>
    </div>
  );
}
