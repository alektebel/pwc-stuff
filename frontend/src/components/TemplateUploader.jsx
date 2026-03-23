import React, { useRef, useState } from "react";

export default function TemplateUploader({ file, onChange }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);

  function handleFiles(files) {
    const f = files[0];
    if (!f) return;
    if (!f.name.endsWith(".docx")) {
      alert("Please select a .docx file.");
      return;
    }
    onChange(f);
  }

  function onDrop(e) {
    e.preventDefault();
    setDragging(false);
    handleFiles(e.dataTransfer.files);
  }

  function onDragOver(e) {
    e.preventDefault();
    setDragging(true);
  }

  return (
    <div className="uploader-wrapper">
      <label className="uploader-label">Word Template (.docx)</label>

      <div
        className={`drop-zone ${dragging ? "dragging" : ""} ${file ? "has-file" : ""}`}
        onClick={() => inputRef.current.click()}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={() => setDragging(false)}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".docx"
          style={{ display: "none" }}
          onChange={(e) => handleFiles(e.target.files)}
        />

        {file ? (
          <div className="file-info">
            <span className="file-icon">📄</span>
            <div>
              <div className="file-name">{file.name}</div>
              <div className="file-size">{(file.size / 1024).toFixed(1)} KB</div>
            </div>
            <button
              className="remove-file"
              onClick={(e) => {
                e.stopPropagation();
                onChange(null);
                inputRef.current.value = "";
              }}
            >
              ✕
            </button>
          </div>
        ) : (
          <div className="drop-placeholder">
            <span className="drop-icon">⬆</span>
            <p>Drop your template here or <strong>click to browse</strong></p>
            <p className="drop-hint">
              The document processor will extract styles, headers, footers,
              and images from this template.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
