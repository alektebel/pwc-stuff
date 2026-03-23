/**
 * API service — mirrors an AWS API Gateway + Lambda architecture.
 *
 * All calls go through /api (proxied to the local Flask server that
 * emulates API Gateway → Lambda → DynamoDB).
 */

const BASE = "/api";

async function _handleResponse(res) {
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try {
      const json = await res.json();
      msg = json.error || msg;
    } catch {
      // non-JSON error body
    }
    throw new Error(msg);
  }
  return res;
}

/**
 * POST /api/process
 *
 * @param {File}   templateFile  – .docx Word template
 * @param {string} markdownText  – markdown string
 * @param {function} onProgress  – optional progress callback (0–100)
 * @returns {{ documentId, filename, createdAt, outputSize }}
 */
export async function processDocument(templateFile, markdownText, onProgress) {
  const form = new FormData();
  form.append("template", templateFile);
  form.append("markdown", markdownText);

  // Use XMLHttpRequest for upload progress tracking
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();

    xhr.upload.addEventListener("progress", (e) => {
      if (e.lengthComputable && onProgress) {
        onProgress(Math.round((e.loaded / e.total) * 80)); // 0–80% for upload
      }
    });

    xhr.addEventListener("load", () => {
      if (onProgress) onProgress(100);
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          resolve(JSON.parse(xhr.responseText));
        } catch {
          reject(new Error("Invalid JSON response from server"));
        }
      } else {
        let msg = `HTTP ${xhr.status}`;
        try {
          const json = JSON.parse(xhr.responseText);
          msg = json.error || msg;
        } catch {
          // ignore
        }
        reject(new Error(msg));
      }
    });

    xhr.addEventListener("error", () => reject(new Error("Network error")));
    xhr.addEventListener("abort", () => reject(new Error("Request aborted")));

    xhr.open("POST", `${BASE}/process`);
    xhr.send(form);
  });
}

/**
 * GET /api/documents
 * @returns {{ documents: Array, count: number }}
 */
export async function listDocuments(limit = 50) {
  const res = await fetch(`${BASE}/documents?limit=${limit}`);
  await _handleResponse(res);
  return res.json();
}

/**
 * GET /api/documents/:id
 * @returns {Object} document metadata
 */
export async function getDocument(documentId) {
  const res = await fetch(`${BASE}/documents/${documentId}`);
  await _handleResponse(res);
  return res.json();
}

/**
 * GET /api/documents/:id/download
 * Triggers a browser file download.
 */
export async function downloadDocument(documentId, filename) {
  const res = await fetch(`${BASE}/documents/${documentId}/download`);
  await _handleResponse(res);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || `document_${documentId}.docx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/**
 * DELETE /api/documents/:id
 */
export async function deleteDocument(documentId) {
  const res = await fetch(`${BASE}/documents/${documentId}`, {
    method: "DELETE",
  });
  await _handleResponse(res);
  return res.json();
}

/**
 * GET /api/health
 */
export async function healthCheck() {
  const res = await fetch(`${BASE}/health`);
  await _handleResponse(res);
  return res.json();
}
