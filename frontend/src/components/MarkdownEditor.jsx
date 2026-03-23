import React, { useRef } from "react";

const SAMPLE_MD_PATH = "/sample_content.md";

export default function MarkdownEditor({ value, onChange }) {
  const textareaRef = useRef(null);

  function handleKeyDown(e) {
    // Tab → insert 2 spaces instead of changing focus
    if (e.key === "Tab") {
      e.preventDefault();
      const start = e.target.selectionStart;
      const end = e.target.selectionEnd;
      const newValue = value.substring(0, start) + "  " + value.substring(end);
      onChange(newValue);
      requestAnimationFrame(() => {
        e.target.selectionStart = start + 2;
        e.target.selectionEnd   = start + 2;
      });
    }
  }

  async function loadSample() {
    try {
      const res = await fetch(SAMPLE_MD_PATH);
      if (!res.ok) throw new Error("Not found");
      const text = await res.text();
      onChange(text);
    } catch {
      // Fallback inline sample
      onChange(FALLBACK_SAMPLE);
    }
  }

  const charCount = value.length;
  const lineCount = value.split("\n").length;

  return (
    <div className="editor-wrapper">
      <div className="editor-header">
        <label className="editor-label">Markdown Content</label>
        <div className="editor-actions">
          <button className="btn-ghost" onClick={loadSample} title="Load sample markdown">
            Load sample
          </button>
          <button
            className="btn-ghost"
            onClick={() => onChange("")}
            title="Clear editor"
            disabled={!value}
          >
            Clear
          </button>
        </div>
      </div>

      <textarea
        ref={textareaRef}
        className="markdown-textarea"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={handleKeyDown}
        placeholder={`# Document Title\n\n## Section Heading\n\nWrite your markdown content here…\n\n- Bullet item\n- Another item\n\n> Blockquote text\n\n\`\`\`\ncode block\n\`\`\``}
        spellCheck={false}
      />

      <div className="editor-footer">
        <span>{lineCount} lines</span>
        <span>{charCount.toLocaleString()} characters</span>
      </div>
    </div>
  );
}

const FALLBACK_SAMPLE = `# Sample Document

## Introduction

This is a **sample** markdown document. It demonstrates the various formatting options supported by the document processor.

> This is a blockquote that will use the template's Quote style.

---

## Features

### Text Formatting

You can use **bold**, *italic*, ~~strikethrough~~, and \`inline code\`.

### Lists

Unordered:
- First item
- Second item
  - Nested item
  - Another nested item
- Third item

Ordered:
1. Step one
2. Step two
3. Step three

### Code Block

\`\`\`python
def hello(name: str) -> str:
    return f"Hello, {name}!"
\`\`\`

### Table

| Name     | Role         | Location    |
|----------|:------------:|------------:|
| Alice    | Engineer     | New York    |
| Bob      | Designer     | London      |
| Carol    | Manager      | Paris       |

---

## Conclusion

The document processor maps each markdown element to the corresponding style in your Word template.
`;
