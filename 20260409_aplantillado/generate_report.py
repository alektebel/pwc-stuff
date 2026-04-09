"""
Generate a .docx report from a Markdown file using plantilla-corrected-copy-repaired.docx
as the template.

Strategy:
  1. Use pandoc to convert the markdown → .docx, with the template as --reference-doc
     so that all heading / paragraph / list styles come from the template.
  2. Open the repaired template (to keep its cover page, TOC area, headers/footers,
     section properties and page layout).
  3. Splice the pandoc body elements into the template, replacing the instruction
     paragraphs that start at the first Heading 1 ("Alcance").
  4. Remap pandoc's auto-generated list numIds to the template's actual numIds so
     bullets and numbered lists render correctly.
  5. Save the result as `informe-generado.docx`.

Usage:
    python generate_report.py [markdown_file] [output_file]

Defaults:
    markdown_file  → ../backend/sample.md
    output_file    → informe-generado.docx
"""

import os
import sys
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

SCRIPT_DIR  = Path(__file__).parent
TEMPLATE    = SCRIPT_DIR / "plantilla-corrected-copy-repaired.docx"
DEFAULT_MD  = SCRIPT_DIR.parent / "backend" / "sample.md"
DEFAULT_OUT = SCRIPT_DIR / "informe-generado.docx"

# numIds from the template's numbering.xml (verified by inspection)
TMPL_BULLET_NUM_ID  = "1"   # abstractNumId=0, format=bullet
TMPL_DECIMAL_NUM_ID = "3"   # abstractNumId=34, format=decimal


def _find_content_start(body_children: list) -> int:
    """Return index of the first element AFTER all leading section breaks (cover + TOC)."""
    last_sectpr_para = -1
    for i, child in enumerate(body_children):
        if child.tag == qn("w:p"):
            pPr = child.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                last_sectpr_para = i
    return last_sectpr_para + 1


def _pandoc_convert(md_path: Path, reference_docx: Path) -> Path:
    """Run pandoc and return path to the generated .docx (caller must delete it)."""
    out = Path(tempfile.mktemp(suffix=".docx"))
    cmd = [
        "pandoc", str(md_path),
        "-o", str(out),
        "--reference-doc", str(reference_docx),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"pandoc failed:\n{result.stderr}")
    return out


def _classify_pandoc_numids(pandoc_docx_path: Path) -> dict[str, str]:
    """
    Read pandoc's numbering.xml and return a mapping
    {pandoc_numId → template_numId} based on the list format (bullet or decimal).
    """
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    mapping: dict[str, str] = {}

    with zipfile.ZipFile(pandoc_docx_path, "r") as z:
        if "word/numbering.xml" not in z.namelist():
            return mapping
        root = etree.fromstring(z.read("word/numbering.xml"))

    # Build abstractNumId → format map
    abstract_fmt: dict[str, str] = {}
    for a in root.findall(f"{{{ns}}}abstractNum"):
        abs_id = a.get(f"{{{ns}}}abstractNumId")
        lvl0 = a.find(f"{{{ns}}}lvl[@{{{ns}}}ilvl='0']")
        if lvl0 is not None:
            fmt_el = lvl0.find(f"{{{ns}}}numFmt")
            if fmt_el is not None:
                abstract_fmt[abs_id] = fmt_el.get(f"{{{ns}}}val", "bullet")

    # Map each pandoc numId to a template numId
    for n in root.findall(f"{{{ns}}}num"):
        num_id = n.get(f"{{{ns}}}numId")
        abs_ref = n.find(f"{{{ns}}}abstractNumId")
        if abs_ref is None:
            continue
        abs_val = abs_ref.get(f"{{{ns}}}val")
        fmt = abstract_fmt.get(abs_val, "bullet")
        tmpl_id = TMPL_DECIMAL_NUM_ID if fmt == "decimal" else TMPL_BULLET_NUM_ID
        mapping[num_id] = tmpl_id

    return mapping


def _remap_numids(body_elements: list, numid_map: dict[str, str]) -> None:
    """
    Walk all w:numId elements in *body_elements* and replace their val
    attribute according to *numid_map*.
    """
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for elem in body_elements:
        for numid_el in elem.iter(f"{{{ns}}}numId"):
            old = numid_el.get(f"{{{ns}}}val")
            if old in numid_map:
                numid_el.set(f"{{{ns}}}val", numid_map[old])


def generate(md_path: Path, template_path: Path, output_path: Path) -> None:
    print(f"Markdown  : {md_path}")
    print(f"Template  : {template_path}")
    print(f"Output    : {output_path}")

    # ── Step 1: convert markdown → styled docx via pandoc ─────────────────────
    print("Running pandoc …")
    pandoc_out = _pandoc_convert(md_path, template_path)
    try:
        numid_map = _classify_pandoc_numids(pandoc_out)
        print(f"List numId remapping: {numid_map}")
        pandoc_doc = Document(str(pandoc_out))
        pandoc_body_elements = list(pandoc_doc.element.body)
    finally:
        pandoc_out.unlink(missing_ok=True)

    # ── Step 2: open template, locate where to splice ─────────────────────────
    tmpl = Document(str(template_path))
    tmpl_body = tmpl.element.body
    children = list(tmpl_body)

    content_start = _find_content_start(children)
    print(f"Content starts at element index {content_start} (first Heading 1 / main body)")

    final_sectpr_idx = next(
        (i for i, c in enumerate(children) if c.tag == qn("w:sectPr")),
        len(children) - 1,
    )

    # ── Step 3: remove instruction paragraphs [content_start .. final_sectpr) ─
    for elem in children[content_start:final_sectpr_idx]:
        tmpl_body.remove(elem)

    # ── Step 4: remap pandoc numIds → template numIds in inserted elements ─────
    insertable = [e for e in pandoc_body_elements if e.tag != qn("w:sectPr")]
    if numid_map:
        _remap_numids(insertable, numid_map)

    # ── Step 5: insert pandoc content before final sectPr ─────────────────────
    children_after = list(tmpl_body)
    final_sectpr_idx_new = next(
        i for i, c in enumerate(children_after) if c.tag == qn("w:sectPr")
    )
    for i, elem in enumerate(insertable):
        tmpl_body.insert(final_sectpr_idx_new + i, elem)

    # ── Step 6: save ───────────────────────────────────────────────────────────
    tmpl.save(str(output_path))
    print(f"Saved → {output_path}  ({output_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    md   = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_MD
    out  = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_OUT

    if not md.exists():
        sys.exit(f"ERROR: Markdown file not found: {md}")
    if not TEMPLATE.exists():
        sys.exit(f"ERROR: Template not found: {TEMPLATE}")

    generate(md, TEMPLATE, out)
