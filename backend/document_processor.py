"""
Core document processor: loads a Word template, clears its body,
then rebuilds content from parsed markdown — preserving all template
styles, page layout, headers, footers, and images.
"""

import base64
import re
import urllib.request
import zipfile
from io import BytesIO
import os
import subprocess
import tempfile
from typing import Optional

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE


# ---------------------------------------------------------------------------
# Mapfre font map — derived from the template's styles.xml definitions.
# Keys are OOXML style IDs (w:styleId); value is (ascii, hAnsi, cs).
# None covers paragraphs with no explicit pStyle (defaults to Normal).
# ---------------------------------------------------------------------------

_DISPLAY = "Mapfre Display"
_TEXT = "Mapfre Text"

# Headings 1-3: Display for all three font slots (as in styles.xml)
# Headings 4-9: Display for ascii/hAnsi; cs follows the template (Times New Roman)
# Body, lists, Normal, fallback: Text for all slots
_MAPFRE_RFONTS: dict[Optional[str], tuple[str, str, str]] = {
    "Heading1": (_DISPLAY, _DISPLAY, _DISPLAY),
    "Heading2": (_DISPLAY, _DISPLAY, _DISPLAY),
    "Heading3": (_DISPLAY, _DISPLAY, _DISPLAY),
    "Heading4": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Heading5": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Heading6": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Heading7": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Heading8": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Heading9": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Caption": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Cabecera1": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Cabecera2": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Anexo1": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "Anexo2": (_DISPLAY, _DISPLAY, "Times New Roman"),
    "BodyText": (_TEXT, _TEXT, _TEXT),
    None: (_TEXT, _TEXT, _TEXT),  # Normal (no explicit pStyle)
    "Normal": (_TEXT, _TEXT, _TEXT),
    "ListParagraph": (_TEXT, _TEXT, _TEXT),
    "ListBullet": (_TEXT, _TEXT, _TEXT),
    "ListBullet2": (_TEXT, _TEXT, _TEXT),
    "ListBullet3": (_TEXT, _TEXT, _TEXT),
    "ListNumber": (_TEXT, _TEXT, _TEXT),
    "ListNumber2": (_TEXT, _TEXT, _TEXT),
    "ListNumber3": (_TEXT, _TEXT, _TEXT),
    "IntenseQuote": (_TEXT, _TEXT, _TEXT),
    "Quote": (_TEXT, _TEXT, _TEXT),
    "NoSpacing": (_TEXT, _TEXT, _TEXT),
}


def _make_rfonts(ascii_f: str, hansi_f: str, cs_f: str):
    """Create a fresh <w:rFonts> element with the given font slots."""
    elem = OxmlElement("w:rFonts")
    elem.set(qn("w:ascii"), ascii_f)
    elem.set(qn("w:hAnsi"), hansi_f)
    elem.set(qn("w:cs"), cs_f)
    return elem


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _collect_style_names(doc: Document) -> dict[str, set]:
    paragraph_styles = {s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH}
    character_styles = {s.name for s in doc.styles if s.type == WD_STYLE_TYPE.CHARACTER}
    return {"paragraph": paragraph_styles, "character": character_styles}


def _fetch_image(src: str) -> Optional[bytes]:
    """Return raw bytes for a data-URI or http(s) image URL."""
    if src.startswith("data:"):
        m = re.match(r"data:[^;]+;base64,(.+)", src, re.DOTALL)
        if m:
            return base64.b64decode(m.group(1))
    elif src.startswith(("http://", "https://")):
        try:
            req = urllib.request.Request(src, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=5) as resp:
                return resp.read()
        except Exception:
            return None
    return None


class DocumentProcessor:
    """
    Loads a .docx template, clears its body,
    then rebuilds content from parsed markdown — preserving all template
    styles, page layout, headers, footers, and images.
    """

    def __init__(self):
        self.doc: Optional[Document] = None
        self.styles: dict[str, set] = {"paragraph": set(), "character": set()}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def process(
        self,
        template_bytes: bytes,
        markdown_text: str,
        cover_fields: Optional[dict] = None,
    ) -> bytes:
        # Step 1: Use Pandoc to convert markdown to a temporary DOCX file,
        # referencing the original template for content styling.
        with tempfile.NamedTemporaryFile(
            suffix=".docx", delete=False
        ) as temp_template_file:
            temp_template_file.write(template_bytes)
            temp_template_path = temp_template_file.name

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".md", delete=False
        ) as temp_markdown_file:
            temp_markdown_file.write(markdown_text)
            temp_markdown_path = temp_markdown_file.name

        pandoc_output_docx_path = ""
        try:
            pandoc_output_docx_path = os.path.join(
                tempfile.gettempdir(), "pandoc_content.docx"
            )
            command = [
                "pandoc",
                "-s",
                temp_markdown_path,
                "-o",
                pandoc_output_docx_path,
                "--reference-doc",
                temp_template_path,
            ]
            subprocess.run(command, check=True, capture_output=True)

            # Step 2: Load the ORIGINAL template into python-docx. This is our base document.
            self.doc = Document(BytesIO(template_bytes))
            self.styles = _collect_style_names(self.doc)

            # Step 3: Clear the existing body content of the template, preserving headers, footers, etc.
            self._clear_body()

            # Step 4: Load the Pandoc-generated DOCX as a source for its content.
            pandoc_doc = Document(pandoc_output_docx_path)

            # Step 5: Append content from the Pandoc-generated document to the template's body.
            for element in pandoc_doc.element.body:
                self.doc.element.body.append(element)

            # Step 6: Apply cover fields using python-docx
            if cover_fields:
                self._fill_cover_page(cover_fields)

            # Step 7: Stamp fonts for consistency (if needed, this might be redundant if pandoc+reference-doc is perfect)
            self._stamp_run_fonts()

            final_output_buffer = BytesIO()
            self.doc.save(final_output_buffer)
            return final_output_buffer.getvalue()

        finally:
            # Clean up temporary files
            os.remove(temp_template_path)
            os.remove(temp_markdown_path)
            if os.path.exists(pandoc_output_docx_path):
                os.remove(pandoc_output_docx_path)

    def _clear_body(self) -> None:
        """Keep the cover-page section; remove everything after it except the final sectPr."""
        body = self.doc.element.body
        children = list(body)
        cover_end = self._find_cover_end()
        self._cover_end = cover_end

        to_remove = [c for c in children[cover_end + 1 :] if c.tag != qn("w:sectPr")]
        for elem in to_remove:
            body.remove(elem)

    def _find_cover_end(self) -> int:
        """Return the index of the cover-page section-break paragraph."""
        body = self.doc.element.body
        for i, child in enumerate(body):
            if child.tag == qn("w:p"):
                pPr = child.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                    return i
        return 0

    def _stamp_run_fonts(self) -> None:
        """
        Stamp the hardcoded Mapfre rFonts (from _MAPFRE_RFONTS) onto every
        generated run that doesn't already carry an explicit <w:rFonts>.
        Runs inside the cover page are skipped.
        Code runs (Courier New) are also skipped.
        """
        body = self.doc.element.body
        children = list(body)
        cover_end = self._find_cover_end()

        for child in children[cover_end + 1 :]:
            if child.tag != qn("w:p"):
                continue

            pPr = child.find(qn("w:pPr"))
            pStyle_elem = pPr.find(qn("w:pStyle")) if pPr is not None else None
            style_id = pStyle_elem.get(qn("w:val")) if pStyle_elem is not None else None

            fonts = _MAPFRE_RFONTS.get(style_id, _MAPFRE_RFONTS[None])

            for run in child.findall(qn("w:r")):
                rPr = run.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    run.insert(0, rPr)
                if rPr.find(qn("w:rFonts")) is None:
                    rPr.insert(0, _make_rfonts(*fonts))

    @staticmethod
    def _replace_across_runs(element, find: str, replace: str) -> None:
        """Replace *find* with *replace* even when the text spans multiple w:t nodes."""
        while True:
            t_nodes = list(element.iter(qn("w:t")))
            texts = [t.text or "" for t in t_nodes]
            full = "".join(texts)
            if find not in full:
                break

            start = full.index(find)
            end = start + len(find)

            # Map character ranges to node indices
            cum, spans = 0, []
            for txt in texts:
                spans.append((cum, cum + len(txt)))
                cum += len(txt)

            affected = [
                i for i, (ns, ne) in enumerate(spans) if ne > start and ns < end
            ]
            if not affected:
                break

            fi, li = affected[0], affected[-1]
            prefix = texts[fi][: max(0, start - spans[fi][0])]
            suffix = texts[li][max(0, end - spans[li][0]) :]

            if fi == li:
                t_nodes[fi].text = prefix + replace + suffix
            else:
                t_nodes[fi].text = prefix + replace
                for i in affected[1:-1]:
                    t_nodes[i].text = ""
                t_nodes[li].text = suffix

    def _fill_cover_page(self, cover_fields: dict) -> None:
        """Replace cover-page placeholders with the values in *cover_fields*."""
        body = self.doc.element.body
        children = list(body)
        cover_end = self._find_cover_end()

        audit_code = cover_fields.get("audit_code", "")
        audit_title = cover_fields.get("audit_title", "")
        uai = cover_fields.get("uai", "")
        date = cover_fields.get("date", "")
        recipients = cover_fields.get("recipients", [])

        for elem in children[: cover_end + 1]:
            if audit_code:
                self._replace_across_runs(elem, "26XX-XXXX", audit_code)
            if audit_title:
                self._replace_across_runs(
                    elem, "TÍTULO DE LA AUDITORÍA", audit_title.upper()
                )
            if uai:
                self._replace_across_runs(elem, "UAI de XXX", uai)
            if date:
                for placeholder in ("xx/xx/xxxx", "XX/XX/XXXX", "xx/XX/XXXX"):
                    self._replace_across_runs(elem, placeholder, date)

        # Recipients — each `D. ` or `Dª. ` w:t node is a placeholder slot
        if recipients:
            slot_tags = ("D. ", "Dª. ")
            slot_idx = 0
            for elem in children[: cover_end + 1]:
                for t in elem.iter(qn("w:t")):
                    if slot_idx >= len(recipients):
                        break
                    if t.text in slot_tags:
                        t.text = recipients[slot_idx]
                        slot_idx += 1
