"""
Core document processor: loads a Word template, clears its body,
then rebuilds content from parsed markdown — preserving all template
styles, page layout, headers, footers, and images.
"""

import base64
import re
import urllib.request
from io import BytesIO
from typing import Optional

import markdown
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _collect_style_names(doc: Document) -> set:
    return {s.name for s in doc.styles}


def _best_style(available: set, *candidates: str) -> Optional[str]:
    """Return the first candidate that exists in the document styles."""
    for name in candidates:
        if name and name in available:
            return name
    return None


def _add_horizontal_rule(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_hyperlink(paragraph, text: str, url: str) -> None:
    if not url:
        paragraph.add_run(text)
        return
    try:
        r_id = paragraph.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        new_run = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        rStyle = OxmlElement("w:rStyle")
        rStyle.set(qn("w:val"), "Hyperlink")
        rPr.append(rStyle)
        new_run.append(rPr)
        t = OxmlElement("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        new_run.append(t)
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
    except Exception:
        run = paragraph.add_run(text)
        run.underline = True


def _fetch_image(src: str) -> Optional[bytes]:
    """Return raw bytes for a data-URI or http(s) image URL."""
    if src.startswith("data:"):
        m = re.match(r"data:[^;]+;base64,(.+)", src, re.DOTALL)
        if m:
            return base64.b64decode(m.group(1))
    elif src.startswith(("http://", "https://")):
        try:
            req = urllib.request.Request(
                src, headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                return resp.read()
        except Exception:
            return None
    return None


# ---------------------------------------------------------------------------
# Inline content builder
# ---------------------------------------------------------------------------

def _add_run_with_fmt(
    paragraph,
    text: str,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    strike: bool = False,
    code: bool = False,
) -> None:
    if not text:
        return
    run = paragraph.add_run(text)
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if underline:
        run.underline = True
    if strike:
        run.font.strike = True
    if code:
        run.font.name = "Courier New"
        run.font.size = Pt(9)


def _inline(
    paragraph,
    element,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    strike: bool = False,
    code: bool = False,
) -> None:
    """Recursively walk inline HTML elements and add runs to *paragraph*."""
    for child in element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                _add_run_with_fmt(
                    paragraph, text, bold, italic, underline, strike, code
                )
            continue

        if not hasattr(child, "name") or child.name is None:
            continue

        tag = child.name.lower()

        if tag in ("strong", "b"):
            _inline(paragraph, child, bold=True, italic=italic,
                    underline=underline, strike=strike, code=code)
        elif tag in ("em", "i"):
            _inline(paragraph, child, bold=bold, italic=True,
                    underline=underline, strike=strike, code=code)
        elif tag == "u":
            _inline(paragraph, child, bold=bold, italic=italic,
                    underline=True, strike=strike, code=code)
        elif tag in ("s", "del", "strike"):
            _inline(paragraph, child, bold=bold, italic=italic,
                    underline=underline, strike=True, code=code)
        elif tag == "code":
            _inline(paragraph, child, bold=bold, italic=italic,
                    underline=underline, strike=strike, code=True)
        elif tag == "a":
            href = child.get("href", "")
            link_text = child.get_text()
            if href:
                _add_hyperlink(paragraph, link_text, href)
            else:
                _add_run_with_fmt(paragraph, link_text, bold, italic,
                                   underline, strike, code)
        elif tag == "br":
            paragraph.add_run("\n")
        elif tag == "img":
            src = child.get("src", "")
            alt = child.get("alt", "")
            if src:
                img_bytes = _fetch_image(src)
                if img_bytes:
                    try:
                        run = paragraph.add_run()
                        run.add_picture(BytesIO(img_bytes), width=Inches(3))
                    except Exception:
                        _add_run_with_fmt(paragraph, f"[Image: {alt}]")
                else:
                    _add_run_with_fmt(paragraph, f"[Image: {alt}]")
            elif alt:
                _add_run_with_fmt(paragraph, f"[Image: {alt}]")
        elif tag == "span":
            _inline(paragraph, child, bold, italic, underline, strike, code)
        else:
            # Unknown inline tag — just emit its text
            text = child.get_text()
            if text:
                _add_run_with_fmt(paragraph, text, bold, italic,
                                   underline, strike, code)


# ---------------------------------------------------------------------------
# Main processor class
# ---------------------------------------------------------------------------

class DocumentProcessor:
    """
    Loads a .docx template, clears its body, and fills it with
    the content described by markdown_text, using the template's own styles.
    """

    def __init__(self):
        self.doc: Optional[Document] = None
        self.styles: set = set()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def process(self, template_bytes: bytes, markdown_text: str) -> bytes:
        self.doc = Document(BytesIO(template_bytes))
        self.styles = _collect_style_names(self.doc)

        self._clear_body()

        html = markdown.markdown(
            markdown_text,
            extensions=[
                "tables",
                "fenced_code",
                "nl2br",
                "sane_lists",
                "attr_list",
                "def_list",
                "footnotes",
                "md_in_html",
            ],
        )

        soup = BeautifulSoup(f"<root>{html}</root>", "html.parser")
        self._process_children(soup.find("root"))

        out = BytesIO()
        self.doc.save(out)
        return out.getvalue()

    # ------------------------------------------------------------------
    # Body management
    # ------------------------------------------------------------------

    def _clear_body(self) -> None:
        """Remove every body child except the final sectPr."""
        body = self.doc.element.body
        to_remove = [c for c in body if c.tag != qn("w:sectPr")]
        for elem in to_remove:
            body.remove(elem)

    # ------------------------------------------------------------------
    # Tree traversal
    # ------------------------------------------------------------------

    def _process_children(self, element) -> None:
        for child in element.children:
            if isinstance(child, NavigableString):
                text = str(child).strip()
                if text:
                    self._add_paragraph(text, "Normal")
            elif hasattr(child, "name") and child.name:
                self._process_block(child)

    def _process_block(self, element) -> None:
        tag = element.name.lower()

        # --- Headings ---
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            style = _best_style(self.styles, f"Heading {level}", "Normal")
            p = self.doc.add_paragraph(style=style)
            _inline(p, element)

        # --- Paragraph ---
        elif tag == "p":
            # Image-only paragraph
            imgs = element.find_all("img")
            if imgs and not element.get_text(strip=True):
                p = self.doc.add_paragraph()
                for img in imgs:
                    src = img.get("src", "")
                    alt = img.get("alt", "")
                    if src:
                        img_bytes = _fetch_image(src)
                        if img_bytes:
                            try:
                                run = p.add_run()
                                run.add_picture(BytesIO(img_bytes), width=Inches(4))
                                continue
                            except Exception:
                                pass
                    p.add_run(f"[Image: {alt}]" if alt else "[Image]")
            else:
                style = _best_style(self.styles, "Body Text", "Normal")
                p = self.doc.add_paragraph(style=style)
                _inline(p, element)

        # --- Lists ---
        elif tag == "ul":
            self._process_list(element, ordered=False, level=0)
        elif tag == "ol":
            self._process_list(element, ordered=True, level=0)

        # --- Blockquote ---
        elif tag == "blockquote":
            q_style = _best_style(
                self.styles, "Intense Quote", "Quote", "Body Text", "Normal"
            )
            for child in element.children:
                if isinstance(child, NavigableString):
                    text = str(child).strip()
                    if text:
                        self.doc.add_paragraph(text, style=q_style)
                elif hasattr(child, "name") and child.name in ("p", "div", "span"):
                    p = self.doc.add_paragraph(style=q_style)
                    _inline(p, child)
                else:
                    # Recurse for nested blockquotes or other blocks
                    self._process_block(child)

        # --- Code block ---
        elif tag == "pre":
            code_elem = element.find("code")
            code_text = (code_elem if code_elem else element).get_text()
            style = _best_style(self.styles, "Code", "No Spacing", "Normal")
            p = self.doc.add_paragraph(style=style)
            run = p.add_run(code_text)
            if style in (None, "Normal", "No Spacing"):
                run.font.name = "Courier New"
                run.font.size = Pt(9)

        # --- Table ---
        elif tag == "table":
            self._process_table(element)

        # --- Horizontal rule ---
        elif tag == "hr":
            p = self.doc.add_paragraph()
            _add_horizontal_rule(p)

        # --- Standalone image ---
        elif tag == "img":
            p = self.doc.add_paragraph()
            src = element.get("src", "")
            alt = element.get("alt", "")
            if src:
                img_bytes = _fetch_image(src)
                if img_bytes:
                    try:
                        run = p.add_run()
                        run.add_picture(BytesIO(img_bytes), width=Inches(4))
                        return
                    except Exception:
                        pass
            p.add_run(f"[Image: {alt}]" if alt else "[Image]")

        # --- Generic container ---
        elif tag in ("div", "section", "article", "main", "aside", "header",
                     "footer", "figure", "figcaption"):
            self._process_children(element)

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------

    def _process_list(self, element, ordered: bool, level: int) -> None:
        base_style = "List Number" if ordered else "List Bullet"
        # python-docx ships "List Bullet 2", "List Bullet 3", etc.
        level_suffix = "" if level == 0 else f" {min(level + 1, 3)}"
        style = _best_style(
            self.styles,
            f"{base_style}{level_suffix}",
            base_style,
            "Normal",
        )

        for child in element.children:
            if not hasattr(child, "name") or child.name != "li":
                continue

            # Collect nested sub-list (if any)
            nested = child.find(["ul", "ol"])

            p = self.doc.add_paragraph(style=style)
            if level > 0:
                p.paragraph_format.left_indent = Inches(0.25 * (level + 1))

            # Add li text, skipping nested list nodes
            for item_child in child.children:
                if hasattr(item_child, "name") and item_child.name in ("ul", "ol"):
                    continue
                if isinstance(item_child, NavigableString):
                    text = str(item_child)
                    if text.strip():
                        p.add_run(text)
                elif hasattr(item_child, "name"):
                    _inline(p, item_child)

            # Recurse into nested list
            if nested:
                self._process_list(
                    nested, ordered=(nested.name == "ol"), level=level + 1
                )

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------

    def _process_table(self, element) -> None:
        rows = element.find_all("tr")
        if not rows:
            return

        max_cols = max(
            (len(r.find_all(["td", "th"])) for r in rows), default=0
        )
        if max_cols == 0:
            return

        table = self.doc.add_table(rows=len(rows), cols=max_cols)

        # Apply a table style if one exists
        for candidate in ("Table Grid", "Light Grid", "Medium Grid 1",
                          "Medium Shading 1 Accent 1", "Table Normal"):
            if candidate in self.styles:
                try:
                    table.style = candidate
                except Exception:
                    pass
                break

        for row_idx, tr in enumerate(rows):
            cells = tr.find_all(["td", "th"])
            for col_idx, cell_elem in enumerate(cells):
                if col_idx >= max_cols:
                    break
                cell = table.cell(row_idx, col_idx)
                # Clear the auto-created empty paragraph
                cell.paragraphs[0].clear()
                p = cell.paragraphs[0]
                _inline(p, cell_elem)
                if cell_elem.name == "th":
                    for run in p.runs:
                        run.bold = True

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    def _add_paragraph(self, text: str, style: str) -> None:
        resolved = _best_style(self.styles, style, "Normal")
        self.doc.add_paragraph(text, style=resolved)
