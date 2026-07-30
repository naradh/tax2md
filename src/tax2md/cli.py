"""tax2md: Convert a ClearTax computation HTML report into clean, AI-ready Markdown.

The hard part of a ClearTax/Office HTML export is that the whole document is
usually built out of *layout* tables (a table wrapping the real content for
positioning), and cells often contain block content — nested tables, several
paragraphs, ``<br>`` line breaks. Pandoc's GitHub-Flavored-Markdown pipe-table
writer cannot represent any of that and collapses the entire table to the
literal string ``[TABLE]`` — i.e. total data loss.

So instead of streaming the HTML straight to pandoc, we first parse it into a
small DOM, drop scripting/styling/Office cruft, *unwrap* layout tables, and
*inline-flatten* the cells of genuine data tables. Only then is well-behaved
HTML handed to pandoc, followed by a light Markdown clean-up pass.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from html.parser import HTMLParser
from pathlib import Path

USAGE = """Usage:
tax2md <html file>

Tip: Drag the HTML file from Finder into Terminal after typing 'tax2md '.
"""


class ToolError(Exception):
    """A user-facing error that should be printed without a traceback."""


# --- Tag/attribute policy -------------------------------------------------

# Attributes that only affect visual presentation, never content/structure.
# Structural attributes (colspan, rowspan, href, ...) are deliberately absent
# so they are preserved.
PRESENTATION_ATTRS = {
    "style", "class", "align", "valign", "bgcolor", "color", "face", "size",
    "width", "height", "lang", "dir", "id", "name", "border", "cellspacing",
    "cellpadding", "background", "link", "vlink", "alink", "topmargin",
    "leftmargin", "marginwidth", "marginheight", "clear", "target", "nowrap",
    "hspace", "vspace",
}

# Tags with no semantic meaning for a text report: unwrap them (keep children).
UNWRAP_TAGS = {"font", "span", "div", "center", "u"}

# Office/Word namespaced wrappers that hold real content: unwrap, keep children.
NS_UNWRAP = {"o:p", "w:sdt", "w:sdtcontent", "w:smarttag", "st1:place",
             "st1:country-region"}

# Tags that are pure presentation/metadata: drop them *and* their content.
DROP_TAGS = {
    "script", "style", "link", "meta", "noscript", "iframe", "object",
    "embed", "base", "head", "title", "xml", "col", "colgroup",
}

# Void elements never have a closing tag.
VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
    "meta", "param", "source", "track", "wbr",
}

# A layout table's own scaffolding (everything *except* a nested <table>, which
# is preserved whole when the surrounding layout table is unwrapped).
LAYOUT_SCAFFOLD = {"thead", "tbody", "tfoot", "tr", "td", "th", "caption",
                   "colgroup", "col"}

# Block-level tags that may not appear inside a Markdown table cell.
BLOCK_IN_CELL = {"p", "div", "br", "hr", "ul", "ol", "li", "table", "tr",
                 "td", "th", "blockquote", "pre", "h1", "h2", "h3", "h4",
                 "h5", "h6", "center"}

HIDDEN_STYLE_RE = re.compile(r"display\s*:\s*none|visibility\s*:\s*hidden", re.I)


# --- A tiny tolerant HTML DOM ---------------------------------------------

class Element:
    __slots__ = ("tag", "attrs", "children")

    def __init__(self, tag: str, attrs=None):
        self.tag = tag
        self.attrs: list[tuple[str, str]] = attrs or []
        self.children: list = []


# Elements whose end tag is optional, and the tags that implicitly close them.
_IMPLICIT_CLOSERS = {
    "p": {"p", "div", "table", "tr", "td", "th", "ul", "ol", "li", "h1",
          "h2", "h3", "h4", "h5", "h6", "hr", "blockquote", "pre", "thead",
          "tbody", "tfoot"},
    "li": {"li"},
    "tr": {"tr", "thead", "tbody", "tfoot"},
    "td": {"td", "th", "tr", "thead", "tbody", "tfoot"},
    "th": {"td", "th", "tr", "thead", "tbody", "tfoot"},
    "dt": {"dt", "dd"},
    "dd": {"dt", "dd"},
    "option": {"option"},
    "thead": {"tbody", "tfoot"},
    "tbody": {"thead", "tfoot"},
}


class DOMBuilder(HTMLParser):
    """Builds a forgiving DOM tree, tolerating the missing/optional end tags
    and stray closes common in generated report HTML."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Element("#root")
        self._stack = [self.root]

    def _implicit_close(self, new_tag: str) -> None:
        # Pop any open elements that `new_tag` is defined to auto-close.
        while len(self._stack) > 1:
            open_tag = self._stack[-1].tag
            closers = _IMPLICIT_CLOSERS.get(open_tag)
            if closers and new_tag in closers:
                self._stack.pop()
            else:
                break

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        self._implicit_close(tag)
        el = Element(tag, attrs)
        self._stack[-1].children.append(el)
        if tag not in VOID_TAGS:
            self._stack.append(el)

    def handle_startendtag(self, tag, attrs):
        tag = tag.lower()
        self._implicit_close(tag)
        self._stack[-1].children.append(Element(tag, attrs))

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in VOID_TAGS:
            return
        # Pop up to and including the nearest matching open tag; ignore strays.
        for depth in range(len(self._stack) - 1, 0, -1):
            if self._stack[depth].tag == tag:
                del self._stack[depth:]
                return

    def handle_data(self, data):
        self._stack[-1].children.append(data)


# --- DOM transformations ---------------------------------------------------

def _is_hidden(attrs) -> bool:
    for name, value in attrs:
        if name and name.lower() == "style" and value and HIDDEN_STYLE_RE.search(value):
            return True
    return False


def _text_content(node) -> str:
    if isinstance(node, str):
        return node
    return "".join(_text_content(c) for c in node.children)


def _is_refresh_link(node: Element) -> bool:
    """A ClearTax report opens with an <a class="refresh-btn">Click here to
    refresh</a> — a presentation-only reload link that should be removed."""
    for name, value in node.attrs:
        if name and name.lower() == "class" and value and "refresh" in value.lower():
            return True
    return "click here to refresh" in _text_content(node).strip().lower()


def _keep_attr(name: str) -> bool:
    if not name:
        return False
    name = name.lower()
    return (name not in PRESENTATION_ATTRS
            and ":" not in name
            and not name.startswith("on"))


def _get_rows(table: Element) -> list[Element]:
    rows: list[Element] = []

    def walk(node: Element) -> None:
        for ch in node.children:
            if isinstance(ch, Element):
                if ch.tag == "tr":
                    rows.append(ch)
                elif ch.tag in ("thead", "tbody", "tfoot"):
                    walk(ch)

    walk(table)
    return rows


def _cell_count(row: Element) -> int:
    return sum(1 for ch in row.children
               if isinstance(ch, Element) and ch.tag in ("td", "th"))


def _contains_table(node: Element) -> bool:
    for ch in node.children:
        if isinstance(ch, Element):
            if ch.tag == "table" or _contains_table(ch):
                return True
    return False


def _promote(node: Element) -> list:
    """Pull the meaningful content out of a layout table, discarding the
    table/row/cell scaffolding but keeping everything else in document order."""
    result: list = []
    for ch in node.children:
        if isinstance(ch, str):
            result.append(ch)
        elif ch.tag in LAYOUT_SCAFFOLD:
            result.extend(_promote(ch))
        else:
            # Nested <table>, headings, paragraphs, etc. are kept whole.
            result.append(ch)
    return result


def _inline_flatten(children: list) -> list:
    """Return an inline-only version of a cell's children: block elements and
    ``<br>`` become spaces (so distinct values never fuse), inline formatting
    such as ``<b>``/``<i>`` is preserved."""
    out: list = []
    for ch in children:
        if isinstance(ch, str):
            out.append(ch)
        elif ch.tag in BLOCK_IN_CELL:
            out.append(" ")
            out.extend(_inline_flatten(ch.children))
            out.append(" ")
        else:
            ch.children = _inline_flatten(ch.children)
            ch.attrs = [(n, v) for n, v in ch.attrs if _keep_attr(n)]
            out.append(ch)
    return out


def _handle_table(table: Element) -> list:
    rows = _get_rows(table)
    max_cells = max((_cell_count(r) for r in rows), default=0)
    is_layout = _contains_table(table) or max_cells <= 1
    if is_layout:
        return _promote(table)
    # Genuine data table: make every cell inline-only so pandoc can render it.
    for row in rows:
        for cell in row.children:
            if isinstance(cell, Element) and cell.tag in ("td", "th"):
                cell.children = _inline_flatten(cell.children)
                cell.attrs = [(n, v) for n, v in cell.attrs
                              if n and n.lower() in ("colspan", "rowspan", "scope")]
    table.attrs = []
    return [table]


def _transform(node):
    """Post-order clean-up of one node; returns its replacement list."""
    if isinstance(node, str):
        return [node]
    tag = node.tag
    if tag in DROP_TAGS or _is_hidden(node.attrs):
        return []
    if ":" in tag and tag not in NS_UNWRAP:
        return []  # unknown namespaced element (o:OfficeDocumentSettings, ...)

    new_children: list = []
    for ch in node.children:
        new_children.extend(_transform(ch))
    node.children = new_children

    if tag == "a" and _is_refresh_link(node):
        return []
    if tag in UNWRAP_TAGS or tag in NS_UNWRAP:
        return node.children
    if tag == "table":
        return _handle_table(node)
    node.attrs = [(n, v) for n, v in node.attrs if _keep_attr(n)]
    return [node]


# --- Serialization back to HTML for pandoc ---------------------------------

def _escape_text(text: str) -> str:
    return (text.replace("\xa0", " ")
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))


def _serialize(node, out: list) -> None:
    if isinstance(node, str):
        out.append(_escape_text(node))
        return
    attrs = "".join(f' {n}="{str(v).replace(chr(34), "&quot;")}"'
                    for n, v in node.attrs if v is not None)
    attrs += "".join(f" {n}" for n, v in node.attrs if v is None)
    out.append(f"<{node.tag}{attrs}>")
    if node.tag in VOID_TAGS:
        return
    for ch in node.children:
        _serialize(ch, out)
    out.append(f"</{node.tag}>")


# --- Pipeline stages -------------------------------------------------------

def decode_html(raw: bytes) -> str:
    """Detect UTF-8 (with or without BOM) vs Windows-1252 and decode.

    A genuine UTF-8 multi-byte sequence is effectively never valid-but-wrong
    Windows-1252, so a strict UTF-8 attempt is a reliable discriminator. The
    Windows-1252 fallback uses ``errors="replace"`` because a handful of 1252
    byte positions are undefined and would otherwise raise.
    """
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace")


def clean_html(html_text: str) -> str:
    builder = DOMBuilder()
    builder.feed(html_text)
    builder.close()

    children: list = []
    for ch in builder.root.children:
        children.extend(_transform(ch))

    out: list[str] = []
    for ch in children:
        _serialize(ch, out)
    return "".join(out)


def html_to_markdown(html_text: str) -> str:
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise ToolError(
            "pandoc is required but was not found on PATH.\n"
            "Install it with: brew install pandoc"
        )
    result = subprocess.run(
        [pandoc, "-f", "html", "-t", "gfm-raw_html", "--wrap=none"],
        input=html_text.encode("utf-8"),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode != 0:
        raise ToolError("pandoc failed:\n" + result.stderr.decode("utf-8", "replace"))
    return result.stdout.decode("utf-8")


_CELL_SPLIT_RE = re.compile(r"(?<!\\)\|")
_SEP_CELL_RE = re.compile(r"^:?-{1,}:?$")


def _split_row(line: str) -> list[str]:
    inner = line.strip()
    inner = inner[1:] if inner.startswith("|") else inner
    inner = inner[:-1] if inner.endswith("|") else inner
    return [c.strip() for c in _CELL_SPLIT_RE.split(inner)]


def _is_blank_cell(cell: str) -> bool:
    return cell.replace("\\", "").strip() == ""


def _is_separator_row(cells: list[str]) -> bool:
    return bool(cells) and all(_SEP_CELL_RE.match(c) for c in cells)


def _prune_table_block(block: list[str]) -> list[str]:
    """Drop columns and body rows that are empty everywhere. Leaves the block
    untouched if it is at all irregular, so a malformed table is never mangled."""
    rows = [_split_row(line) for line in block]
    width = len(rows[0])
    if width < 2 or any(len(r) != width for r in rows):
        return block
    sep_idx = next((i for i, r in enumerate(rows) if _is_separator_row(r)), None)
    if sep_idx is None:
        return block

    keep = [c for c in range(width)
            if not all(_is_blank_cell(rows[i][c])
                       for i in range(len(rows)) if i != sep_idx)]
    if not keep or len(keep) == width:
        kept_rows = rows
    else:
        kept_rows = [[r[c] for c in keep] for r in rows]

    out_rows = [
        r for i, r in enumerate(kept_rows)
        if i <= sep_idx or not all(_is_blank_cell(c) for c in r)
    ]
    return ["| " + " | ".join(r) + " |" for r in out_rows]


def _prune_tables(md: str) -> str:
    lines = md.split("\n")
    out: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i].lstrip().startswith("|"):
            j = i
            while j < len(lines) and lines[j].lstrip().startswith("|"):
                j += 1
            out.extend(_prune_table_block(lines[i:j]))
            i = j
        else:
            out.append(lines[i])
            i += 1
    return "\n".join(out)


# Trailing boilerplate every ClearTax report ends with (signature block and
# "Generated by ClearTax" credit) — no value to an AI, so drop it.
_FOOTER_RE = (
    re.compile(r"^signature$", re.I),
    re.compile(r"^for .{1,60}$", re.I),
    re.compile(r"^generated by cleartax\b", re.I),
)


def _strip_footer(md: str) -> str:
    lines = md.split("\n")
    changed = True
    while changed:
        changed = False
        while lines and lines[-1].strip() == "":
            lines.pop()
            changed = True
        if lines and any(p.match(lines[-1].strip()) for p in _FOOTER_RE):
            lines.pop()
            changed = True
    return "\n".join(lines)


def clean_markdown(md: str) -> str:
    # Drop any raw HTML tags pandoc still left behind.
    md = re.sub(r"</?(?:br|hr|div|span|font|o:p)\b[^>]*/?>", "", md, flags=re.I)
    # Strip pandoc header/span identifier & class attributes, e.g. "## Foo {#foo}".
    md = re.sub(r"[ \t]*\{[^{}\n]*\}", "", md)
    # Remove empty link and emphasis markers left after the strips above.
    md = re.sub(r"\[\]\(\)", "", md)
    md = re.sub(r"\*\*\s*\*\*", "", md)
    md = re.sub(r"__\s*__", "", md)
    md = re.sub(r"(?<!\*)\*\s+\*(?!\*)", "", md)
    md = re.sub(r"(?<!_)_\s+_(?!_)", "", md)
    # Unescape characters pandoc escapes defensively but which read fine as text.
    md = re.sub(r"\\([%$#&_{}.\-+!()\[\]])", r"\1", md)
    # Normalize line endings; strip trailing hard-break backslashes (leftover
    # from presentational <br>) and trailing whitespace.
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = "\n".join(re.sub(r"(?<!\\)\\$", "", line).rstrip() for line in md.split("\n"))
    # Drop any now-empty stray-backslash lines.
    md = "\n".join(line for line in md.split("\n") if line.strip() != "\\")
    # Remove all-empty columns and spacer rows from tables.
    md = _prune_tables(md)
    # Collapse runs of blank lines down to a single blank line.
    md = re.sub(r"\n{3,}", "\n\n", md)
    # Drop ClearTax's trailing signature/credit boilerplate.
    md = _strip_footer(md)
    return md.strip() + "\n"


def convert(html_path: Path) -> Path:
    raw = html_path.read_bytes()
    text = decode_html(raw)
    cleaned_html = clean_html(text)
    md = html_to_markdown(cleaned_html)
    md = clean_markdown(md)
    out_path = html_path.with_suffix(".md")
    out_path.write_text(md, encoding="utf-8")
    return out_path


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv

    if len(argv) != 1:
        print(USAGE)
        return 0 if len(argv) == 0 else 1

    html_path = Path(argv[0]).expanduser()

    if not html_path.exists():
        print(f"Error: file not found: {html_path}", file=sys.stderr)
        return 1
    if not html_path.is_file():
        print(f"Error: not a file: {html_path}", file=sys.stderr)
        return 1
    if html_path.suffix.lower() not in (".html", ".htm"):
        print(f"Error: expected an .html file, got: {html_path.name}", file=sys.stderr)
        return 1

    try:
        out_path = convert(html_path)
    except ToolError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"Error: could not process {html_path.name}: {exc}", file=sys.stderr)
        return 1

    print("\u2713 Markdown created:")
    print(str(out_path.resolve()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
