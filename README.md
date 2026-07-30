# tax2md

Convert a ClearTax computation HTML report into clean, AI-ready Markdown.

```
tax2md <html_file>
```

The Markdown is written **beside** the HTML with the same name
(`…Computation.html` → `…Computation.md`). Typically you drag the HTML file
from Finder into Terminal after typing `tax2md `.

## What it does

- Accepts exactly one `.html`/`.htm` file.
- Auto-detects **UTF-8** (with or without BOM) vs **Windows-1252** and decodes
  to UTF-8 internally.
- Removes CSS, JavaScript, comments, meta-refresh redirects, hidden elements
  and other presentation-only markup.
- **Unwraps layout tables** and **inline-flattens data-table cells** so the
  report survives conversion. (Office/ClearTax exports wrap everything in
  layout tables and put block content inside cells; fed to pandoc directly,
  that collapses to the literal string `[TABLE]` — total data loss. tax2md
  handles this before pandoc ever sees the document.)
- Converts HTML → Markdown with **Pandoc**, then cleans the result: leftover
  HTML, defensive backslash-escapes and excess blank lines are removed.
- Preserves headings, tables, notes and **all numerical values exactly** —
  including negatives like `(50,000)`, percentages, and the `₹` symbol.

## Requirements

- Python ≥ 3.9 (standard library only — no third-party Python packages)
- [Pandoc](https://pandoc.org): `brew install pandoc`

## Install

```sh
python3 -m pip install --user -e .
```

pip installs the `tax2md` script into your Python user `bin` (e.g.
`~/Library/Python/3.9/bin`). If that directory isn't on your `PATH`, symlink
the script somewhere that is — for example:

```sh
ln -s ~/Library/Python/3.9/bin/tax2md /opt/homebrew/bin/tax2md
```

On this machine that symlink is already in place, so `tax2md` works from any
directory.

## Test

```sh
python3 tests/test_convert.py     # no dependencies
# or, if you have pytest:
python3 -m pytest
```
