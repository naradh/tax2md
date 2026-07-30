# tax2md

Convert a **ClearTax computation report** into clean, AI-ready Markdown — so you
can paste it into Claude, ChatGPT, etc. using **~80% fewer tokens** than the raw
HTML (and get better answers, because the model isn't wading through markup).

Two ways to use it:

- **Web app** — [tax2md.naradh.in](https://tax2md.naradh.in). 100% in your
  browser; the file is never uploaded and nothing is stored server-side.
- **CLI** — a Python command for local/batch use.

> **What it's for:** ClearTax *computation reports* saved from your browser as
> **“Webpage, Complete”** (a `.html` file). Not PDFs, not partial saves.

## Why it's not just "strip the tags"

ClearTax/Office HTML exports are built entirely out of **layout tables**, and
cells routinely contain block content (nested tables, multiple paragraphs,
`<br>` line breaks). A naïve HTML→Markdown conversion collapses those to the
literal string `[TABLE]` — i.e. the whole report is lost. tax2md instead:

- Detects **UTF-8** (with/without BOM) vs **Windows-1252** and decodes cleanly.
- Removes CSS, JavaScript, comments, meta-refresh redirects, refresh links and
  hidden elements.
- **Unwraps layout tables** and **inline-flattens data-table cells** (so two
  numbers separated by `<br>` never fuse into one).
- **Prunes empty spacer columns/rows** that clutter the tables.
- Preserves headings, tables, notes and **every numerical value exactly** —
  negatives like `-2,596`, percentages, and the `₹` symbol.

Verified faithful on a real report: every visible number preserved,
number-for-number.

## Web app

Static, dependency-free files in [`web/`](web/):

- `index.html` — the page
- `styles.css` — styling (light/dark aware)
- `tax2md.js` — the converter (uses the browser's native `DOMParser`; no server)

Open `web/index.html` directly, or serve the folder with any static host.

### Deploying

The site is 100% static, so any static host works. It's designed to run with a
domain whose DNS stays where it is (e.g. GoDaddy) via a simple CNAME:

- **Vercel** — import this repo, set **Root Directory** = `web`, framework
  **Other**, no build command. Add the domain in Vercel and point a CNAME at
  `cname.vercel-dns.com`.
- **Cloudflare** — [`wrangler.jsonc`](wrangler.jsonc) serves `./web` as a
  static-assets Worker (`npx wrangler deploy`). Note: a Cloudflare *custom
  domain* requires the domain's DNS zone to be on Cloudflare.
- **GitHub Pages / Netlify** — also work; point the site at the `web/` folder.

## CLI

The command-line version uses [Pandoc](https://pandoc.org) under the hood.

### Requirements

- Python ≥ 3.9 (standard library only — no third-party Python packages)
- Pandoc: `brew install pandoc`

### Install

```sh
python3 -m pip install --user -e .
```

pip installs the `tax2md` script into your Python user `bin` (e.g.
`~/Library/Python/3.9/bin`). If that directory isn't on your `PATH`, symlink the
script somewhere that is:

```sh
ln -s ~/Library/Python/3.9/bin/tax2md /opt/homebrew/bin/tax2md
```

### Use

```sh
tax2md <html_file>
```

The Markdown is written **beside** the HTML with the same name
(`…Computation.html` → `…Computation.md`). Tip: type `tax2md ` then drag the HTML
file from Finder into Terminal.

### Test

```sh
python3 tests/test_convert.py     # no dependencies
# or, with pytest:
python3 -m pytest
```

## License

[MIT](LICENSE)
