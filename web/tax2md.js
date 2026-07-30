/*
 * tax2md — browser converter (no server, no upload).
 *
 * Converts a ClearTax computation report saved as "Webpage, Complete" (.html)
 * into clean, AI-ready Markdown, entirely in the browser. This is a faithful
 * port of the Python CLI: it decodes UTF-8 / Windows-1252, unwraps the layout
 * tables that Office/ClearTax exports are built from, inline-flattens genuine
 * data-table cells (so distinct numbers never fuse), and prunes empty spacer
 * columns and rows.
 *
 * Public API (attached to window.tax2md):
 *   decodeBytes(ArrayBuffer|Uint8Array) -> string
 *   htmlToMarkdown(string) -> string
 *   convertBytes(ArrayBuffer|Uint8Array) -> string
 */
(function (global) {
  "use strict";

  // --- Tag policy (mirrors the Python CLI) ---------------------------------

  var DROP = new Set([
    "script", "style", "link", "meta", "noscript", "iframe", "object",
    "embed", "base", "title", "head", "xml", "col", "colgroup",
  ]);

  // Office/Word namespaced wrappers that hold real content: render through.
  var NS_UNWRAP = new Set([
    "o:p", "w:sdt", "w:sdtcontent", "w:smarttag",
    "st1:place", "st1:country-region",
  ]);

  // Semantically empty wrappers: render their children through.
  var TRANSPARENT = new Set(["font", "span", "div", "center", "u", "small",
    "o:p", "w:sdt", "w:sdtcontent", "w:smarttag", "st1:place",
    "st1:country-region"]);

  // Block-level tags that force spacing when flattened into a table cell.
  var BLOCK_IN_CELL = new Set(["p", "div", "br", "hr", "ul", "ol", "li",
    "table", "tr", "td", "th", "blockquote", "pre", "h1", "h2", "h3", "h4",
    "h5", "h6", "center"]);

  // A layout table's own scaffolding (everything but a nested <table>).
  var SCAFFOLD = new Set(["thead", "tbody", "tfoot", "tr", "td", "th",
    "caption", "colgroup", "col"]);

  var HIDDEN_RE = /display\s*:\s*none|visibility\s*:\s*hidden/i;

  // --- Small helpers --------------------------------------------------------

  function tagOf(el) { return el.tagName.toLowerCase(); }

  function childNodes(el) { return Array.prototype.slice.call(el.childNodes); }

  function collapse(s) { return s.replace(/[\s ]+/g, " "); }

  function isHidden(el) {
    var s = el.getAttribute && el.getAttribute("style");
    return !!s && HIDDEN_RE.test(s);
  }

  function textContent(el) {
    return (el.textContent || "");
  }

  function isRefreshLink(el) {
    var cls = el.getAttribute("class");
    if (cls && /refresh/i.test(cls)) return true;
    return textContent(el).trim().toLowerCase().indexOf("click here to refresh") !== -1;
  }

  function shouldDrop(el) {
    var tag = tagOf(el);
    if (DROP.has(tag)) return true;
    if (tag.indexOf(":") !== -1 && !NS_UNWRAP.has(tag)) return true;
    if (isHidden(el)) return true;
    if (tag === "a" && isRefreshLink(el)) return true;
    return false;
  }

  function hasBlockChild(el) {
    return !!el.querySelector(
      "table,p,div,h1,h2,h3,h4,h5,h6,ul,ol,li,tr,hr,blockquote,pre");
  }

  // --- Inline rendering (cells, headings, list items, paragraphs) ----------

  function renderInline(nodes) {
    var s = "";
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (n.nodeType === 3) {           // text
        s += n.nodeValue;
        continue;
      }
      if (n.nodeType !== 1) continue;   // comments etc.
      var el = n, tag = tagOf(el);
      if (shouldDrop(el)) continue;
      if (tag === "br") { s += " "; continue; }
      if (tag === "b" || tag === "strong") {
        var b = collapse(renderInline(childNodes(el))).trim();
        if (b) s += "**" + b + "**";
      } else if (tag === "i" || tag === "em") {
        var it = collapse(renderInline(childNodes(el))).trim();
        if (it) s += "*" + it + "*";
      } else if (tag === "a") {
        var inner = collapse(renderInline(childNodes(el))).trim();
        var href = el.getAttribute("href");
        if (href && !isRefreshLink(el)) s += "[" + inner + "](" + href + ")";
        else s += inner;
      } else if (BLOCK_IN_CELL.has(tag)) {
        s += " " + renderInline(childNodes(el)) + " ";
      } else {                          // span/font/other inline: transparent
        s += renderInline(childNodes(el));
      }
    }
    return s;
  }

  function inlineText(nodes) { return collapse(renderInline(nodes)).trim(); }

  // --- Tables ---------------------------------------------------------------

  function directCells(tr) {
    var out = [];
    for (var i = 0; i < tr.children.length; i++) {
      var t = tagOf(tr.children[i]);
      if (t === "td" || t === "th") out.push(tr.children[i]);
    }
    return out;
  }

  function tableRows(table) {
    // All <tr> belonging to this table (this is only called on data tables,
    // which by construction contain no nested table).
    return Array.prototype.slice.call(table.querySelectorAll("tr"));
  }

  function isLayoutTable(table) {
    if (table.querySelector("table")) return true;   // wraps another table
    var rows = tableRows(table);
    var maxCells = 0;
    for (var i = 0; i < rows.length; i++) {
      maxCells = Math.max(maxCells, directCells(rows[i]).length);
    }
    return maxCells <= 1;                             // single-column / banner
  }

  function promote(node) {
    // Pull meaningful content out of a layout table, dropping the
    // table/row/cell scaffolding but keeping nested tables/headings whole.
    var out = [];
    var kids = childNodes(node);
    for (var i = 0; i < kids.length; i++) {
      var ch = kids[i];
      if (ch.nodeType === 3) { out.push(ch); continue; }
      if (ch.nodeType !== 1) continue;
      if (SCAFFOLD.has(tagOf(ch))) out.push.apply(out, promote(ch));
      else out.push(ch);
    }
    return out;
  }

  function renderTable(table) {
    var rows = tableRows(table);
    var matrix = [];
    for (var r = 0; r < rows.length; r++) {
      var cells = directCells(rows[r]);
      if (cells.length === 0) continue;
      var row = [];
      for (var c = 0; c < cells.length; c++) {
        var span = parseInt(cells[c].getAttribute("colspan") || "1", 10) || 1;
        var text = inlineText(childNodes(cells[c])).replace(/\|/g, "\\|");
        row.push(text);
        for (var k = 1; k < span; k++) row.push("");
      }
      matrix.push(row);
    }
    if (matrix.length === 0) return "";

    var width = 0;
    matrix.forEach(function (row) { width = Math.max(width, row.length); });
    matrix.forEach(function (row) { while (row.length < width) row.push(""); });

    // Prune columns that are empty in every row.
    var keep = [];
    for (var col = 0; col < width; col++) {
      var any = matrix.some(function (row) { return row[col].trim() !== ""; });
      if (any) keep.push(col);
    }
    if (keep.length === 0) return "";
    var m = matrix.map(function (row) {
      return keep.map(function (col) { return row[col]; });
    });
    // Prune all-empty body rows (keep the header row).
    m = m.filter(function (row, idx) {
      return idx === 0 || row.some(function (cell) { return cell.trim() !== ""; });
    });
    if (m.length === 0) return "";

    var fmt = function (row) { return "| " + row.join(" | ") + " |"; };
    var sep = m[0].map(function () { return "----"; });
    var lines = [fmt(m[0]), fmt(sep)];
    for (var j = 1; j < m.length; j++) lines.push(fmt(m[j]));
    return lines.join("\n");
  }

  function renderTableOrPromote(table) {
    if (isLayoutTable(table)) return renderBlocks(promote(table));
    return renderTable(table);
  }

  // --- Lists ----------------------------------------------------------------

  function renderList(el, ordered) {
    var items = [];
    var kids = childNodes(el);
    var n = 1;
    for (var i = 0; i < kids.length; i++) {
      var li = kids[i];
      if (li.nodeType !== 1 || tagOf(li) !== "li") continue;
      var text = inlineText(childNodes(li));
      if (!text) continue;
      items.push((ordered ? (n++) + ". " : "- ") + text);
    }
    return items.join("\n");
  }

  // --- Block rendering ------------------------------------------------------

  function renderBlocks(nodes) {
    var blocks = [];
    var buf = [];
    function flush() {
      var t = inlineText(buf);
      if (t) blocks.push(t);
      buf = [];
    }
    for (var i = 0; i < nodes.length; i++) {
      var n = nodes[i];
      if (n.nodeType === 3) { buf.push(n); continue; }
      if (n.nodeType !== 1) continue;
      var el = n, tag = tagOf(el);
      if (shouldDrop(el)) continue;

      if (/^h[1-6]$/.test(tag)) {
        flush();
        var t = inlineText(childNodes(el));
        if (t) blocks.push(new Array(+tag[1] + 1).join("#") + " " + t);
      } else if (tag === "p") {
        flush();
        var p = inlineText(childNodes(el));
        if (p) blocks.push(p);
      } else if (tag === "table") {
        flush();
        var tb = renderTableOrPromote(el);
        if (tb) blocks.push(tb);
      } else if (tag === "ul" || tag === "ol") {
        flush();
        var ls = renderList(el, tag === "ol");
        if (ls) blocks.push(ls);
      } else if (tag === "hr") {
        flush();
        blocks.push("---");
      } else if (TRANSPARENT.has(tag)) {
        if (hasBlockChild(el)) { flush(); var inner = renderBlocks(childNodes(el)); if (inner) blocks.push(inner); }
        else buf.push(el);
      } else if (BLOCK_IN_CELL.has(tag)) {
        flush();
        var bb = renderBlocks(childNodes(el));
        if (bb) blocks.push(bb);
      } else {
        buf.push(el);            // inline element (b, i, a, ...) at block level
      }
    }
    flush();
    return blocks.join("\n\n");
  }

  // --- Public pipeline ------------------------------------------------------

  // Trailing boilerplate every ClearTax report ends with (signature block and
  // "Generated by ClearTax" credit) — no value to an AI, so drop it.
  var FOOTER_RE = [/^signature$/i, /^for .{1,60}$/i, /^generated by cleartax\b/i];

  function stripFooter(md) {
    var lines = md.split("\n");
    var isBoiler = function (l) {
      l = l.trim();
      return l !== "" && FOOTER_RE.some(function (re) { return re.test(l); });
    };
    var changed = true;
    while (changed) {
      changed = false;
      while (lines.length && lines[lines.length - 1].trim() === "") { lines.pop(); changed = true; }
      if (lines.length && isBoiler(lines[lines.length - 1])) { lines.pop(); changed = true; }
    }
    return lines.join("\n");
  }

  function htmlToMarkdown(html) {
    var doc = new DOMParser().parseFromString(html, "text/html");
    var md = renderBlocks(childNodes(doc.body));
    md = md.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
    md = stripFooter(md);
    return md + "\n";
  }

  function decodeBytes(buf) {
    var bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder("windows-1252").decode(bytes);
    }
  }

  function convertBytes(buf) {
    return htmlToMarkdown(decodeBytes(buf));
  }

  global.tax2md = {
    decodeBytes: decodeBytes,
    htmlToMarkdown: htmlToMarkdown,
    convertBytes: convertBytes,
  };
})(typeof window !== "undefined" ? window : this);
