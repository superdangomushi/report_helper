#!/usr/bin/env node
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import JSZip from "jszip";
import { imageSize } from "image-size";

const __filename = fileURLToPath(import.meta.url);
const SCRIPT_DIR = path.dirname(__filename);
const CWD = process.cwd();
const TEMPLATES_DIR = path.join(SCRIPT_DIR, "templates");
const USER_CONFIG_DIR = path.join(os.homedir(), ".mkreport");
const USER_CONFIG_FILE = path.join(USER_CONFIG_DIR, "config.env");

// ---------- CLI ----------
function printUsageAndExit(code = 1) {
  console.error(
    [
      "Usage:",
      "  mkreport setup                          Open a GUI to set the Gemini API key (saved per-user)",
      "  mkreport start <project-name>           Scaffold a new project from templates/",
      "  mkreport [--ocr] <project-path>         Build report from a project folder",
      "  mkreport [--ocr] <project-path> <out>   Build to a specific output path",
      "",
      "Flags:",
      "  --ocr   Run Gemini OCR on ocr_hyo/ images. Refreshes any existing",
      "          <OCR-...>…</OCR-...> blocks too. Without this flag, OCR fences",
      "          render as a placeholder caption and cached OCR blocks render as-is.",
      "",
      "Examples:",
      "  mkreport setup",
      "  mkreport start my_first_report",
      "  mkreport ./my_first_report",
      "  mkreport --ocr ./my_first_report report.docx",
    ].join("\n")
  );
  process.exit(code);
}

const FLAGS = { ocr: false };
const args = [];
for (const a of process.argv.slice(2)) {
  if (a === "--ocr") FLAGS.ocr = true;
  else if (a === "--help" || a === "-h") args.push(a);
  else if (a.startsWith("--")) {
    console.error(`error: unknown flag ${a}`);
    printUsageAndExit();
  } else {
    args.push(a);
  }
}
if (args.length < 1) printUsageAndExit();

// PROJECT_ROOT is set by buildReport(); helper functions use it.
let PROJECT_ROOT = CWD;

// ---------- helpers ----------
const xmlEscape = (s) =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

// Report convention: full-width Japanese commas/periods must be ASCII.
const normalizePunctuation = (s) =>
  String(s ?? "").replace(/、/g, ",").replace(/。/g, ".");

// Look up an input file. Lookup order:
//   1. PROJECT_ROOT       — project-specific override (highest priority)
//   2. USER_CONFIG_DIR    — per-user shared (~/.mkreport/, populated by `mkreport setup`)
//   3. TEMPLATES_DIR      — bundled defaults
// This lets users keep one customized seed.docx / hyo1.docx in ~/.mkreport/
// and have it picked up by every project automatically.
function findInputFile(name) {
  const candidates = [
    path.join(PROJECT_ROOT, name),
    path.join(USER_CONFIG_DIR, name),
    path.join(TEMPLATES_DIR, name),
  ];
  for (const p of candidates) if (fs.existsSync(p)) return p;
  return null;
}

// Look up an asset file under a subfolder (figures/, ocr_hyo/).
function findAssetFile(name, folder) {
  const candidates = [
    path.join(PROJECT_ROOT, folder, name),
    path.join(TEMPLATES_DIR, folder, name),
  ];
  for (const p of candidates) if (fs.existsSync(p)) return p;
  return null;
}

function readTextIfExists(p) {
  if (!p || !fs.existsSync(p)) return "";
  return fs.readFileSync(p, "utf8").replace(/^﻿/, "");
}

// Minimal .env reader (KEY=VALUE per line). Strips surrounding quotes.
// Lookup order: process env → project .env → per-user config (mkreport setup)
// → bundled templates/.env → script dir .env.
function readEnvVar(key) {
  if (process.env[key]) return process.env[key];
  const candidates = [
    path.join(PROJECT_ROOT, ".env"),
    USER_CONFIG_FILE,
    path.join(TEMPLATES_DIR, ".env"),
    path.join(SCRIPT_DIR, ".env"),
  ];
  for (const p of candidates) {
    if (!fs.existsSync(p)) continue;
    const text = fs.readFileSync(p, "utf8");
    const re = new RegExp(`^\\s*${key}\\s*=\\s*(.+?)\\s*$`, "m");
    const m = text.match(re);
    if (m) return m[1].replace(/^["']|["']$/g, "");
  }
  return null;
}

// ---------- title.env parser (KEY=VALUE, supports multi-line via backslash) ----------
function parseTitleEnv(text) {
  const map = {};
  if (!text) return map;
  const lines = text.split(/\r?\n/);
  for (const raw of lines) {
    const line = raw.trim();
    if (!line || line.startsWith("#") && !/^#[0-9T]/.test(line.replace(/^\s+/, ""))) {
      // skip pure comments (lines starting with # but not a placeholder key)
      if (!/^#[0-9T]/.test(line)) continue;
    }
    const m = line.match(/^(#(?:T\d+|\d+))\s*=\s*(.*)$/);
    if (m) map[m[1]] = m[2];
  }
  return map;
}

// ---------- alignment preprocessing ----------
// Markers (C++ block-comment style): /c ... c/ → center, /r ... r/ → right.
// Default alignment is left. Markers may span paragraphs.
// Returns chunks: [{align, text}, ...]
function splitByAlignment(text) {
  if (!text) return [{ align: "left", text: "" }];
  const chunks = [];
  let pos = 0;
  let current = "left";
  // /c …c/ → center, /r …r/ → right. The lookarounds keep table cell-merge
  // markers (<-/c … c/->) from being mistaken for alignment markers, which
  // would otherwise split a table row apart before it reaches the parser.
  const re = /(?<!<-)\/([cr])|([cr])\/(?!->)/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > pos) chunks.push({ align: current, text: text.slice(pos, m.index) });
    if (m[1]) current = m[1] === "c" ? "center" : "right";
    else current = "left";
    pos = m.index + m[0].length;
  }
  if (pos < text.length) chunks.push({ align: current, text: text.slice(pos) });
  return chunks;
}

// ---------- markdown parser ----------
// Produces blocks: {type:'p',text,align?} | {type:'table',rows} | {type:'fence',content}
function parseMarkdown(text) {
  const blocks = [];
  const lines = text.split(/\r?\n/);
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }

    // code fence
    if (/^\s*```/.test(line)) {
      // Single-line fence: ```content```
      const inline = line.match(/^\s*```(.*?)```\s*$/);
      if (inline) {
        blocks.push({ type: "fence", content: inline[1].trim() });
        i++;
        continue;
      }
      // Multi-line fence
      const buf = [];
      // strip any language tag after opening ```
      const firstLine = line.replace(/^\s*```\S*\s*/, "");
      if (firstLine) buf.push(firstLine);
      i++;
      while (i < lines.length && !/^\s*```/.test(lines[i])) {
        buf.push(lines[i]);
        i++;
      }
      i++; // consume closing fence
      blocks.push({ type: "fence", content: buf.join("\n").trim() });
      continue;
    }

    // table: header row + separator row + body rows
    if (/^\s*\|.*\|\s*$/.test(line) && i + 1 < lines.length && /^\s*\|[-:\s|]+\|\s*$/.test(lines[i + 1])) {
      const parseRow = (l) =>
        l.trim().replace(/^\|/, "").replace(/\|\s*$/, "").split("|").map((c) => c.trim());
      const header = parseRow(line);
      i += 2; // skip header + separator
      const rows = [header];
      while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) {
        rows.push(parseRow(lines[i]));
        i++;
      }
      blocks.push({ type: "table", rows });
      continue;
    }

    // paragraph: accumulate until blank
    const buf = [line];
    i++;
    while (i < lines.length && lines[i].trim() && !/^\s*```/.test(lines[i]) && !/^\s*\|.*\|\s*$/.test(lines[i])) {
      buf.push(lines[i]);
      i++;
    }
    blocks.push({ type: "p", text: buf.join("\n") });
  }
  return blocks;
}

// ---------- placeholder substitution in document.xml ----------
// Keys that appear multiple times in the seed (e.g. #4 = 氏名 + 班番号) can be
// filled from one CSV value: occurrence N gets CSV part N (clamped to the last).
// Single-occurrence keys keep their value verbatim so commas inside e.g. #T1 stay intact.
function substitutePlaceholders(xml, map) {
  const re = /#(T\d+|\d+)(?!\d)/g;
  const counts = {};
  for (const m of xml.matchAll(re)) {
    const k = "#" + m[1];
    counts[k] = (counts[k] || 0) + 1;
  }
  const parts = {};
  const cursor = {};
  for (const k of Object.keys(map)) {
    if (counts[k] > 1 && /[,、]/.test(map[k])) {
      parts[k] = map[k].split(/[,、]/).map((s) => s.trim());
      cursor[k] = 0;
    }
  }
  return xml.replace(re, (full, key) => {
    const k = "#" + key;
    if (!(k in map)) return full;
    let value;
    if (parts[k]) {
      const i = Math.min(cursor[k], parts[k].length - 1);
      value = parts[k][i];
      cursor[k]++;
    } else {
      value = map[k];
    }
    return xmlEscape(normalizePunctuation(value));
  });
}

// ---------- XML builders for body content ----------
const W_NS = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
const M_NS = "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"";

// Leading whitespace becomes a nesting level: one tab, or SPACES_PER_LEVEL
// half-width spaces, per level. A literal tab/space only indents the first
// visual line, so wrapped lines fall back to the margin; converting the indent
// to a paragraph-level left indent keeps every wrapped line aligned.
const INDENT_PER_LEVEL = 840; // twips, matches the document's default tab stop
const SPACES_PER_LEVEL = 4;   // four leading half-width spaces == one nesting level
const BULLET_HANGING = 420;   // hang width so wrapped bullet lines align under the text
function stripLeadingIndent(line) {
  let i = 0, tabs = 0, spaces = 0;
  while (i < line.length && (line[i] === "\t" || line[i] === " ")) {
    if (line[i] === "\t") tabs++; else spaces++;
    i++;
  }
  // Full-width spaces (　, used for Japanese first-line indent) are intentionally
  // left in `rest`; only tabs and half-width spaces count toward nesting.
  const level = tabs + Math.floor(spaces / SPACES_PER_LEVEL);
  return { level, rest: line.slice(i) };
}

// Classify a source line: a leading "- " / "* " (after any indent) marks a bullet.
function classifyLine(line) {
  const { level, rest } = stripLeadingIndent(line);
  const m = rest.match(/^[-*]\s+(.*)$/);
  if (m) return { level, kind: "bullet", content: m[1] };
  return { level, kind: "text", content: rest };
}

// ---------- inline markup → runs ----------
// Renders the inline syntax that can appear inside a paragraph or table cell:
//   $...$     → an inline Word equation (OMML), reusing the math parser
//   <br>      → a hard line break (also <br/> and <br />)
// Everything else becomes plain text runs (punctuation normalized + XML-escaped).
// This is what lets complex Markdown tables (math headers, multi-line cells)
// survive conversion instead of leaking raw LaTeX / "<br>" into the document.
function textRuns(text) {
  // Split on <br> variants; each gap becomes a <w:br/> line break.
  const parts = String(text ?? "").split(/<br\s*\/?>/i);
  return parts
    .map((part, idx) => {
      const run =
        part === ""
          ? ""
          : `<w:r><w:t xml:space="preserve">${xmlEscape(normalizePunctuation(part))}</w:t></w:r>`;
      return idx < parts.length - 1 ? `${run}<w:r><w:br/></w:r>` : run;
    })
    .join("");
}

function renderInline(text) {
  const s = String(text ?? "");
  // $...$ inline math (single line, non-greedy; needs at least one inner char).
  const re = /\$([^$\n]+?)\$/g;
  let out = "";
  let last = 0;
  let m;
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) out += textRuns(s.slice(last, m.index));
    out += `<m:oMath>${mathToOMML(m[1])}</m:oMath>`;
    last = m.index + m[0].length;
  }
  if (last < s.length) out += textRuns(s.slice(last));
  return out;
}

function pPlain(text, align = "left", opts = {}) {
  // paragraph with simple text; preserve leading/trailing spaces
  if (text === "") return `<w:p/>`;
  const keep = (opts.keepNext ? `<w:keepNext/>` : "") + (opts.keepLines ? `<w:keepLines/>` : "");
  const jc =
    align === "center" ? `<w:jc w:val="center"/>` :
    align === "right"  ? `<w:jc w:val="right"/>`  : "";

  const lines = String(text).split(/\n/).map(classifyLine);
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const cur = lines[i];
    if (cur.kind === "bullet") {
      // Hanging indent: bullet glyph at level*840, text (and wrapped lines) at +420.
      const left = cur.level * INDENT_PER_LEVEL + BULLET_HANGING;
      const ind = `<w:ind w:left="${left}" w:hanging="${BULLET_HANGING}"/>`;
      out.push(
        `<w:p><w:pPr>${keep}${ind}${jc}</w:pPr>` +
          `<w:r><w:t>・</w:t></w:r><w:r><w:tab/></w:r>` +
          renderInline(cur.content) +
          `</w:p>`
      );
      i++;
      continue;
    }
    // Group consecutive plain-text lines at the same indent level into one
    // paragraph (joined by soft breaks) so wrapped lines keep the indent.
    const level = cur.level;
    const items = [];
    while (i < lines.length && lines[i].kind === "text" && lines[i].level === level) {
      items.push(lines[i].content);
      i++;
    }
    const ind = level > 0 ? `<w:ind w:left="${level * INDENT_PER_LEVEL}"/>` : "";
    const pPr = (keep || ind || jc) ? `<w:pPr>${keep}${ind}${jc}</w:pPr>` : "";
    const runs = items
      .map((line, idx, arr) => {
        const t = renderInline(line);
        return idx < arr.length - 1 ? `${t}<w:r><w:br/></w:r>` : t;
      })
      .join("");
    out.push(`<w:p>${pPr}${runs}</w:p>`);
  }
  return out.join("");
}

function pCenter(text, opts = {}) {
  const keep = (opts.keepNext ? `<w:keepNext/>` : "") + (opts.keepLines ? `<w:keepLines/>` : "");
  return `<w:p><w:pPr>${keep}<w:jc w:val="center"/></w:pPr>${renderInline(text)}</w:p>`;
}

// ---------- cell-merge markers (Word-like spanned cells) ----------
// Horizontal merge: a cell beginning with "<-" opens a span and a cell ending
// with "->" closes it; the cells in between collapse into one Word cell via
// w:gridSpan. An optional /c, /l, /r (either side of the arrow) sets the
// alignment, defaulting to center — e.g. "<-/c理論値 | c/->" spans two columns,
// centered, showing "理論値".
// Vertical merge: a cell ending with "-^" toggles a vertical span in its column.
// The first "-^" cell opens the span (its text stays at the top), the next
// closes it, and every cell between them joins via w:vMerge — e.g. "395-^" …
// "" … "-^" merges three rows into one showing "395".
function alignFromMarker(mk) {
  if (!mk) return null;
  const c = mk.toLowerCase();
  if (c.includes("c")) return "center";
  if (c.includes("l")) return "left";
  if (c.includes("r")) return "right";
  return null;
}
function parseHMergeStart(s) {
  const m = /^<-\s*(\/?[clrCLR]\/?)?\s*([\s\S]*)$/.exec(s);
  if (!m) return null;
  return { align: alignFromMarker(m[1]), content: m[2].trim() };
}
function parseHMergeEnd(s) {
  const m = /^([\s\S]*?)\s*(\/?[clrCLR]\/?)?->\s*$/.exec(s);
  if (!m) return null;
  return { align: alignFromMarker(m[2]), content: m[1].trim() };
}
function parseVMerge(s) {
  const m = /^([\s\S]*?)\s*-\^\s*$/.exec(s);
  if (!m) return null;
  return { content: m[1].trim() };
}

// Resolve the cell-merge markers into a grid model shared by the Word table
// builder and the Excel export. Returns { colCount, R, M } where each
// M[r][c] = { content, align, gridSpan, vMerge:'restart'|'continue'|null, hSkip }.
function resolveMergedCells(rows) {
  const colCount = Math.max(...rows.map((r) => r.length));
  const R = rows.length;

  const M = rows.map((row) =>
    Array.from({ length: colCount }, (_, c) => ({
      content: row[c] ?? "",
      align: "left",
      gridSpan: 1,
      vMerge: null, // 'restart' | 'continue'
      hSkip: false, // covered by a gridSpan cell to its left → not emitted
    }))
  );

  // Horizontal merges (per row): "<-" opens, "->" closes, gridSpan collapses.
  for (let r = 0; r < R; r++) {
    let c = 0;
    while (c < colCount) {
      const cell = M[r][c];
      const start = parseHMergeStart(cell.content);
      if (!start) { c++; continue; }
      let content = start.content;
      let align = start.align || "center";
      // The opening cell can also carry the closing arrow (e.g. "<-/c理論値->").
      const self = parseHMergeEnd(content);
      if (self) {
        cell.content = self.content;
        cell.align = align;
        c++;
        continue;
      }
      let end = -1;
      for (let k = c + 1; k < colCount; k++) {
        const info = parseHMergeEnd(M[r][k].content);
        if (info) {
          end = k;
          if (!content) content = info.content;
          if (!start.align && info.align) align = info.align;
          break;
        }
      }
      cell.content = content;
      cell.align = align;
      if (end > c) {
        cell.gridSpan = end - c + 1;
        for (let k = c + 1; k <= end; k++) M[r][k].hSkip = true;
        c = end + 1;
      } else {
        c++; // lone "<-": no closer found, keep it as a single centered cell
      }
    }
  }

  // Vertical merges (per column): paired "-^" markers fold rows together.
  for (let c = 0; c < colCount; c++) {
    let open = false;
    for (let r = 0; r < R; r++) {
      const cell = M[r][c];
      if (cell.hSkip) continue;
      const v = parseVMerge(cell.content);
      if (v) {
        if (!open) {
          cell.vMerge = "restart";
          cell.content = v.content;
          open = true;
        } else {
          cell.vMerge = "continue";
          cell.content = "";
          open = false;
        }
      } else if (open) {
        cell.vMerge = "continue";
        cell.content = "";
      }
    }
  }

  return { colCount, R, M };
}

function tableXml(rows) {
  // Apply hyo1.docx style: tblStyle "aa", double top border, no left/right; rows distributed evenly
  const { colCount, R, M } = resolveMergedCells(rows);
  // total width ~9000 dxa
  const colWidth = Math.floor(9000 / colCount);
  const gridCols = Array.from({ length: colCount }, () => `<w:gridCol w:w="${colWidth}"/>`).join("");

  const trXml = M
    .map((row, rIdx) => {
      // keepNext on every row but the last glues the rows together, so a table
      // that fits on one page jumps to the next page intact rather than splitting.
      const keepNext = rIdx < R - 1;
      const tcs = [];
      for (let c = 0; c < colCount; c++) {
        const cell = row[c];
        if (cell.hSkip) continue;
        const w = colWidth * cell.gridSpan;
        const tcPr =
          `<w:tcW w:w="${w}" w:type="dxa"/>` +
          (cell.gridSpan > 1 ? `<w:gridSpan w:val="${cell.gridSpan}"/>` : "") +
          (cell.vMerge === "restart" ? `<w:vMerge w:val="restart"/>` :
           cell.vMerge === "continue" ? `<w:vMerge/>` : "") +
          (cell.align === "center" ? `<w:vAlign w:val="center"/>` : "");
        // A vMerge continuation row contributes an empty cell body.
        const body =
          cell.vMerge === "continue"
            ? `<w:p/>`
            : pPlain(cell.content, cell.align, { keepNext });
        tcs.push(`<w:tc><w:tcPr>${tcPr}</w:tcPr>${body}</w:tc>`);
      }
      // cantSplit stops a single row from straddling a page boundary.
      return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${tcs.join("")}</w:tr>`;
    })
    .join("");

  return (
    `<w:tbl>` +
    `<w:tblPr>` +
    `<w:tblStyle w:val="aa"/>` +
    `<w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="double" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
    `<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>` +
    `</w:tblBorders>` +
    `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>` +
    `</w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>` +
    trXml +
    `</w:tbl>` +
    `<w:p/>` // spacer paragraph after table (Word requires a paragraph after a table)
  );
}

// ---------- Excel (.xlsx) export ----------
// One sheet per table, same content as the Word table including cell merges.
// An .xlsx is just a zipped OOXML package (like .docx), so we build it by hand
// with JSZip rather than pulling in a spreadsheet dependency.

// 0-based column index → spreadsheet letters (0→A, 25→Z, 26→AA, ...).
function colLetter(n) {
  let s = "";
  n += 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Excel sheet names: ≤31 chars, may not contain : \ / ? * [ ], and must be
// unique within the workbook. `used` is a Set the caller carries across sheets.
function sanitizeSheetName(name, fallback, used) {
  let base = String(name ?? "").replace(/[:\\/?*\[\]]/g, " ").replace(/\s+/g, " ").trim();
  if (!base) base = fallback;
  base = base.slice(0, 31);
  let out = base;
  let n = 2;
  while (used.has(out)) {
    const suffix = `_${n++}`;
    out = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(out);
  return out;
}

function sheetXml(rows) {
  const { colCount, R, M } = resolveMergedCells(rows);
  const isNumeric = (s) => /^-?\d+(?:\.\d+)?$/.test(s);
  const merges = [];
  const rowXmls = [];
  for (let r = 0; r < R; r++) {
    const cells = [];
    for (let c = 0; c < colCount; c++) {
      const cell = M[r][c];
      // hSkip cells are covered by a gridSpan; vMerge continuations are blank
      // bodies folded into the cell above — neither carries its own value.
      if (cell.hSkip || cell.vMerge === "continue") continue;
      // Compute the merge rectangle: gridSpan columns × contiguous vMerge rows.
      let rowSpan = 1;
      if (cell.vMerge === "restart") {
        while (r + rowSpan < R && M[r + rowSpan][c].vMerge === "continue") rowSpan++;
      }
      const ref = `${colLetter(c)}${r + 1}`;
      if (cell.gridSpan > 1 || rowSpan > 1) {
        merges.push(`${ref}:${colLetter(c + cell.gridSpan - 1)}${r + rowSpan}`);
      }
      const v = cell.content;
      if (v === "") continue;
      if (isNumeric(v)) {
        cells.push(`<c r="${ref}"><v>${v}</v></c>`);
      } else {
        cells.push(`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(normalizePunctuation(v))}</t></is></c>`);
      }
    }
    rowXmls.push(`<row r="${r + 1}">${cells.join("")}</row>`);
  }
  const mergeXml = merges.length
    ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>`
    : "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<sheetData>${rowXmls.join("")}</sheetData>` +
    mergeXml + // schema order: sheetData then mergeCells
    `</worksheet>`
  );
}

// Build a workbook buffer from [{ name, rows }, ...]. Names are pre-sanitized.
async function buildXlsx(sheets) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheets
        .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
        .join("") +
      `</Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets>` +
      sheets.map((s, i) => `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
      `</sheets></workbook>`
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheets
        .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
        .join("") +
      `</Relationships>`
  );
  sheets.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows)));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ---------- Math parser (LaTeX-like → OMML) ----------
// Supports: _ (subscript), ^ (superscript), \frac{a}{b}, \sqrt{x}, \sqrt[n]{x},
//           * → × (multiplication sign), Greek letters, common math commands.
const MATH_SYMBOLS = {
  // lowercase Greek
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ϵ",
  zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ", iota: "ι", kappa: "κ",
  lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", omicron: "ο", pi: "π", varpi: "ϖ",
  rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ",
  phi: "φ", varphi: "ϕ", chi: "χ", psi: "ψ", omega: "ω",
  // uppercase Greek
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
  Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  // operators / relations
  times: "×", cdot: "·", div: "÷", pm: "±", mp: "∓",
  le: "≤", leq: "≤", ge: "≥", geq: "≥", ne: "≠", neq: "≠",
  approx: "≈", equiv: "≡", sim: "∼", simeq: "≃", cong: "≅", propto: "∝",
  ll: "≪", gg: "≫",
  // arrows
  to: "→", rightarrow: "→", leftarrow: "←", leftrightarrow: "↔",
  Rightarrow: "⇒", Leftarrow: "⇐", Leftrightarrow: "⇔", mapsto: "↦",
  // set / logic
  in: "∈", notin: "∉", subset: "⊂", supset: "⊃", subseteq: "⊆", supseteq: "⊇",
  cup: "∪", cap: "∩", emptyset: "∅", forall: "∀", exists: "∃",
  land: "∧", lor: "∨", lnot: "¬",
  // misc symbols
  infty: "∞", partial: "∂", nabla: "∇", hbar: "ℏ", ell: "ℓ",
  prime: "′", dprime: "″",
  cdots: "⋯", ldots: "…", vdots: "⋮", ddots: "⋱",
  // big operators (rendered as plain symbols here; LaTeX-style limits not implemented)
  sum: "∑", prod: "∏", int: "∫", oint: "∮", iint: "∬", iiint: "∭",
  lim: "lim",
  // math functions
  sin: "sin", cos: "cos", tan: "tan", cot: "cot", sec: "sec", csc: "csc",
  arcsin: "arcsin", arccos: "arccos", arctan: "arctan",
  sinh: "sinh", cosh: "cosh", tanh: "tanh",
  log: "log", ln: "ln", exp: "exp",
  max: "max", min: "min",
};

// Accent commands (LaTeX → combining unicode char used in <m:acc>)
const MATH_ACCENTS = {
  bar: "̄", overline: "̄",
  hat: "̂", widehat: "̂",
  tilde: "̃", widetilde: "̃",
  vec: "⃗", overrightarrow: "⃗",
  dot: "̇", ddot: "̈",
  acute: "́", grave: "̀",
  check: "̌", breve: "̆",
  mathring: "̊",
};

function tokenizeMath(s) {
  const tokens = [];
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (c === "\n") { tokens.push({ type: "newline" }); i++; continue; }
    if (/\s/.test(c)) { i++; continue; }
    if (c === "\\") {
      let j = i + 1;
      // command name = letters; if next char is non-letter, take it as a one-char command
      if (j < s.length && /[a-zA-Z]/.test(s[j])) {
        while (j < s.length && /[a-zA-Z]/.test(s[j])) j++;
        tokens.push({ type: "cmd", value: s.slice(i + 1, j) });
      } else {
        tokens.push({ type: "cmd", value: s[j] || "" });
        j = i + 2;
      }
      i = j; continue;
    }
    if (c === "{") { tokens.push({ type: "lbrace" }); i++; continue; }
    if (c === "}") { tokens.push({ type: "rbrace" }); i++; continue; }
    if (c === "[") { tokens.push({ type: "lbracket" }); i++; continue; }
    if (c === "]") { tokens.push({ type: "rbracket" }); i++; continue; }
    if (c === "_" || c === "^") { tokens.push({ type: c }); i++; continue; }
    if (c === "*") { tokens.push({ type: "sym", value: "×" }); i++; continue; }
    if (/[0-9]/.test(c)) {
      let j = i;
      while (j < s.length && /[0-9.]/.test(s[j])) j++;
      tokens.push({ type: "num", value: s.slice(i, j) });
      i = j; continue;
    }
    if (/[a-zA-Z]/.test(c)) {
      tokens.push({ type: "id", value: c });
      i++; continue;
    }
    tokens.push({ type: "sym", value: c });
    i++;
  }
  return tokens;
}

// Parser: returns an array of nodes for a "sequence" until end or matching rbrace.
function parseMathSeq(tokens, i, stopOn) {
  const out = [];
  while (i < tokens.length) {
    const t = tokens[i];
    if (stopOn && t.type === stopOn) break;
    if (t.type === "newline") { out.push({ type: "newline" }); i++; continue; }
    const r = parseMathAtom(tokens, i);
    out.push(r.node);
    i = r.next;
  }
  return { node: out, next: i };
}

function parseMathGroup(tokens, i) {
  // { ... } group; or a single atom if no brace
  if (i < tokens.length && tokens[i].type === "lbrace") {
    const r = parseMathSeq(tokens, i + 1, "rbrace");
    return { node: { type: "group", items: r.node }, next: r.next + 1 };
  }
  return parseMathPrimary(tokens, i);
}

// Sub/sup argument: gobble consecutive ids or digits without braces.
function parseSubSupArg(tokens, i) {
  if (i >= tokens.length) return { node: { type: "text", value: "" }, next: i };
  const t = tokens[i];
  if (t.type === "lbrace") return parseMathGroup(tokens, i);
  if (t.type === "id") {
    let j = i;
    while (j < tokens.length && tokens[j].type === "id") j++;
    if (j - i > 1) {
      return { node: { type: "text", value: tokens.slice(i, j).map((x) => x.value).join("") }, next: j };
    }
    return parseMathPrimary(tokens, i);
  }
  if (t.type === "num") {
    return parseMathPrimary(tokens, i);
  }
  return parseMathPrimary(tokens, i);
}

// Parse a primary (without sub/sup attachment)
function parseMathPrimary(tokens, i) {
  const t = tokens[i];
  if (!t) return { node: { type: "text", value: "" }, next: i };
  if (t.type === "lbrace") return parseMathGroup(tokens, i);
  if (t.type === "id") return { node: { type: "var", value: t.value }, next: i + 1 };
  if (t.type === "num") return { node: { type: "num", value: t.value }, next: i + 1 };
  if (t.type === "sym") return { node: { type: "sym", value: t.value }, next: i + 1 };
  if (t.type === "cmd") {
    if (t.value === "frac") {
      const num = parseMathGroup(tokens, i + 1);
      const den = parseMathGroup(tokens, num.next);
      return { node: { type: "frac", num: num.node, den: den.node }, next: den.next };
    }
    if (t.value === "sqrt") {
      let degree = null;
      let j = i + 1;
      if (tokens[j] && tokens[j].type === "lbracket") {
        const d = parseMathSeq(tokens, j + 1, "rbracket");
        degree = { type: "group", items: d.node };
        j = d.next + 1;
      }
      const arg = parseMathGroup(tokens, j);
      return { node: { type: "sqrt", degree, arg: arg.node }, next: arg.next };
    }
    if (MATH_ACCENTS[t.value]) {
      const arg = parseMathGroup(tokens, i + 1);
      return { node: { type: "accent", char: MATH_ACCENTS[t.value], arg: arg.node }, next: arg.next };
    }
    if (MATH_SYMBOLS[t.value]) {
      // Function names like sin/cos/log etc: emit as plain text (Word will keep upright)
      return { node: { type: "func", value: MATH_SYMBOLS[t.value], isWord: /^[a-zA-Z]+$/.test(MATH_SYMBOLS[t.value]) && MATH_SYMBOLS[t.value].length > 1 }, next: i + 1 };
    }
    // Unknown command: emit as plain word
    return { node: { type: "text", value: "\\" + t.value }, next: i + 1 };
  }
  return { node: { type: "text", value: "?" }, next: i + 1 };
}

// Parse one atom = primary + optional _sub ^sup
function parseMathAtom(tokens, i) {
  let prim = parseMathPrimary(tokens, i);
  let cur = prim.node;
  i = prim.next;
  let sub = null, sup = null;
  while (i < tokens.length && (tokens[i].type === "_" || tokens[i].type === "^")) {
    const which = tokens[i].type;
    i++;
    const arg = parseSubSupArg(tokens, i);
    i = arg.next;
    if (which === "_") sub = arg.node;
    else sup = arg.node;
  }
  if (sub && sup) cur = { type: "subsup", base: cur, sub, sup };
  else if (sub) cur = { type: "sub", base: cur, sub };
  else if (sup) cur = { type: "sup", base: cur, sup };
  return { node: cur, next: i };
}

// AST → OMML XML fragment (without enclosing <m:oMath>)
function mathNodeToOMML(node) {
  if (Array.isArray(node)) return node.map(mathNodeToOMML).join("");
  switch (node.type) {
    case "group":
      return mathNodeToOMML(node.items);
    case "var":
      return `<m:r><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
    case "num":
      return `<m:r><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
    case "sym":
      return `<m:r><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
    case "text":
      return `<m:r><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
    case "func":
      // Use upright style for multi-char function names (sin, cos, log, etc.)
      if (node.isWord) {
        return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
      }
      return `<m:r><m:t>${xmlEscape(node.value)}</m:t></m:r>`;
    case "newline":
      return ""; // ignore inside math; line breaks handled by caller
    case "sub":
      return `<m:sSub><m:e>${mathNodeToOMML(node.base)}</m:e><m:sub>${mathNodeToOMML(node.sub)}</m:sub></m:sSub>`;
    case "sup":
      return `<m:sSup><m:e>${mathNodeToOMML(node.base)}</m:e><m:sup>${mathNodeToOMML(node.sup)}</m:sup></m:sSup>`;
    case "subsup":
      return `<m:sSubSup><m:e>${mathNodeToOMML(node.base)}</m:e><m:sub>${mathNodeToOMML(node.sub)}</m:sub><m:sup>${mathNodeToOMML(node.sup)}</m:sup></m:sSubSup>`;
    case "frac":
      return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${mathNodeToOMML(node.num)}</m:num><m:den>${mathNodeToOMML(node.den)}</m:den></m:f>`;
    case "sqrt":
      if (node.degree) {
        return `<m:rad><m:deg>${mathNodeToOMML(node.degree)}</m:deg><m:e>${mathNodeToOMML(node.arg)}</m:e></m:rad>`;
      }
      return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${mathNodeToOMML(node.arg)}</m:e></m:rad>`;
    case "accent":
      return `<m:acc><m:accPr><m:chr m:val="${xmlEscape(node.char)}"/></m:accPr><m:e>${mathNodeToOMML(node.arg)}</m:e></m:acc>`;
  }
  return "";
}

function mathToOMML(content) {
  // Multi-line: render each line in its own <m:oMath> separated by line breaks
  const lines = content.split(/\n/);
  return lines
    .map((line) => {
      const tokens = tokenizeMath(line);
      const ast = parseMathSeq(tokens, 0).node;
      return mathNodeToOMML(ast);
    })
    .join(`<m:r><w:br/></m:r>`);
}

function equationXml(content, eqnNumber) {
  // center-aligned via tabs: tab to center (equation) + tab to right ((N))
  const mathOmml = mathToOMML(content);
  return (
    `<w:p>` +
    `<w:pPr>` +
    `<w:tabs>` +
    `<w:tab w:val="center" w:pos="4536"/>` +
    `<w:tab w:val="right" w:pos="9072"/>` +
    `</w:tabs>` +
    `</w:pPr>` +
    `<w:r><w:tab/></w:r>` +
    `<m:oMath>${mathOmml}</m:oMath>` +
    `<w:r><w:tab/><w:t xml:space="preserve">(${eqnNumber})</w:t></w:r>` +
    `</w:p>`
  );
}

// ---------- Gemini OCR (table extraction) ----------
// Returns { rows } on success, or { error } on any failure (missing key, network error, parse failure).
async function geminiOcrTable(imagePath) {
  const apiKey = readEnvVar("GEMINI_API_KEY");
  if (!apiKey) return { error: "GEMINI_API_KEY が未設定です (`mkreport setup` で設定するか、.env に GEMINI_API_KEY=... を書いてください)" };
  try {
    const buf = fs.readFileSync(imagePath);
    const ext = path.extname(imagePath).slice(1).toLowerCase();
    const mimeType =
      ext === "png" ? "image/png" :
      ext === "gif" ? "image/gif" :
      ext === "webp" ? "image/webp" : "image/jpeg";
    const body = {
      contents: [{
        parts: [
          { text: "この画像に写っている表を、Markdown 表記（| 列1 | 列2 |\\n|---|---|\\n| 値 | 値 | 形式）でそのまま出力してください。表の内容のみを出力し、説明文や前置きは一切不要です。" },
          { inline_data: { mime_type: mimeType, data: buf.toString("base64") } },
        ],
      }],
    };
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const errText = await res.text().catch(() => "");
      let apiMsg = "";
      try { apiMsg = JSON.parse(errText)?.error?.message || ""; } catch {}
      const detail = apiMsg || errText.slice(0, 300);
      return { error: `Gemini API エラー (HTTP ${res.status} ${res.statusText})${detail ? `: ${detail}` : ""}` };
    }
    const json = await res.json();
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) {
      const blockReason = json?.promptFeedback?.blockReason;
      const finishReason = json?.candidates?.[0]?.finishReason;
      const reason = blockReason ? `ブロック理由: ${blockReason}` :
                     finishReason ? `finishReason: ${finishReason}` :
                     "応答にテキストが含まれていません";
      return { error: `Gemini 応答が空 (${reason})` };
    }
    // Strip any ```markdown ... ``` wrappers Gemini might add.
    const cleaned = text.replace(/^```\w*\s*\n?|\n?```\s*$/g, "");
    const blocks = parseMarkdown(cleaned);
    const tbl = blocks.find((b) => b.type === "table");
    if (!tbl) return { error: `Markdown 表が検出できません (応答冒頭: ${cleaned.slice(0, 200).replace(/\s+/g, " ")})` };
    if (!tbl.rows || tbl.rows.length === 0) return { error: "Markdown 表は検出されましたが行が空です" };
    return { rows: tbl.rows };
  } catch (e) {
    return { error: `例外: ${e?.message || String(e)}` };
  }
}

// ---------- OCR block markers (<OCR-filename>…</OCR-filename>) ----------
// These wrap an OCR'd caption + Markdown table inside source .md files so the
// user can spot what came from OCR and re-OCR easily. They're parser-invisible
// in normal builds (just stripped) and act as re-OCR sentinels under --ocr.
function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function preprocessOcrMarkers(text, ocrEnabled) {
  if (ocrEnabled) {
    // Convert existing OCR blocks back to a bare fence so the OCR pipeline
    // refreshes them on this run.
    return text.replace(
      /<OCR-([^>\n]+)>[\s\S]*?<\/OCR-\1>/g,
      (_, filename) => "```" + filename + "```"
    );
  }
  // OCR disabled: drop marker lines, keep the cached caption + table visible.
  return text.replace(/^[ \t]*<\/?OCR-[^>\n]+>[ \t]*\r?\n?/gm, "");
}

function rowsToMarkdown(rows) {
  if (!rows || !rows.length) return "";
  const escape = (c) => String(c ?? "").replace(/\|/g, "\\|").replace(/\r?\n/g, " ");
  const colCount = Math.max(...rows.map((r) => r.length));
  const pad = (r) => Array.from({ length: colCount }, (_, i) => escape(r[i] ?? ""));
  const header = `| ${pad(rows[0]).join(" | ")} |`;
  const sep = `| ${Array.from({ length: colCount }, () => "---").join(" | ")} |`;
  const body = rows.slice(1).map((r) => `| ${pad(r).join(" | ")} |`).join("\n");
  return [header, sep, body].filter(Boolean).join("\n");
}

// Replace either an existing <OCR-filename>…</OCR-filename> block or a bare
// ```filename``` fence with `replacement`. Used to write OCR results back into
// the source .md so the table becomes editable and re-OCR'able.
function replaceOcrSourceInMd(text, filename, replacement) {
  const blockRe = new RegExp(
    `<OCR-${escapeRegex(filename)}>[\\s\\S]*?</OCR-${escapeRegex(filename)}>`
  );
  if (blockRe.test(text)) {
    return { text: text.replace(blockRe, replacement), replaced: true };
  }
  const eol = text.includes("\r\n") ? "\r\n" : "\n";
  const lines = text.split(eol);
  const out = [];
  let i = 0;
  let replaced = false;
  while (i < lines.length) {
    const line = lines[i];
    if (!replaced) {
      const inline = line.match(/^\s*```(.*?)```\s*$/);
      if (inline && inline[1].trim() === filename) {
        out.push(replacement);
        replaced = true;
        i++;
        continue;
      }
      if (/^\s*```/.test(line) && !inline) {
        const startIdx = i;
        const firstLineContent = line.replace(/^\s*```\S*\s*/, "");
        const buf = [];
        if (firstLineContent) buf.push(firstLineContent);
        i++;
        while (i < lines.length && !/^\s*```/.test(lines[i])) {
          buf.push(lines[i]);
          i++;
        }
        const endIdx = i;
        i++; // consume closing fence
        const content = buf.join("\n").trim();
        if (content === filename) {
          out.push(replacement);
          replaced = true;
          continue;
        }
        for (let j = startIdx; j <= endIdx; j++) out.push(lines[j]);
        continue;
      }
    }
    out.push(line);
    i++;
  }
  return { text: out.join(eol), replaced };
}

function parseTableFilename(filename) {
  // "1-1-1売上表.png" -> {number:"1.1.1", name:"売上表"}
  const m = filename.match(/^(\d+(?:-\d+)*)(.+?)\.(jpg|jpeg|png|gif|bmp|webp)$/i);
  if (!m) return null;
  return { number: m[1].replace(/-/g, "."), name: m[2].trim() };
}

// ---------- figure (image) handling ----------
function parseFigureFilename(filename) {
  // "1-1-1こんにちは.jpg" -> {number:"1.1.1", name:"こんにちは"}
  const m = filename.match(/^(\d+(?:-\d+)*)(.+?)\.(jpg|jpeg|png|gif|bmp)$/i);
  if (!m) return null;
  return {
    number: m[1].replace(/-/g, "."),
    name: m[2].trim(),
    ext: m[3].toLowerCase(),
  };
}

function figureXml(rId, cx, cy, captionText) {
  // cx/cy in EMU. drawing inline.
  // keepNext keeps the figure glued to its caption; keepLines stops the figure
  // from straddling a page boundary, so an oversized figure starts on a fresh page.
  const drawing =
    `<w:p><w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${rId}" name="Picture ${rId}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${rId}" name="Picture ${rId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`;
  const caption = pCenter(captionText, { keepLines: true });
  return drawing + caption;
}

// ---------- setup (GUI input for Gemini API key) ----------
// Pops a native OS dialog so non-technical users can paste the key once.
// The key is written to ~/.mkreport/config.env and reused by every later run.
function promptApiKeyGUI() {
  const title = "mkreport setup";
  const prompt = "Gemini API キーを入力してください (取得: https://aistudio.google.com/apikey)";

  if (process.platform === "darwin") {
    const scriptLines = [
      "try",
      `set d to display dialog "${prompt}" default answer "" with title "${title}" with hidden answer`,
      "return text returned of d",
      "on error",
      'return ""',
      "end try",
    ];
    const args = scriptLines.flatMap((l) => ["-e", l]);
    const r = spawnSync("osascript", args, { encoding: "utf8" });
    if (r.error || r.status !== 0) return { error: r.error?.message || `osascript exit ${r.status}` };
    return { key: r.stdout.trim() };
  }

  if (process.platform === "win32") {
    const ps = [
      "Add-Type -AssemblyName System.Windows.Forms,System.Drawing",
      "$f = New-Object System.Windows.Forms.Form",
      `$f.Text = '${title}'`,
      "$f.Size = New-Object System.Drawing.Size(460,200)",
      "$f.StartPosition = 'CenterScreen'",
      "$f.Topmost = $true",
      "$l = New-Object System.Windows.Forms.Label",
      "$l.Location = New-Object System.Drawing.Point(10,15)",
      "$l.Size = New-Object System.Drawing.Size(430,40)",
      `$l.Text = '${prompt}'`,
      "$f.Controls.Add($l)",
      "$t = New-Object System.Windows.Forms.TextBox",
      "$t.Location = New-Object System.Drawing.Point(10,60)",
      "$t.Size = New-Object System.Drawing.Size(420,25)",
      "$t.UseSystemPasswordChar = $true",
      "$f.Controls.Add($t)",
      "$ok = New-Object System.Windows.Forms.Button",
      "$ok.Location = New-Object System.Drawing.Point(350,110)",
      "$ok.Size = New-Object System.Drawing.Size(80,28)",
      "$ok.Text = 'OK'",
      "$ok.DialogResult = [System.Windows.Forms.DialogResult]::OK",
      "$f.AcceptButton = $ok",
      "$f.Controls.Add($ok)",
      "$c = New-Object System.Windows.Forms.Button",
      "$c.Location = New-Object System.Drawing.Point(260,110)",
      "$c.Size = New-Object System.Drawing.Size(80,28)",
      "$c.Text = 'Cancel'",
      "$c.DialogResult = [System.Windows.Forms.DialogResult]::Cancel",
      "$f.CancelButton = $c",
      "$f.Controls.Add($c)",
      "$r = $f.ShowDialog()",
      "if ($r -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $t.Text }",
    ].join("\n");
    // UTF-16LE base64 avoids all quoting headaches on the PowerShell command line.
    const encoded = Buffer.from(ps, "utf16le").toString("base64");
    const r = spawnSync("powershell", ["-NoProfile", "-EncodedCommand", encoded], { encoding: "utf8" });
    if (r.error || r.status !== 0) return { error: r.error?.message || `powershell exit ${r.status}` };
    return { key: r.stdout.trim() };
  }

  // Linux: try zenity then kdialog.
  let r = spawnSync("zenity", ["--entry", "--hide-text", `--title=${title}`, `--text=${prompt}`], { encoding: "utf8" });
  if (!r.error && r.status === 0) return { key: r.stdout.trim() };
  r = spawnSync("kdialog", ["--password", prompt, "--title", title], { encoding: "utf8" });
  if (!r.error && r.status === 0) return { key: r.stdout.trim() };
  return { error: "GUI ダイアログを起動できませんでした (zenity または kdialog をインストールしてください)" };
}

function runSetup() {
  console.log("Gemini API キーを設定します。ダイアログに貼り付けてください。");
  const { key, error } = promptApiKeyGUI();
  if (error) {
    console.error(`error: ${error}`);
    process.exit(1);
  }
  if (!key) {
    console.error("キャンセルされました (キーが空でした)。");
    process.exit(1);
  }
  fs.mkdirSync(USER_CONFIG_DIR, { recursive: true });
  fs.writeFileSync(USER_CONFIG_FILE, `GEMINI_API_KEY=${key}\n`, { mode: 0o600 });
  console.log(`保存しました: ${USER_CONFIG_FILE}`);
  console.log("以降、このコンピュータ上のすべての mkreport 実行で同じキーが使用されます。");
}

// ---------- scaffold ----------
function scaffold(name) {
  if (!name) {
    console.error("error: project name is required\n  usage: mkreport start <project-name>");
    process.exit(1);
  }
  const target = path.resolve(CWD, name);
  if (fs.existsSync(target)) {
    console.error(`error: ${target} already exists`);
    process.exit(1);
  }
  if (!fs.existsSync(TEMPLATES_DIR)) {
    console.error(`error: templates directory not found at ${TEMPLATES_DIR}`);
    process.exit(1);
  }
  fs.cpSync(TEMPLATES_DIR, target, { recursive: true });
  console.log(`created ${target}`);
  console.log("");
  console.log(`Next steps:`);
  console.log(`  1. Edit files inside ${name}/ (see ${name}/README.md)`);
  console.log(`  2. Run: mkreport ${name}`);
}

// ---------- build ----------
async function buildReport(projectArg, outputArg) {
  const projectRoot = path.resolve(CWD, projectArg);
  if (!fs.existsSync(projectRoot) || !fs.statSync(projectRoot).isDirectory()) {
    console.error(`error: ${projectRoot} is not a directory`);
    process.exit(1);
  }
  PROJECT_ROOT = projectRoot;
  const projectName = path.basename(projectRoot);
  const OUTPUT = outputArg
    ? path.resolve(CWD, outputArg)
    : path.join(projectRoot, `${projectName}.docx`);

  const seedPath = findInputFile("seed.docx");
  if (!seedPath) {
    console.error(`error: seed.docx not found in ${projectRoot} or bundled templates/`);
    process.exit(1);
  }

  // Load placeholders
  const titlePath = findInputFile("title.env");
  const titleMap = parseTitleEnv(readTextIfExists(titlePath));

  // Load section md files
  const sections = [
    { num: 1, headingRegexName: "目的", file: "1mokuteki.md" },
    { num: 2, headingRegexName: "実験装置", file: "2souchi.md" },
    { num: 3, headingRegexName: "実験方法", file: "3houhou.md" },
    { num: 4, headingRegexName: "実験結果", file: "4kekka.md" },
    { num: 5, headingRegexName: "考察", file: "5kousatsu.md" },
  ];

  // Citation
  const inyouPath = findInputFile("6inyou.env");
  const inyouText = readTextIfExists(inyouPath);

  // Load and unzip the seed docx
  const seedBuf = fs.readFileSync(seedPath);
  const zip = await JSZip.loadAsync(seedBuf);

  let docXml = await zip.file("word/document.xml").async("string");

  // Body text width in EMU = (page width - left - right margin) twips * 635.
  // Used by the `<->` figure marker to stretch an image to the full text column.
  let bodyWidthEmu = 5400040; // fallback: A4 with 30mm margins
  {
    const szM = docXml.match(/<w:pgSz\b[^>]*w:w="(\d+)"/);
    const marM = docXml.match(/<w:pgMar\b[^>]*?w:right="(\d+)"[^>]*?w:left="(\d+)"/) ||
                 docXml.match(/<w:pgMar\b[^>]*?w:left="(\d+)"[^>]*?w:right="(\d+)"/);
    if (szM && marM) {
      const pageW = Number(szM[1]);
      const a = Number(marM[1]), b = Number(marM[2]);
      const textW = pageW - a - b;
      if (textW > 0) bodyWidthEmu = textW * 635;
    }
  }

  // 1) substitute title placeholders
  docXml = substitutePlaceholders(docXml, titleMap);

  // 2) build body content per section, including image relationships
  // Track equation counter and rId/image registry
  let eqnCounter = 0;
  const imageEntries = []; // {rId, target, dataBuf, contentType}
  let nextRId = 100; // start far above any rIds already in the seed

  // Read existing relationships to avoid rId collisions
  const relsPath = "word/_rels/document.xml.rels";
  let relsXml = await zip.file(relsPath).async("string");
  const existingIds = new Set();
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) existingIds.add(Number(m[1]));
  while (existingIds.has(nextRId)) nextRId++;

  function allocRId() {
    while (existingIds.has(nextRId)) nextRId++;
    const id = nextRId++;
    existingIds.add(id);
    return id;
  }

  const mdRewrites = []; // {fileName, fenceContent, replacement}
  const tableSheets = []; // {caption, rows} — one Excel sheet per table

  async function renderBlocks(blocks, fileName) {
    const out = [];
    for (let idx = 0; idx < blocks.length; idx++) {
      const b = blocks[idx];
      const next = blocks[idx + 1];
      if (b.type === "p") {
        // A paragraph immediately before a table is treated as that table's
        // title: keepNext glues it to the table so the caption never sits at the
        // bottom of one page with the table pushed onto the next. (The --ocr path
        // sets this directly; this covers cached <OCR-…> blocks built without --ocr
        // and any plain "title + table" written by hand.)
        const keepNext = next && next.type === "table";
        out.push(pPlain(b.text, b.align || "left", { keepNext }));
      } else if (b.type === "table") {
        out.push(tableXml(b.rows));
        // A paragraph immediately above the table is its caption (used as the
        // Excel sheet name); fall back to a running 表N number otherwise.
        const prev = blocks[idx - 1];
        const caption = prev && prev.type === "p" ? prev.text.split(/\n/)[0].trim() : "";
        tableSheets.push({ caption, rows: b.rows });
      } else if (b.type === "fence") {
        // `<->` anywhere in the fence stretches a figure to the full text width.
        const fullWidth = b.content.includes("<->");
        const content = b.content.replace(/<->/g, "").trim();

        // 1) OCR table: file in ocr_hyo/
        const ocrPath = findAssetFile(content, "ocr_hyo");
        if (ocrPath && /\.(jpg|jpeg|png|gif|bmp|webp)$/i.test(content)) {
          const info = parseTableFilename(content);
          const captionText = info ? `表${info.number} ${info.name}` : content;
          if (!FLAGS.ocr) {
            // Without --ocr we don't call the API. Drop a centered placeholder so
            // the doc still shows where the table belongs.
            out.push(pCenter(`${captionText} (OCR未実行 — --ocr を付けて再実行してください)`));
            continue;
          }
          out.push(pCenter(captionText, { keepNext: true }));
          const result = await geminiOcrTable(ocrPath);
          if (result.rows && result.rows.length > 0) {
            out.push(tableXml(result.rows));
            tableSheets.push({ caption: captionText, rows: result.rows });
            // Queue source .md rewrite: replace fence (or stale OCR block) with
            // a fresh <OCR-…>…</OCR-…> block so the user can edit / re-OCR later.
            const replacement =
              `<OCR-${content}>\n/c ${captionText} c/\n\n${rowsToMarkdown(result.rows)}\n</OCR-${content}>`;
            mdRewrites.push({ fileName, fenceContent: content, replacement });
          } else {
            const reason = result.error || "不明なエラー";
            console.warn(`[OCR 失敗] ${content}: ${reason}`);
            out.push(pCenter(`失敗: ${reason}`));
          }
          continue;
        }

        // 2) Figure: file in figures/
        const figPath = findAssetFile(content, "figures");
        if (figPath && /\.(jpg|jpeg|png|gif|bmp)$/i.test(content)) {
          const info = parseFigureFilename(content);
          const buf = fs.readFileSync(figPath);
          let cx = 4572000, cy = 3429000; // default 5in x 3.75in
          try {
            const dim = imageSize(buf);
            if (dim?.width && dim?.height) {
              const pxToEmu = 9525;
              let w = dim.width * pxToEmu;
              let h = dim.height * pxToEmu;
              if (fullWidth) {
                // Scale to exactly the text-column width, keeping aspect ratio.
                h = h * (bodyWidthEmu / w); w = bodyWidthEmu;
              } else {
                const maxCx = 5029200;
                if (w > maxCx) { h = h * (maxCx / w); w = maxCx; }
              }
              cx = Math.round(w); cy = Math.round(h);
            }
          } catch {}
          const rId = allocRId();
          const ext = (info?.ext || path.extname(content).slice(1)).toLowerCase();
          const target = `media/image_${rId}.${ext === "jpg" ? "jpeg" : ext}`;
          imageEntries.push({
            rId,
            target,
            dataBuf: buf,
            contentType: ext === "png" ? "image/png" :
                         ext === "gif" ? "image/gif" :
                         ext === "bmp" ? "image/bmp" : "image/jpeg",
          });
          const caption = info ? `図${info.number} ${info.name}` : content;
          out.push(figureXml(rId, cx, cy, caption));
          continue;
        }

        // 3) Equation
        eqnCounter++;
        out.push(equationXml(content, eqnCounter));
      }
    }
    return out.join("");
  }

  // Build per-section XML
  const sectionXmls = {};
  for (const s of sections) {
    const rawText = readTextIfExists(findInputFile(s.file));
    const mdText = preprocessOcrMarkers(rawText, FLAGS.ocr);
    const chunks = splitByAlignment(mdText);
    const blocks = [];
    for (const c of chunks) {
      const bs = parseMarkdown(c.text);
      for (const b of bs) {
        if (b.type === "p") b.align = c.align;
        blocks.push(b);
      }
    }
    sectionXmls[s.num] = await renderBlocks(blocks, s.file);
  }

  // Apply queued OCR rewrites back to the source .md files. Always writes to
  // PROJECT_ROOT so we never modify bundled templates in place.
  if (mdRewrites.length) {
    const byFile = {};
    for (const r of mdRewrites) {
      (byFile[r.fileName] = byFile[r.fileName] || []).push(r);
    }
    for (const [fileName, rewrites] of Object.entries(byFile)) {
      const dst = path.join(PROJECT_ROOT, fileName);
      const srcPath = fs.existsSync(dst) ? dst : findInputFile(fileName);
      if (!srcPath) continue;
      let text = fs.readFileSync(srcPath, "utf8").replace(/^﻿/, "");
      let any = false;
      for (const r of rewrites) {
        const res = replaceOcrSourceInMd(text, r.fenceContent, r.replacement);
        if (res.replaced) { text = res.text; any = true; }
      }
      if (any) {
        fs.writeFileSync(dst, text);
        console.log(`updated ${dst}: OCR結果を <OCR-…> ブロックに反映しました`);
      }
    }
  }

  // Build citation XML (6, 引用)
  const citationLines = inyouText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  const citationXml = citationLines
    .map((line, idx) => {
      // CSV: year, author, title -> [N] author, title, year.
      const parts = line.split(",").map((p) => p.trim());
      const [year, author, title] = parts;
      return pPlain(`[${idx + 1}] ${author ?? ""}, ${title ?? ""}, ${year ?? ""}.`);
    })
    .join("");
  sectionXmls[6] = citationXml;

  // 3) Inject content right after each section heading paragraph
  // The heading paragraphs look like: <w:p ...>...N, ...HEADINGNAME...</w:p>
  const headingMap = [
    { num: 1, name: "目的" },
    { num: 2, name: "実験装置" },
    { num: 3, name: "実験方法" },
    { num: 4, name: "実験結果" },
    { num: 5, name: "考察" },
    { num: 6, name: "引用" },
  ];

  for (const h of headingMap) {
    // Regex matches the <w:p> that contains both "N, " and the heading name (e.g., "目的")
    // We use a tempered group: (?:(?!</w:p>).)* to avoid matching across paragraphs.
    const re = new RegExp(
      `(<w:p\\b(?:(?!</w:p>).)*?${h.num},\\s(?:(?!</w:p>).)*?${h.name}(?:(?!</w:p>).)*?</w:p>)`,
      "s"
    );
    const replacement = `$1${sectionXmls[h.num] || ""}`;
    if (re.test(docXml)) {
      docXml = docXml.replace(re, replacement);
    } else {
      console.warn(`warning: heading "${h.num}, ${h.name}" not found in seed.docx`);
    }
  }

  // 4) Update relationships and [Content_Types].xml for new images
  if (imageEntries.length) {
    // Add relationships
    const newRels = imageEntries
      .map(
        (e) =>
          `<Relationship Id="rId${e.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${e.target}"/>`
      )
      .join("");
    relsXml = relsXml.replace(/<\/Relationships>\s*$/, `${newRels}</Relationships>`);
    zip.file(relsPath, relsXml);

    // Add image binaries
    for (const e of imageEntries) {
      zip.file(`word/${e.target}`, e.dataBuf);
    }

    // Update [Content_Types].xml to declare image extensions
    const ctPath = "[Content_Types].xml";
    let ctXml = await zip.file(ctPath).async("string");
    const neededExts = new Set();
    for (const e of imageEntries) {
      const ext = e.target.split(".").pop().toLowerCase();
      neededExts.add(ext);
    }
    for (const ext of neededExts) {
      const ctype =
        ext === "png" ? "image/png" :
        ext === "gif" ? "image/gif" :
        ext === "bmp" ? "image/bmp" : "image/jpeg";
      const probe = new RegExp(`Extension="${ext}"`, "i");
      if (!probe.test(ctXml)) {
        ctXml = ctXml.replace(
          /<\/Types>\s*$/,
          `<Default Extension="${ext}" ContentType="${ctype}"/></Types>`
        );
      }
    }
    zip.file(ctPath, ctXml);
  }

  // Save the modified document.xml back
  zip.file("word/document.xml", docXml);

  // Write output
  const outBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(OUTPUT, outBuf);
  console.log(`wrote ${OUTPUT}`);

  // Companion .xlsx: one sheet per table, same name as the Word file.
  if (tableSheets.length) {
    const usedNames = new Set();
    const sheets = tableSheets.map((t, i) => ({
      name: sanitizeSheetName(t.caption, `表${i + 1}`, usedNames),
      rows: t.rows,
    }));
    const xlsxOut = OUTPUT.replace(/\.docx$/i, "") + ".xlsx";
    const xlsxBuf = await buildXlsx(sheets);
    fs.writeFileSync(xlsxOut, xlsxBuf);
    console.log(`wrote ${xlsxOut} (${sheets.length} 表)`);
  }
}

// ---------- entry point ----------
async function main() {
  if (args[0] === "setup") {
    runSetup();
    return;
  }
  if (args[0] === "start") {
    scaffold(args[1]);
    return;
  }
  if (args[0] === "--help" || args[0] === "-h" || args[0] === "help") {
    printUsageAndExit(0);
  }
  await buildReport(args[0], args[1]);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
