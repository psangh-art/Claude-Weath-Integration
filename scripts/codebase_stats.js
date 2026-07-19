// Size + freshness of this application's own source, for the Overview banner
// (user request 2026-07-19: "confirm the number of lines of code the application
// uses and a last updated date").
//
// Counts only HAND-WRITTEN source. Three categories are deliberately excluded, or
// the number would measure the wrong thing:
//   * dependencies (node_modules) — not this application's code;
//   * GENERATED artefacts that happen to live in the tree — architecture.html /
//     alert_rules.html are rendered from the .pptx decks (1.9 MB of machine output
//     that would swamp everything), review_deck.html likewise, and package-lock.json;
//   * run scratch — *_tmp.json manifests, logs/, screenshots/, data/, __pycache__.
// Docs (.md) are excluded too: CLAUDE.md alone is thousands of lines of prose.
//
// "Last updated" is the last COMMIT date, not a file mtime — an mtime changes when
// a generated file is rewritten or OneDrive touches something, which would make the
// banner claim an update that never happened. Falls back to the newest counted
// source file's mtime outside a git checkout.
import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.join(__dirname, '..');

const COUNTED_EXT = new Set(['.js', '.py', '.html', '.css', '.ps1', '.bat', '.mjs']);
const SKIP_DIRS = new Set([
  'node_modules', '.git', '__pycache__', '.claude', '.venv', 'venv',
  'logs', 'screenshots', 'data', '_annotated_charts', 'dist', 'build',
]);
const SKIP_FILES = new Set([
  'package-lock.json', 'architecture.html', 'alert_rules.html', 'review_deck.html',
]);
const SKIP_RE = /_tmp\.json$|\.min\.(js|css)$/i;

function walk(dir, out) {
  let entries;
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return out; }
  for (const e of entries) {
    if (e.name.startsWith('.') && e.name !== '.claude') { if (e.isDirectory()) continue; }
    const abs = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (!SKIP_DIRS.has(e.name)) walk(abs, out);
      continue;
    }
    if (!e.isFile()) continue;
    if (SKIP_FILES.has(e.name) || SKIP_RE.test(e.name)) continue;
    if (!COUNTED_EXT.has(path.extname(e.name).toLowerCase())) continue;
    out.push(abs);
  }
  return out;
}

function countLines(file) {
  try {
    const buf = fs.readFileSync(file);
    if (!buf.length) return 0;
    let n = 0;
    for (let i = 0; i < buf.length; i++) if (buf[i] === 0x0a) n++;
    return buf[buf.length - 1] === 0x0a ? n : n + 1;   // count a final unterminated line
  } catch { return 0; }
}

function lastCommitDate() {
  const r = spawnSync('git', ['-C', REPO_ROOT, 'log', '-1', '--format=%cI'], { encoding: 'utf-8' });
  const out = (r.stdout || '').trim();
  return r.status === 0 && out ? out : null;
}

let cache = null;
const CACHE_MS = 60_000;

/** {lines, files, by_language:{ext:{files,lines}}, last_updated, source} */
export function codebaseStats() {
  if (cache && Date.now() - cache.at < CACHE_MS) return cache.value;

  const files = walk(REPO_ROOT, []);
  const byLang = {};
  let lines = 0, newestMtime = 0;
  for (const f of files) {
    const n = countLines(f);
    const ext = path.extname(f).toLowerCase().slice(1);
    const bucket = byLang[ext] || (byLang[ext] = { files: 0, lines: 0 });
    bucket.files++; bucket.lines += n;
    lines += n;
    try { newestMtime = Math.max(newestMtime, fs.statSync(f).mtimeMs); } catch { /* raced */ }
  }

  const commit = lastCommitDate();
  const value = {
    lines,
    files: files.length,
    by_language: byLang,
    last_updated: commit || (newestMtime ? new Date(newestMtime).toISOString() : null),
    source: commit ? 'last commit' : 'newest source file',
  };
  cache = { at: Date.now(), value };
  return value;
}
