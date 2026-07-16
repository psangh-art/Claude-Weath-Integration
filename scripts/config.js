// Single source of truth for machine-specific paths/IDs/ports — the Node-side
// twin of config.py. Both read scripts/config.json; moving machines means
// editing that ONE file. Added 2026-07-12 (external-review follow-up approved
// by the user).
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf-8'));

export function downloadsDir() {
  return CFG.downloadsDir.startsWith('~')
    ? path.join(os.homedir(), CFG.downloadsDir.slice(1))
    : CFG.downloadsDir;
}

export function downloadsFile(key) {
  return path.join(downloadsDir(), CFG[key]);
}

// First configured Python interpreter that actually exists on this machine.
export function pythonExe() {
  for (const p of CFG.pythonCandidates) {
    if (p === 'python') return p;
    try { if (fs.existsSync(p)) return p; } catch { /* keep looking */ }
  }
  return 'python';
}

export function financeSheetUrl() {
  return `https://docs.google.com/spreadsheets/d/${CFG.financeSheetId}/edit`;
}

// Folder the pipeline app mirrors built products into so OneDrive syncs them —
// Office web apps (PowerPoint/Excel Online) can only open files that live
// inside an actual OneDrive-synced folder, not ~/Downloads.
export function onedriveProductsDir() {
  const raw = CFG.onedriveProductsDir || '';
  return raw.startsWith('~') ? path.join(os.homedir(), raw.slice(1)) : raw;
}

// One-time-pasted OneDrive share/web URL for a product, keyed by
// config.json's productWebLinks (reviewDeck/architecturePptx/spendingSummary).
// Empty string (the default) means "not pasted yet".
export function productWebLink(key) {
  const v = CFG.productWebLinks && CFG.productWebLinks[key];
  return v ? v : null;
}
