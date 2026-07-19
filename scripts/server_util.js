// Static-file serving shared by the two local servers (added 2026-07-19).
//
// pipeline_app_server.js (Production Centre, 4590) and dashboard_server.js
// (Investment Dashboard, 4600) each carried their own serveFile, their own
// content-type map and a near-identical isAllowedAsset — the asset-proxy
// whitelist in particular was duplicated with a comment in one file saying "same
// rule as the other", which is exactly the shape of duplication that lets a
// security rule get tightened in one place and not the other.
import fs from 'fs';
import path from 'path';

// Union of what both servers serve. Anything unlisted falls back to a download.
const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

export function contentTypeFor(filePath) {
  return CONTENT_TYPES[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
}

/** Stream `filePath` to `res`, 404 if it isn't there. `download` forces a save-as. */
export function serveFile(res, filePath, { download } = {}) {
  if (!fs.existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
    return;
  }
  const headers = { 'Content-Type': contentTypeFor(filePath) };
  if (download) headers['Content-Disposition'] = `attachment; filename="${path.basename(filePath)}"`;
  res.writeHead(200, headers);
  fs.createReadStream(filePath).pipe(res);
}

/**
 * Whitelist for the `/asset?p=` image proxy both servers expose: an image file
 * inside one of `roots` and nothing else — never an arbitrary path off the query
 * string. Kept deliberately strict (extension AND containment).
 */
export function isAllowedAsset(abs, roots) {
  if (!abs) return false;
  const resolved = path.resolve(abs);
  return roots.map(r => path.resolve(r))
    .some(root => resolved === root || resolved.startsWith(root + path.sep))
    && /\.(png|jpe?g)$/i.test(resolved);
}
