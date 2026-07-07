import { getClient, evaluate } from '../connection.js';
import { writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENSHOT_DIR = join(dirname(dirname(__dirname)), 'screenshots');

export async function captureScreenshot({ region, filename, scale } = {}) {
  mkdirSync(SCREENSHOT_DIR, { recursive: true });

  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const fname = (filename || `tv_${region}_${ts}`).replace(/[\/\\]/g, '_');
  const filePath = join(SCREENSHOT_DIR, `${fname}.png`);

  const client = await getClient();
  const params = { format: 'png' };

  // Rendering at a higher device scale factor (without changing CSS layout) sharpens
  // axis labels/text in dense multi-pane grids — plain 1x captures compress each pane's
  // price axis to the point individual digits become unreliable to read back.
  const useScale = scale && scale > 1;
  if (useScale) {
    const dims = await evaluate(`({w: window.innerWidth, h: window.innerHeight})`);
    await client.Emulation.setDeviceMetricsOverride({ width: dims.w, height: dims.h, deviceScaleFactor: scale, mobile: false });
  }

  let data;
  try {
    ({ data } = await client.Page.captureScreenshot(params));
  } finally {
    if (useScale) await client.Emulation.clearDeviceMetricsOverride();
  }

  writeFileSync(filePath, Buffer.from(data, 'base64'));

  return {
    success: true, method: 'cdp', file_path: filePath, region,
    size_bytes: Buffer.from(data, 'base64').length,
  };
}
