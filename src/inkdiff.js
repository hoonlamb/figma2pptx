'use strict';
// Element-level verification.
//
// A page-level "6.32% of pixels differ" cannot tell a shifted table from text
// that renders slightly narrow — the two look identical as a number, and the
// only way to find out which one it was has been to write a throwaway script
// and measure by hand. So measure every text box instead: crop the same
// rectangle out of both renders, find where the ink actually starts and ends,
// and report the difference in the units the design was drawn in.

const BG_TOLERANCE = 40;    // per-pixel distance from the local background to count as ink
const MIN_INK = 6;          // fewer marked pixels than this is noise, not a glyph

/** Most common quantised colour in the region — the local background. */
function backgroundOf(img, x0, y0, x1, y1) {
  const seen = new Map();
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const i = (y * img.width + x) << 2;
      // 5-bit buckets: robust to antialiasing without merging distinct fills
      const k = ((img.data[i] >> 3) << 10) | ((img.data[i + 1] >> 3) << 5) | (img.data[i + 2] >> 3);
      seen.set(k, (seen.get(k) || 0) + 1);
    }
  }
  let best = 0, bestK = 0;
  for (const [k, n] of seen) if (n > best) { best = n; bestK = k; }
  return [((bestK >> 10) & 31) << 3, ((bestK >> 5) & 31) << 3, (bestK & 31) << 3];
}

/** Bounding box of everything that is not the local background.
 *  `blocked` lists rectangles to ignore — a rule or border crossing the text's
 *  box is ink too, and if it lands inside the box in one render and a pixel
 *  outside it in the other, the measurement blames the text for it. */
function inkExtent(img, x0, y0, x1, y1, blocked = []) {
  x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
  x1 = Math.min(img.width, Math.ceil(x1)); y1 = Math.min(img.height, Math.ceil(y1));
  if (x1 - x0 < 2 || y1 - y0 < 2) return null;

  const bg = backgroundOf(img, x0, y0, x1, y1);
  const masked = (x, y) => blocked.some((m) => x >= m[0] && x <= m[2] && y >= m[1] && y <= m[3]);

  let l = Infinity, r = -Infinity, t = Infinity, b = -Infinity, n = 0;
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const i = (y * img.width + x) << 2;
      const d = Math.abs(img.data[i] - bg[0]) + Math.abs(img.data[i + 1] - bg[1]) + Math.abs(img.data[i + 2] - bg[2]);
      if (d <= BG_TOLERANCE) continue;
      if (masked(x, y)) continue;
      n++;
      if (x < l) l = x; if (x > r) r = x;
      if (y < t) t = y; if (y > b) b = y;
    }
  }
  return n < MIN_INK ? null : { l, r, t, b, n };
}

/**
 * Compare each text box between the two aligned renders.
 * @param A  Figma render, @param B  PowerPoint render — both already the same size
 * @param slide  normalised slide (ops carry frame-pixel geometry)
 * @param scale  aligned pixels per frame pixel
 * @returns findings in FRAME pixels, worst first
 */
function compareText(A, B, slide, scale, cfg = {}) {
  // Never claim precision the raster cannot support: one render pixel is
  // 1/scale design pixels, so on a 1920px frame rendered at 1600 every
  // measurement is quantised to 1.2px. Reporting those as findings buries the
  // real ones in noise.
  const floor = 1.5 / scale;
  const tolPx = Math.max(cfg.elementTolerancePx === undefined ? 2 : cfg.elementTolerancePx, floor);
  const tolRatio = cfg.elementToleranceRatio === undefined ? 0.02 : cfg.elementToleranceRatio;
  const out = [];

  // Rules and borders are the neighbours that bleed into a text box; a filled
  // chip behind the text does not, because it becomes the local background.
  const rules = slide.ops
    .filter((o) => o.kind === 'line' || o.kind === 'outline')
    .map((o) => [(o.x - 1) * scale, (o.y - 1) * scale, (o.x + o.w + 1) * scale, (o.y + o.h + 1) * scale]);

  for (const op of slide.ops) {
    if (op.kind !== 'text') continue;
    // a little slack so a glyph that overhangs its box is not clipped off
    const pad = 2 / scale;
    const box = [(op.x - pad) * scale, (op.y - pad) * scale, (op.x + op.w + pad) * scale, (op.y + op.h + pad) * scale];
    const near = rules.filter((m) => m[0] <= box[2] && m[2] >= box[0] && m[1] <= box[3] && m[3] >= box[1]);
    const a = inkExtent(A, ...box, near);
    const b = inkExtent(B, ...box, near);

    if (!a && !b) continue;
    const label = (op.runs || []).map((r) => r.text).join('').trim().slice(0, 22);
    if (!a || !b) {
      out.push({ label, kind: a ? 'missing' : 'extra', dx: 0, dw: 0, dy: 0, aw: 0, ratio: 0, severity: 1e9 });
      continue;
    }

    const aw = (a.r - a.l) / scale, bw = (b.r - b.l) / scale;
    const dx = (b.l - a.l) / scale;
    const dy = (b.t - a.t) / scale;
    const dw = bw - aw;

    // Narrower text drags its own left edge inward when the box is centred or
    // right-aligned. That displacement is the width difference showing up again,
    // not an independent move — subtract it, or every centred label in the deck
    // gets reported twice under two different faults.
    const alignH = (op.style || {}).textAlignHorizontal;
    const expectedDx = alignH === 'CENTER' ? -dw / 2 : alignH === 'RIGHT' ? -dw : 0;
    const moveX = dx - expectedDx;

    out.push({
      label, kind: 'measured',
      dx, dy, dw, aw, moveX,
      ratio: aw > 1 ? dw / aw : 0,          // signed: negative means PPT is narrower
      tolPx, tolRatio,
      severity: Math.abs(moveX) + Math.abs(dy),
    });
  }
  return out;
}

module.exports = { compareText, inkExtent };
