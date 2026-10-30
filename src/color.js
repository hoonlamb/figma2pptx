'use strict';
// PowerPoint has no per-shape alpha that survives every renderer reliably, so any
// translucent paint is pre-blended against the background it actually sits on.
// The normalizer tracks that background while walking the tree.

const clamp255 = (v) => Math.max(0, Math.min(255, Math.round(v)));

function toRgb(figmaColor, paintOpacity = 1, nodeOpacity = 1) {
  if (!figmaColor) return null;
  const a = (figmaColor.a === undefined ? 1 : figmaColor.a) * paintOpacity * nodeOpacity;
  return { r: clamp255(figmaColor.r * 255), g: clamp255(figmaColor.g * 255), b: clamp255(figmaColor.b * 255), a };
}

function blend(fg, bg) {
  if (!fg) return null;
  if (fg.a >= 0.999) return { r: fg.r, g: fg.g, b: fg.b, a: 1 };
  const b = bg || { r: 255, g: 255, b: 255 };
  return {
    r: clamp255(fg.r * fg.a + b.r * (1 - fg.a)),
    g: clamp255(fg.g * fg.a + b.g * (1 - fg.a)),
    b: clamp255(fg.b * fg.a + b.b * (1 - fg.a)),
    a: 1,
  };
}

const hex = (c) =>
  c ? [c.r, c.g, c.b].map((v) => v.toString(16).padStart(2, '0').toUpperCase()).join('') : '000000';

/** First visible SOLID paint of a paint array. */
function solidPaint(paints, nodeOpacity = 1) {
  if (!Array.isArray(paints)) return null;
  for (const p of paints) {
    if (p.visible === false) continue;
    if (p.type === 'SOLID') return toRgb(p.color, p.opacity === undefined ? 1 : p.opacity, nodeOpacity);
    return null; // gradient/image → caller rasterizes instead
  }
  return null;
}

function hasNonSolidPaint(paints) {
  return Array.isArray(paints) && paints.some((p) => p.visible !== false && p.type !== 'SOLID');
}

module.exports = { toRgb, blend, hex, solidPaint, hasNonSolidPaint };
