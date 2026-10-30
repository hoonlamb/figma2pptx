'use strict';
// Figma node tree -> flat, absolutely-positioned draw list for one slide.
//
// Why absolute coordinates: in Figma a GROUP does not create a new coordinate
// space for its children, while a FRAME does. Accumulating parent offsets by hand
// silently double-shifts everything inside a group. absoluteBoundingBox sidesteps
// the whole class of bug, so every op below is expressed in frame-local pixels
// derived from absolute boxes.

const { solidPaint, hasNonSolidPaint, blend } = require('./color');

const RASTER_TYPES = new Set(['VECTOR', 'BOOLEAN_OPERATION', 'STAR', 'POLYGON', 'REGULAR_POLYGON', 'SLICE']);
const CONTAINER_TYPES = new Set(['FRAME', 'GROUP', 'COMPONENT', 'COMPONENT_SET', 'INSTANCE', 'SECTION']);

const box = (n) => n.absoluteBoundingBox || null;

function intersect(a, b) {
  if (!a) return b;
  if (!b) return a;
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  const r = Math.min(a.x + a.w, b.x + b.w);
  const bt = Math.min(a.y + a.h, b.y + b.h);
  return { x, y, w: Math.max(0, r - x), h: Math.max(0, bt - y) };
}

// Rules and dividers are legitimately zero-thickness on one axis, so a plain
// "empty intersection means invisible" test would delete every hairline.
function clipRect(rect, clip) {
  if (!clip) return rect;
  const x1 = Math.max(rect.x, clip.x);
  const y1 = Math.max(rect.y, clip.y);
  const x2 = Math.min(rect.x + rect.w, clip.x + clip.w);
  const y2 = Math.min(rect.y + rect.h, clip.y + clip.h);
  const w = x2 - x1;
  const h = y2 - y1;
  if (w < 0 || h < 0) return null;

  const thinW = rect.w <= 0.5;
  const thinH = rect.h <= 0.5;
  if (!thinW && w <= 0) return null;
  if (!thinH && h <= 0) return null;
  return { x: x1, y: y1, w: thinW ? rect.w : w, h: thinH ? rect.h : h };
}

/** Average of gradient stops — used when a gradient container must stay editable. */
function gradientAverage(paints) {
  for (const p of paints || []) {
    if (p.visible === false) continue;
    if (p.type && p.type.startsWith('GRADIENT') && Array.isArray(p.gradientStops) && p.gradientStops.length) {
      const n = p.gradientStops.length;
      const acc = p.gradientStops.reduce(
        (a, s) => ({ r: a.r + s.color.r, g: a.g + s.color.g, b: a.b + s.color.b }),
        { r: 0, g: 0, b: 0 }
      );
      return { r: Math.round((acc.r / n) * 255), g: Math.round((acc.g / n) * 255), b: Math.round((acc.b / n) * 255), a: 1 };
    }
  }
  return null;
}

/** Split styled text into runs, breaking at newlines so PPT gets explicit line breaks. */
function buildRuns(node) {
  const chars = node.characters || '';
  const base = node.style || {};
  const overrides = node.characterStyleOverrides || [];
  const table = node.styleOverrideTable || {};

  const styleAt = (i) => {
    const key = overrides[i];
    const ov = key && table[key] ? table[key] : null;
    if (!ov) return base;
    const merged = Object.assign({}, base, ov);
    // Tracking is stored once in px at the base size. A smaller override without
    // its own value keeps the same percentage, not the same px.
    if (ov.fontSize && ov.letterSpacing === undefined && base.fontSize && base.letterSpacing) {
      merged.letterSpacing = base.letterSpacing * (ov.fontSize / base.fontSize);
    }
    return merged;
  };

  // group consecutive characters sharing a style key, splitting on line breaks
  const segs = [];
  let cur = null;
  // 10 = \n, 13 = \r, 11 = vertical tab, 8232 = U+2028 line separator.
  // Figma emits any of these depending on how the break was typed.
  const BREAKS = new Set([10, 13, 11, 8232]);
  for (let i = 0; i < chars.length; i++) {
    const code = chars.charCodeAt(i);
    if (BREAKS.has(code)) {
      if (cur) { segs.push(cur); cur = null; }
      if (code === 13 && chars.charCodeAt(i + 1) === 10) i++; // CRLF is one break
      segs.push({ br: true });
      continue;
    }
    const key = overrides[i] || 0;
    if (!cur || cur.key !== key) {
      if (cur) segs.push(cur);
      cur = { key, text: '', style: styleAt(i) };
    }
    cur.text += chars[i];
  }
  if (cur) segs.push(cur);

  // convert to runs; a break marks the previous run (or an empty spacer run)
  const runs = [];
  for (const s of segs) {
    if (s.br) {
      const last = runs[runs.length - 1];
      if (last && !last.breakLine) last.breakLine = true;
      else runs.push({ text: '', style: last ? last.style : base, breakLine: true });   // an empty line
      continue;
    }
    runs.push({ text: s.text, style: s.style, breakLine: false });
  }
  return runs.length ? splitGradientRuns(runs, node) : [{ text: '', style: base, breakLine: false }];
}

const lerp = (a, b, t) => a + (b - a) * t;

function gradientColor(stops, t) {
  const last = stops[stops.length - 1];
  if (t <= stops[0].position) return stops[0].color;
  if (t >= last.position) return last.color;
  const i = stops.findIndex((s) => s.position >= t);
  const a = stops[i - 1], b = stops[i];
  const u = (t - a.position) / (b.position - a.position || 1);
  return { r: lerp(a.color.r, b.color.r, u), g: lerp(a.color.g, b.color.g, u), b: lerp(a.color.b, b.color.b, u), a: 1 };
}

/** Rough advance in px; good enough to find where a letter falls on a gradient. */
function advance(ch, st) {
  const wide = ch.charCodeAt(0) > 0x2E7F;
  return st.fontSize * (ch === ' ' ? 0.28 : wide ? 0.92 : 0.56) + (st.letterSpacing || 0);
}

/**
 * PPT text can't carry a gradient through our writer, so colour each letter from
 * the gradient at its own position. The text stays editable and keeps the look;
 * Figma lays the gradient across the whole box, not across the words.
 */
function splitGradientRuns(runs, node) {
  const boxW = (node.absoluteBoundingBox || {}).width || 0;
  const align = (node.style || {}).textAlignHorizontal;
  const out = [];
  for (const run of runs) {
    const paint = (run.style.fills || []).find((f) => f.visible !== false && f.type === 'GRADIENT_LINEAR');
    if (!paint || !run.text || !boxW) { out.push(run); continue; }
    const chars = [...run.text];
    const total = chars.reduce((n, c) => n + advance(c, run.style), 0);
    let x = align === 'CENTER' ? (boxW - total) / 2 : align === 'RIGHT' ? boxW - total : 0;
    const [h0, h1] = paint.gradientHandlePositions;
    chars.forEach((c, i) => {
      const w = advance(c, run.style);
      const u = (x + w / 2) / boxW;
      x += w;
      const color = gradientColor(paint.gradientStops, lerp(h0.x, h1.x, u));
      out.push({
        text: c,
        style: { ...run.style, fills: [{ type: 'SOLID', color: { r: color.r, g: color.g, b: color.b, a: 1 } }] },
        breakLine: i === chars.length - 1 && run.breakLine,
      });
    });
  }
  return out;
}

/**
 * @param frameNode  a Figma FRAME (one slide)
 * @param opts.rasterIds  Set that collects node ids that must be rendered as images
 * @param opts.notes      array collecting fidelity warnings
 */
function normalizeFrame(frameNode, opts = {}) {
  const rasterIds = opts.rasterIds || new Set();
  const imageRefs = opts.imageRefs || new Set();
  const notes = opts.notes || [];
  const fb = box(frameNode);
  if (!fb) throw new Error(`프레임에 좌표가 없습니다: ${frameNode.name}`);

  const ops = [];
  const origin = { x: fb.x, y: fb.y };
  const local = (b) => ({ x: b.x - origin.x, y: b.y - origin.y, w: b.width, h: b.height });

  const frameBg = solidPaint(frameNode.fills, 1) || { r: 255, g: 255, b: 255, a: 1 };

  // The frame's own paint only ever became a background colour, so a frame whose
  // fill IS the artwork — a PDF page imported into Figma arrives exactly like
  // this: one FRAME, an IMAGE fill, no children — converted to a blank white
  // slide. walk() never runs on the frame itself, so nothing else could catch it.
  //
  // Only safe to rasterise when there are no children: Figma renders a node
  // together with its contents, so doing this to a frame that has children would
  // bake them into the picture and destroy the editability that is the whole
  // point of this tool.
  if (hasNonSolidPaint(frameNode.fills)) {
    const full = { x: 0, y: 0, w: fb.width, h: fb.height };
    const photo = (frameNode.fills || []).find((p) => p.type === 'IMAGE' && p.visible !== false && p.imageRef);
    if (!(frameNode.children || []).length) {
      rasterIds.add(frameNode.id);
      ops.push({ kind: 'image', id: frameNode.id, ...full, srcRect: { ...full } });
    } else if (photo) {
      // Rendering this frame would flatten its children into the picture, so
      // fetch the original upload by imageRef instead and lay it in behind them.
      const id = `ref:${photo.imageRef}`;
      imageRefs.add(photo.imageRef);
      ops.push({ kind: 'image', id, ...full, srcRect: { ...full } });
    } else {
      notes.push(`${frameNode.name}: 프레임 배경(그라디언트)은 옮기지 못했습니다 (자식 요소의 편집성을 지키기 위해).`);
    }
  }

  function walk(node, clip, bg, inheritedOpacity) {
    if (node.visible === false) return;
    const b = box(node);
    if (!b) return;

    const nodeOpacity = (node.opacity === undefined ? 1 : node.opacity) * inheritedOpacity;
    const rect = local(b);
    const visible = clipRect(rect, clip);
    if (!visible) return; // fully clipped away in Figma → must not appear in PPT either

    const fills = node.fills;
    // A PPT rect has no angle here: a tilted shape would come out axis-aligned.
    const tilted = Math.abs(node.rotation || 0) > 0.001 && !(node.children || []).length && node.type !== 'TEXT';
    const needsRaster =
      RASTER_TYPES.has(node.type) ||
      tilted ||
      (node.type === 'LINE' && node.strokeCap && node.strokeCap !== 'NONE') ||   // arrowheads
      (hasNonSolidPaint(fills) && !(CONTAINER_TYPES.has(node.type) && (node.children || []).length)) ||
      (Array.isArray(node.effects) && node.effects.some((e) => e.visible !== false) && !(node.children || []).length);

    if (needsRaster) {
      rasterIds.add(node.id);
      const rb = node.absoluteRenderBounds;
      ops.push({
        kind: 'image', id: node.id, ...visible, srcRect: rect,
        // Where the render really lands: for a tilted picture or a line with an
        // arrowhead this differs from the bounding box, sometimes in height 0 vs 29.
        render: rb ? local(rb) : null,
        clip: clip ? { ...clip } : null,
      });
      return;
    }

    if (node.type === 'TEXT') {
      ops.push({
        kind: 'text',
        id: node.id,
        ...rect,
        clip: clip ? { ...clip } : null,
        runs: buildRuns(node),
        // per-paragraph list data — Figma stores it beside the runs, not in them
        lineTypes: node.lineTypes || null,
        lineIndentations: node.lineIndentations || null,
        style: node.style || {},
        baseFill: solidPaint(node.fills, nodeOpacity),
        bg,
        opacity: nodeOpacity,
      });
      return;
    }

    // --- background fill -------------------------------------------------
    let childBg = bg;
    let fillColor = solidPaint(fills, nodeOpacity);
    if (!fillColor && hasNonSolidPaint(fills)) {
      const avg = gradientAverage(fills);
      if (avg) {
        fillColor = { ...avg, a: nodeOpacity };
        notes.push(`${node.name}: 그라디언트를 단색으로 근사했습니다 (자식 요소 편집성 유지)`);
      }
    }
    if (fillColor) {
      const solid = blend(fillColor, bg);
      childBg = solid;
      ops.push({
        kind: node.type === 'ELLIPSE' ? 'ellipse' : 'rect',
        id: node.id,
        ...visible,
        fill: solid,
        radius: cornerRadius(node),
      });
    }

    if (Array.isArray(node.effects) && node.effects.some((e) => e.visible !== false)) {
      notes.push(`${node.name}: 그림자/블러 효과는 생략했습니다`);
    }

    // --- strokes ---------------------------------------------------------
    emitStrokes(node, visible, rect, bg, nodeOpacity, ops);

    // --- children --------------------------------------------------------
    const nextClip = node.clipsContent ? intersect(clip, rect) : clip;
    for (const c of node.children || []) walk(c, nextClip, childBg, nodeOpacity);
  }

  function emitStrokes(node, visible, rect, bg, nodeOpacity, out) {
    const sc = solidPaint(node.strokes, nodeOpacity);
    if (!sc) return;
    const color = blend(sc, bg);
    const iw = node.individualStrokeWeights;
    const w = node.strokeWeight === undefined ? 1 : node.strokeWeight;

    // zero-height/width nodes are rules, not outlines
    if (rect.h <= 0.5 || rect.w <= 0.5) {
      out.push({ kind: 'line', id: node.id, x: visible.x, y: visible.y, w: visible.w, h: visible.h, color, width: w || 1 });
      return;
    }

    if (iw && (iw.top !== iw.right || iw.right !== iw.bottom || iw.bottom !== iw.left)) {
      // per-side borders: table cells rely on these, and a rectangle outline cannot express them
      const seg = (x, y, ww, hh, t) =>
        t > 0 && out.push({ kind: 'line', id: node.id, x, y, w: ww, h: hh, color, width: t });
      const r = rect;
      seg(r.x, r.y, 0, r.h, iw.left);
      seg(r.x, r.y, r.w, 0, iw.top);
      seg(r.x + r.w, r.y, 0, r.h, iw.right);
      seg(r.x, r.y + r.h, r.w, 0, iw.bottom);
      return;
    }

    if (w > 0) {
      out.push({
        kind: 'outline',
        id: node.id,
        ...visible,
        color,
        width: iw ? iw.top || w : w,
        radius: cornerRadius(node),
      });
    }
  }

  for (const c of frameNode.children || []) {
    walk(c, frameNode.clipsContent === false ? null : { x: 0, y: 0, w: fb.width, h: fb.height }, frameBg, 1);
  }

  // A frame with a plain colour and nothing in it is a deliberately blank page,
  // not a conversion failure — the background already carries everything it has.
  const blank = !ops.length && !(frameNode.children || []).length && !hasNonSolidPaint(frameNode.fills);
  return { name: frameNode.name, id: frameNode.id, w: fb.width, h: fb.height, bg: frameBg, ops, blank };
}

function cornerRadius(node) {
  if (Array.isArray(node.rectangleCornerRadii)) return Math.max(...node.rectangleCornerRadii) || 0;
  return node.cornerRadius || 0;
}

module.exports = { normalizeFrame };
