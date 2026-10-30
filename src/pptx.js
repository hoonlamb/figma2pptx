'use strict';
// Draw list -> .pptx. Everything except rasterized nodes becomes a native,
// editable PowerPoint object.

const path = require('path');
const PptxGenJS = require('pptxgenjs');
const { hex } = require('./color');

const ALIGN = { LEFT: 'left', CENTER: 'center', RIGHT: 'right', JUSTIFIED: 'justify' };
const VALIGN = { TOP: 'top', CENTER: 'middle', BOTTOM: 'bottom' };

async function buildPptx(slides, assets, cfg, outPath, meta = {}) {
  const first = slides[0];
  if (!first) throw new Error('변환할 프레임이 없습니다.');

  const slideW = cfg.slideWidthInches || 13.3333;
  const slideH = +(slideW * (first.h / first.w)).toFixed(4);
  const PX2IN = slideW / first.w;
  const PX2PT = PX2IN * 72;

  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'FIGMA', width: slideW, height: slideH });
  pres.layout = 'FIGMA';
  if (meta.title) pres.title = meta.title;
  pres.author = 'figma2pptx';

  const warnings = [];
  const measurer = cfg.__measurer;
  const drift = [];               // texts whose rendered width ≠ the design's

  // Pre-measure every run in one batch: the layout decisions below need the
  // width of the font PowerPoint will use, not the one Figma measured with.
  if (measurer && measurer.available) {
    for (const s of slides)
      for (const op of s.ops)
        if (op.kind === 'text') for (const r of runsOf(op, PX2PT, cfg)) measurer.want(r);
    measurer.flush();
  }

  for (const s of slides) {
    const slide = pres.addSlide();
    slide.background = { color: hex(s.bg) };

    for (const op of s.ops) {
      try {
        drawOp(slide, op, { PX2IN, PX2PT, assets, cfg, warnings, measurer, drift, page: s.name, slideW: s.w });
      } catch (e) {
        warnings.push(`${s.name} / ${op.kind} ${op.id}: ${e.message}`);
      }
    }
  }

  await pres.writeFile({ fileName: outPath });
  return { slideW, slideH, warnings, drift, path: path.resolve(outPath) };
}

/** The runs of a text op, described the way the measurer wants them. */
function runsOf(op, PX2PT, cfg) {
  const resolver = cfg.__resolver;
  return op.runs.map((r) => {
    const rs = r.style || op.style || {};
    const { face, bold } = resolver.resolve(rs);
    return {
      text: r.text,
      face,
      bold,
      italic: !!rs.italic,
      size: rs.fontSize || 16,
      kern: rs.letterSpacing || 0,
      breakLine: r.breakLine,
    };
  });
}

/** Widest rendered line, in px, or null when unmeasured. */
function renderedWidth(op, ctx) {
  const { measurer, PX2PT, cfg } = ctx;
  if (!measurer || !measurer.available) return null;
  let best = 0, line = 0, seen = false;
  for (const r of runsOf(op, PX2PT, cfg)) {
    const w = measurer.width(r);
    if (w === null) return null;
    seen = true;
    line += w;
    if (r.breakLine) { best = Math.max(best, line); line = 0; }
  }
  return seen ? Math.max(best, line) : null;
}

function drawOp(slide, op, ctx) {
  const { PX2IN } = ctx;
  const geo = { x: op.x * PX2IN, y: op.y * PX2IN, w: Math.max(op.w, 0) * PX2IN, h: Math.max(op.h, 0) * PX2IN };

  switch (op.kind) {
    case 'rect':
    case 'ellipse': {
      const o = { ...geo, fill: { color: hex(op.fill) }, line: { type: 'none' } };
      if (op.kind === 'ellipse') return slide.addShape('ellipse', o);
      if (op.radius > 0) {
        o.rectRadius = Math.min(op.radius * PX2IN, Math.min(geo.w, geo.h) / 2);
        return slide.addShape('roundRect', o);
      }
      return slide.addShape('rect', o);
    }

    case 'outline': {
      const o = {
        ...geo,
        fill: { type: 'none' },
        line: { color: hex(op.color), width: op.width * ctx.PX2PT },
      };
      if (op.radius > 0) {
        o.rectRadius = Math.min(op.radius * PX2IN, Math.min(geo.w, geo.h) / 2);
        return slide.addShape('roundRect', o);
      }
      return slide.addShape('rect', o);
    }

    case 'line':
      return slide.addShape('line', {
        ...geo,
        line: { color: hex(op.color), width: Math.max(op.width, 0.25) * ctx.PX2PT },
      });

    case 'image': {
      const asset = ctx.assets[op.id];
      if (!asset) {
        ctx.warnings.push(`이미지 렌더를 찾지 못했습니다: ${op.id}`);
        return;
      }
      // assets may be a bare base64 string (older callers, tests) or {data, mime}
      const data = typeof asset === 'string' ? asset : asset.data;
      const mime = typeof asset === 'string' ? 'image/png' : asset.mime;
      const placed = placeRender(op, asset, ctx);
      if (placed) return slide.addImage({ ...placed, data: `${mime};base64,${data}` });
      return slide.addImage({ ...geo, data: `${mime};base64,${data}` });
    }

    case 'text':
      return drawText(slide, op, ctx);

    default:
      return;
  }
}

/**
 * Figma exports a node's whole render, while absoluteRenderBounds stops at the
 * clipping frame. Fitting the one into the other squeezes a tilted picture, so
 * find where the full render sits, keep it at its real size and crop the part
 * the frame hides. Returns addImage geometry, or null when there is nothing to go on.
 */
function placeRender(op, asset, ctx) {
  const rb = op.render;
  if (!rb || !asset.width || !asset.height) return null;
  const { PX2IN, slideW } = ctx;
  const clip = op.clip || { x: 0, y: 0, w: slideW || rb.x + rb.w, h: Infinity };

  const cut = (lo, size, clipLo, clipHi) => ({ start: lo <= clipLo + 0.5, end: lo + size >= clipHi - 0.5 });
  const cx = cut(rb.x, rb.w, clip.x, clip.x + clip.w);
  const cy = cut(rb.y, rb.h, clip.y, clip.y + clip.h);

  // Figma may render below the requested scale on huge nodes, so prefer an axis
  // the clip did not touch to learn the real one.
  let scale = ctx.cfg.raster.scale || 1;
  if (!cx.start && !cx.end && rb.w >= 1) scale = asset.width / rb.w;
  else if (!cy.start && !cy.end && rb.h >= 1) scale = asset.height / rb.h;
  const w = asset.width / scale, h = asset.height / scale;

  const anchor = (lo, size, c, full, centre) =>
    !c.start ? lo : !c.end ? lo + size - full : centre - full / 2;
  const x = anchor(rb.x, rb.w, cx, w, op.srcRect.x + op.srcRect.w / 2);
  const y = anchor(rb.y, rb.h, cy, h, op.srcRect.y + op.srcRect.h / 2);

  const left = Math.max(x, clip.x), top = Math.max(y, clip.y);
  const right = Math.min(x + w, clip.x + clip.w), bottom = Math.min(y + h, clip.y + clip.h);
  const geo = { x: left * PX2IN, y: top * PX2IN, w: w * PX2IN, h: h * PX2IN };
  if (left - x < 0.5 && top - y < 0.5 && x + w - right < 0.5 && y + h - bottom < 0.5) {
    return { x: x * PX2IN, y: y * PX2IN, w: geo.w, h: geo.h };
  }
  return {
    ...geo,
    sizing: { type: 'crop', x: (left - x) * PX2IN, y: (top - y) * PX2IN, w: (right - left) * PX2IN, h: (bottom - top) * PX2IN },
  };
}

/** One px height per paragraph: a box can mix sizes, and each line keeps its own. */
function paragraphLineHeights(op) {
  const st = op.style || {};
  const height = (rs) => {
    const size = rs.fontSize || st.fontSize;
    if (rs.lineHeightUnit === 'FONT_SIZE_%' && rs.lineHeightPercentFontSize) return size * rs.lineHeightPercentFontSize / 100;
    if (rs.lineHeightUnit === 'INTRINSIC_%' && st.fontSize && rs.lineHeightPx === st.lineHeightPx) return st.lineHeightPx * size / st.fontSize;
    return rs.lineHeightPx || 0;
  };
  const out = [];
  let best = 0;
  op.runs.forEach((r, i) => {
    best = Math.max(best, height(r.style || st));
    if (r.breakLine || i === op.runs.length - 1) { out.push(best); best = 0; }
  });
  return out;
}

function drawText(slide, op, ctx) {
  const { PX2IN, PX2PT, cfg } = ctx;
  const st = op.style || {};
  const resolver = ctx.cfg.__resolver;

  // Figma keeps list formatting outside the character runs, in per-paragraph
  // arrays. Dropping them loses both the bullet glyph and the indent it sits in,
  // which shifts the whole line left — the element check caught exactly that.
  const lineTypes = op.lineTypes || [];
  const lineIndents = op.lineIndentations || [];
  const lineHeights = paragraphLineHeights(op);
  let paragraph = 0;

  const runs = op.runs.map((r, i) => {
    const rs = r.style || st;
    const { face, bold } = resolver.resolve(rs);
    const sizePt = (rs.fontSize || 16) * PX2PT;
    const colorSrc = runColor(r, op);
    const first = i === 0 || op.runs[i - 1].breakLine;   // opens a paragraph
    const kind = lineTypes[paragraph];
    const indent = lineIndents[paragraph] || 0;
    const lineHeight = lineHeights[paragraph];
    if (r.breakLine) paragraph++;

    const options = {
      fontFace: face,
      bold,
      italic: !!rs.italic,
      fontSize: sizePt,
      color: hex(colorSrc),
      charSpacing: (rs.letterSpacing || 0) * PX2PT,
      breakLine: r.breakLine,
    };
    if (cfg.text.useExactLineSpacing && lineHeight) options.lineSpacing = lineHeight * PX2PT;
    // Underlines and strikethroughs are ordinary character formatting in Figma
    // and were being dropped on the floor — 208 of them in one real file.
    if (rs.textDecoration === 'UNDERLINE') options.underline = { style: 'sng' };
    else if (rs.textDecoration === 'STRIKETHROUGH') options.strike = 'sngStrike';
    if (first && (kind === 'UNORDERED' || kind === 'ORDERED')) {
      options.bullet = kind === 'ORDERED' ? { type: 'number' } : true;
      if (indent > 1) options.indentLevel = indent - 1;
    }
    return { text: r.text, options };
  });

  // Figma already decided where the lines fall. Re-wrapping in PowerPoint is the
  // single most common way a converted deck breaks, so only let it wrap when the
  // Figma node itself was set to a fixed width.
  const autoResize = st.textAutoResize || 'NONE';
  const canWrap = autoResize !== 'WIDTH_AND_HEIGHT';

  // The box has to clear the text as PowerPoint sets it. Measuring beats the
  // fixed guess: when the render font is wider than the design font a fixed pad
  // silently re-wraps the line, and when it is narrower the slack throws off
  // anything that reads the box (centred labels, chips drawn to fit).
  // THE RULE: a wrapping box keeps exactly the width Figma gave it. That width is
  // the layout constraint — every line break, the row height it produces, and
  // anything drawn beside it all follow from it. Widening it by even a few pixels
  // moves a word to another line and drags the rest of the page with it.
  //
  // This is not hypothetical. renderedWidth() measures the text laid out on ONE
  // line, so feeding a wrapping paragraph through it returned the width of the
  // whole thing unbroken — a 700px table cell asked for a 3006px box. The text
  // stopped wrapping, ran off the slide, and centred cells jumped 1150px left.
  // Only a hugging box (WIDTH_AND_HEIGHT) may be resized, because there the
  // node's width means "as wide as the text" and nothing else is anchored to it.
  let pad = 0;
  if (!canWrap) {
    const real = renderedWidth(op, ctx);
    pad = cfg.text.padNoWrapPx || 0;
    if (real !== null) {
      const delta = real - op.w;                     // >0 → renders wider than designed
      pad = Math.max(4, Math.ceil(delta) + 6);
      if (Math.abs(delta) >= (cfg.text.driftReportPx || 2)) {
        ctx.drift.push({
          page: ctx.page, id: op.id, delta,
          text: op.runs.map((r) => r.text).join('').slice(0, 28),
        });
      }
    }
  }

  const align = ALIGN[st.textAlignHorizontal] || 'left';
  // The slack exists to stop PowerPoint re-wrapping a line; it must not push the
  // box past the slide edge. A full-width centred heading was landing at
  // -24…1945px on a 1920px slide purely because of the padding, which is not a
  // real overflow but is indistinguishable from one in the report.
  const room = ctx.slideW || 0;
  if (room && pad > 0) {
    const grow = align === 'center' ? pad / 2 : align === 'right' ? pad : 0;
    const left = op.x - grow;
    const right = left + op.w + pad;
    const spare = Math.min(left, room - right);
    if (spare < 0) pad = Math.max(0, pad + spare * (align === 'center' ? 2 : 1));
  }

  let x = op.x;
  let w = op.w + pad;
  // Keep the text's optical position: Figma centres/right-aligns inside the node
  // box, so the widened PowerPoint box has to grow around that same anchor.
  if (align === 'center') x -= pad / 2;
  else if (align === 'right') x -= pad;

  // A box that leaves the slide is always a conversion fault, never a design —
  // Figma clips at the frame edge, so nothing the designer drew can be out there.
  // Say so instead of shipping a page with the sentence sliced off.
  const slideW = ctx.slideW || 0;
  if (slideW && (x < -1 || x + w > slideW + 1)) {
    ctx.warnings.push(
      `${ctx.page}: «${op.runs.map((r) => r.text).join('').slice(0, 20)}» 텍스트 상자가 `
      + `슬라이드 밖으로 나갑니다 (${Math.round(x)}~${Math.round(x + w)}px / 폭 ${Math.round(slideW)}px)`);
  }

  const opts = {
    x: x * PX2IN,
    y: op.y * PX2IN,
    w: w * PX2IN,
    h: Math.max(op.h, 1) * PX2IN,
    align,
    valign: VALIGN[st.textAlignVertical] || 'top',
    margin: 0,
    wrap: canWrap,
    isTextBox: true,
    autoFit: false,
    shrinkText: false,
  };

  if (cfg.text.useExactLineSpacing && st.lineHeightPx) {
    opts.lineSpacing = st.lineHeightPx * PX2PT;
  }

  slide.addText(runs, opts);
}

function runColor(run, op) {
  const rs = run.style || {};
  const fills = rs.fills;
  if (Array.isArray(fills)) {
    const p = fills.find((f) => f.visible !== false && f.type === 'SOLID');
    if (p) {
      const { toRgb, blend } = require('./color');
      return blend(toRgb(p.color, p.opacity === undefined ? 1 : p.opacity, op.opacity || 1), op.bg);
    }
  }
  const { blend } = require('./color');
  return op.baseFill ? blend(op.baseFill, op.bg) : { r: 0, g: 0, b: 0, a: 1 };
}

module.exports = { buildPptx };
