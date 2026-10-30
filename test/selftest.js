#!/usr/bin/env node
'use strict';
// Offline self-test: builds a synthetic Figma-shaped tree that reproduces every
// failure mode we hit converting real decks, then asserts the normalizer's output
// and writes a real .pptx for visual inspection.
//
//   node test/selftest.js [--pptx]

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { normalizeFrame } = require('../src/normalize');
const { buildPptx } = require('../src/pptx');
const { loadConfig } = require('../src/pipeline');
const { makeResolver } = require('../src/fonts');

// The frame deliberately does NOT sit at the canvas origin — absolute-vs-local
// mistakes stay invisible when the origin is (0,0).
const OX = 1000, OY = 2000;
const bb = (x, y, w, h) => ({ x: OX + x, y: OY + y, width: w, height: h });
const solid = (r, g, b, a = 1) => [{ type: 'SOLID', color: { r: r / 255, g: g / 255, b: b / 255, a } }];

const NAVY = solid(0, 20, 147);
const PBLUE = solid(221, 233, 252);

const text = (id, x, y, w, h, characters, extra = {}) => {
  const { style: styleOverride, ...rest } = extra;
  return {
    id, name: characters.slice(0, 12), type: 'TEXT', absoluteBoundingBox: bb(x, y, w, h),
    characters, fills: solid(0, 0, 0),
    ...rest,
    style: {
      fontFamily: 'Pretendard', fontWeight: 500, fontSize: 24,
      textAlignHorizontal: 'LEFT', textAlignVertical: 'TOP',
      letterSpacing: -1.2, lineHeightPx: 43.2, textAutoResize: 'WIDTH_AND_HEIGHT',
      ...(styleOverride || {}),
    },
  };
};

const frame = {
  id: '0:1', name: 'Self Test', type: 'FRAME', absoluteBoundingBox: bb(0, 0, 1920, 1080),
  fills: solid(255, 255, 255), clipsContent: true,
  children: [
    // 1) GROUP — children carry absolute coords; a naive walker double-shifts them
    {
      id: '1:1', name: 'Group 1', type: 'GROUP', absoluteBoundingBox: bb(143, 215, 400, 40),
      children: [
        { id: '1:2', name: 'bullet', type: 'RECTANGLE', absoluteBoundingBox: bb(143, 227, 16, 16), fills: NAVY },
        text('1:3', 175, 215, 360, 40, '그룹 안의 텍스트'),
      ],
    },

    // 2) clipsContent — the bar is taller than its card and must be cut off
    {
      id: '2:1', name: 'card', type: 'FRAME', absoluteBoundingBox: bb(143, 300, 520, 260),
      fills: PBLUE, clipsContent: true, cornerRadius: 8,
      children: [
        { id: '2:2', name: 'accent bar', type: 'RECTANGLE', absoluteBoundingBox: bb(143, 300, 6, 320), fills: NAVY },
        text('2:3', 179, 330, 300, 34, '클리핑 카드', { style: { fontWeight: 700, fontSize: 28, lineHeightPx: 42 } }),
      ],
    },

    // 3) mixed character colors inside one text node
    {
      ...text('3:1', 700, 300, 700, 48, '6. 에셋 제작 가이드 ｜ 저작권 관리'),
      style: {
        fontFamily: 'Pretendard', fontWeight: 700, fontSize: 40, textAlignHorizontal: 'LEFT',
        textAlignVertical: 'TOP', letterSpacing: -2, lineHeightPx: 48, textAutoResize: 'WIDTH_AND_HEIGHT',
      },
      characterStyleOverrides: Array.from({ length: 40 }, (_, i) => (i >= 13 && i <= 13 ? 1 : 0)),
      styleOverrideTable: { 1: { fills: solid(48, 85, 145) } },
    },

    // 4) per-side borders — what table cells are made of
    {
      id: '4:1', name: 'cell', type: 'FRAME', absoluteBoundingBox: bb(700, 400, 400, 80),
      fills: [], strokes: solid(221, 233, 252), strokeWeight: 1,
      individualStrokeWeights: { top: 1, right: 0, bottom: 0, left: 1 },
      children: [text('4:2', 712, 420, 300, 24, '표 셀', { style: { fontSize: 16, lineHeightPx: 24 } })],
    },

    // 5) translucent rule on white — PPT gets a pre-blended solid
    {
      id: '5:1', name: 'rule', type: 'LINE', absoluteBoundingBox: bb(143, 620, 800, 0),
      strokes: solid(48, 85, 145), strokeWeight: 1, opacity: 0.4,
    },

    // 6) ellipse
    { id: '6:1', name: 'dot', type: 'ELLIPSE', absoluteBoundingBox: bb(143, 660, 48, 48), fills: NAVY },

    // 7) long single-line text that PowerPoint would otherwise re-wrap
    text('7:1', 143, 740, 1600, 50,
      '기획 → 초고 → 피드백 → 수정 → 최종컨펌 → 발행(CMS) — 이 줄은 절대 두 줄로 깨지면 안 됩니다',
      { style: { fontSize: 28, lineHeightPx: 50.4, letterSpacing: -1.4 } }),

    // 8) fixed-width text that SHOULD wrap
    text('8:1', 143, 820, 700, 120,
      '이 문장은 고정 폭 텍스트라 자동 줄바꿈이 허용되어야 합니다. 피그마에서 이미 여러 줄로 흐르고 있기 때문입니다.',
      { style: { fontSize: 16, lineHeightPx: 24, textAutoResize: 'HEIGHT' } }),

    // 9) invisible node — must not reach the deck
    { id: '9:1', name: 'hidden', type: 'RECTANGLE', visible: false, absoluteBoundingBox: bb(0, 0, 1920, 1080), fills: NAVY },
  ],
};

// ---------------------------------------------------------------------------
const notes = [], rasterIds = new Set();
const slide = normalizeFrame(frame, { notes, rasterIds });
const find = (id) => slide.ops.filter((o) => o.id === id);

let passed = 0;
const check = (label, fn) => {
  try { fn(); console.log(`  ✅ ${label}`); passed++; }
  catch (e) { console.log(`  ❌ ${label}\n     ${e.message}`); process.exitCode = 1; }
};

console.log('\nfigma2pptx 자체 테스트\n');

check('그룹 자식 좌표가 밀리지 않는다', () => {
  const t = find('1:3')[0];
  assert.strictEqual(Math.round(t.x), 175, `x=${t.x}, 기대값 175`);
  assert.strictEqual(Math.round(t.y), 215, `y=${t.y}, 기대값 215`);
});

check('클리핑 프레임 밖으로 넘친 자식이 잘린다', () => {
  const bar = find('2:2')[0];
  assert.strictEqual(Math.round(bar.h), 260, `높이=${bar.h}, 기대값 260 (원본 320)`);
});

check('숨긴 노드는 출력되지 않는다', () => {
  assert.strictEqual(find('9:1').length, 0);
});

check('한 텍스트 안의 혼합 색상이 런으로 분리된다', () => {
  const t = find('3:1')[0];
  const colored = t.runs.filter((r) => r.style && Array.isArray(r.style.fills));
  assert.ok(t.runs.length >= 3, `런 개수=${t.runs.length}`);
  assert.ok(colored.length >= 1, '색이 다른 런을 찾지 못했습니다');
  assert.strictEqual(colored[0].text, '｜');
});

check('셀의 좌·상 테두리만 선으로 나온다', () => {
  const lines = find('4:1').filter((o) => o.kind === 'line');
  assert.strictEqual(lines.length, 2, `선 개수=${lines.length}`);
  assert.ok(lines.some((l) => l.w === 0), '세로선 없음');
  assert.ok(lines.some((l) => l.h === 0), '가로선 없음');
});

check('반투명 선이 배경과 미리 합성된다', () => {
  const l = find('5:1')[0];
  assert.strictEqual(l.color.a, 1, '알파가 남아 있습니다');
  assert.ok(l.color.r > 100 && l.color.r < 200, `합성 결과 R=${l.color.r}`);
});

check('타원이 타원으로 유지된다', () => {
  assert.strictEqual(find('6:1')[0].kind, 'ellipse');
});

check('자동 폭 텍스트는 줄바꿈이 꺼진다', () => {
  const t = find('7:1')[0];
  assert.strictEqual(t.style.textAutoResize, 'WIDTH_AND_HEIGHT');
});

check('고정 폭 텍스트는 줄바꿈이 켜진다', () => {
  const t = find('8:1')[0];
  assert.strictEqual(t.style.textAutoResize, 'HEIGHT');
});

// 10) the failure mode behind three rounds of client revisions: the deck carried
//     Noto Sans KR widths but rendered in Pretendard, so every chip and centred
//     label sat off. The converter has to notice and report that itself.
const { createMeasurer } = require('../src/metrics');

async function driftTest() {
  const cfg2 = loadConfig();
  cfg2.__resolver = makeResolver(cfg2);
  const m = createMeasurer(() => {});
  if (!m.available) {
    console.log('  ⏭️ 서체 폭 차이 감지 (실측 불가 환경 — 건너뜀)');
    return true;
  }
  cfg2.__measurer = m;

  // width as Noto measured it; Pretendard sets the same string ~3.6px narrower
  const drifted = {
    name: 'drift', id: 'd', w: 1920, h: 1080, bg: { r: 255, g: 255, b: 255 },
    ops: [{
      kind: 'text', id: 'd:1', x: 100, y: 100, w: 56.5, h: 20,
      runs: [{ text: '공통사항', style: null, breakLine: false }],
      style: { fontFamily: 'Pretendard', fontWeight: 700, fontSize: 16,
               textAlignHorizontal: 'LEFT', textAlignVertical: 'TOP',
               letterSpacing: -0.8, lineHeightPx: 20, textAutoResize: 'WIDTH_AND_HEIGHT' },
      baseFill: { r: 0, g: 0, b: 0, a: 1 }, bg: { r: 255, g: 255, b: 255 }, opacity: 1,
    }],
  };
  const out = path.join(__dirname, '_drift.pptx');
  const r = await buildPptx([drifted], {}, cfg2, out, {});
  require('fs').unlinkSync(out);
  try {
    assert.strictEqual(r.drift.length, 1, `drift 항목 ${r.drift.length}개`);
    assert.ok(r.drift[0].delta < -2, `delta=${r.drift[0].delta}`);
    console.log('  ✅ 서체 폭 차이를 감지해 리포트한다 ' +
                `(${r.drift[0].delta.toFixed(1)}px)`);
    return true;
  } catch (e) {
    console.log(`  ❌ 서체 폭 차이를 감지해 리포트한다\n     ${e.message}`);
    return false;
  }
}

// 11) the render comparison scaled nothing: Figma rendered 1600x900 while
//     PowerPoint's PDF came out 3200x1800, and padding the smaller one into the
//     corner made a byte-perfect conversion report a 60%+ mismatch.
function scalingTest() {
  try {
    const { PNG } = require('pngjs');
    const pixelmatch = require('pixelmatch');
    // verify.js keeps align() private; run the file with a require rooted in
    // src/ (its own './figma' etc. must resolve) and reach in for the function.
    const SRC = path.join(__dirname, '..', 'src');
    const srcRequire = require('module').createRequire(path.join(SRC, 'verify.js'));
    const vsrc = fs.readFileSync(path.join(SRC, 'verify.js'), 'utf8');
    const m = { exports: {} };
    new Function('module', 'exports', 'require', '__dirname',
      vsrc + '\nmodule.exports.__align = align;')(m, m.exports, srcRequire, SRC);

    // the same picture at two raster sizes — the exact shape of the real failure
    const draw = (size) => {
      const p = new PNG({ width: size, height: size });
      for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
        const i = (y * size + x) << 2;
        const on = x > size * 0.25 && x < size * 0.75 && y > size * 0.25 && y < size * 0.75;
        p.data[i] = p.data[i + 1] = p.data[i + 2] = on ? 255 : 0;
        p.data[i + 3] = 255;
      }
      return p;
    };
    const [A, B, w, h] = m.exports.__align(draw(100), draw(200));
    const d = new PNG({ width: w, height: h });
    const ratio = pixelmatch(A.data, B.data, d.data, w, h, { threshold: 0.12, includeAA: false }) / (w * h);
    assert.ok(ratio <= 0.02, `같은 그림인데 ${(ratio * 100).toFixed(1)}% 차이로 나옴`);
    console.log(`  ✅ 크기가 다른 렌더를 스케일 맞춰 비교한다 (${(ratio * 100).toFixed(2)}%)`);
    return true;
  } catch (e) {
    console.log(`  ❌ 크기가 다른 렌더를 스케일 맞춰 비교한다\n     ${e.message}`);
    return false;
  }
}

// 12) pptxgenjs writes one media file per addImage, so artwork repeated across
//     pages was stored once per use — 114 embeds for 88 distinct pictures, 75MB.
async function dedupeTest() {
  try {
    const os = require('os');
    const cp = require('child_process');
    const { PNG } = require('pngjs');
    const { dedupeMedia } = require('../src/dedupe');

    const mk = (r, g, b) => {
      const p = new PNG({ width: 40, height: 40 });
      for (let i = 0; i < p.data.length; i += 4) {
        p.data[i] = r; p.data[i + 1] = g; p.data[i + 2] = b; p.data[i + 3] = 255;
      }
      return PNG.sync.write(p).toString('base64');
    };
    const red = mk(200, 60, 60), blue = mk(60, 60, 200);
    const plan = [['a', red], ['b', red], ['c', blue], ['d', red], ['e', blue], ['f', red]];
    const assets = {};
    plan.forEach(([id, d]) => { assets[id] = d; });

    const s = {
      name: 'S', id: 's', w: 1920, h: 1080, bg: { r: 255, g: 255, b: 255, a: 1 },
      ops: plan.map(([id], i) => ({
        kind: 'image', id, x: 100 + i * 120, y: 100, w: 100, h: 100,
        srcRect: { x: 100 + i * 120, y: 100, w: 100, h: 100 },
      })),
    };
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    const out = path.join(os.tmpdir(), `dedupe-${Date.now()}.pptx`);
    await buildPptx([s], assets, cfg, out, { title: 'dedupe' });
    const r = await dedupeMedia(out);

    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'z-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const media = fs.readdirSync(path.join(d, 'ppt', 'media'));
    const rels = fs.readFileSync(path.join(d, 'ppt', 'slides', '_rels', 'slide1.xml.rels'), 'utf8');
    const refs = [...rels.matchAll(/media\/([^"]+)/g)].map((x) => x[1]);
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });

    assert.strictEqual(media.length, 2, `고유 그림은 2개여야 하는데 ${media.length}개`);
    assert.strictEqual(refs.length, 6, `이미지 참조 6개가 유지돼야 하는데 ${refs.length}개`);
    const broken = refs.filter((f) => !media.includes(f));
    assert.ok(!broken.length, `끊어진 참조: ${broken.join(' ')}`);
    console.log(`  ✅ 같은 그림은 한 번만 저장한다 (${r.removed}개 제거, 참조 6개 유지)`);
    return true;
  } catch (e) {
    console.log(`  ❌ 같은 그림은 한 번만 저장한다\n     ${e.message}`);
    return false;
  }
}

// 13) the table-cell disaster: renderedWidth() measures text on ONE line, and the
//     old code applied it to wrapping text too — a 700px cell asked for a 3006px
//     box, so the paragraph stopped wrapping, ran off the slide, and centred
//     cells jumped half the overflow to the left.
async function wrapWidthTest() {
  try {
    const os = require('os');
    const cp = require('child_process');
    const long = '협력사의 외부 제작자를 통해 제작한 저작물에 대해서는 협력사가 외주 저작물에 대한 복제권, '
               + '전송권, 2차 저작물 작성권 등을 포함한 일체 저작 재산권을 양도함';
    const cellW = 700;
    const s = {
      name: 'T', id: 't', w: 1920, h: 1080, bg: { r: 255, g: 255, b: 255, a: 1 },
      ops: [{
        kind: 'text', id: 'cell', x: 1100, y: 200, w: cellW, h: 300, clip: null,
        runs: [{ text: long, style: {}, breakLine: false }],
        style: { fontSize: 24, fontFamily: 'Pretendard', textAutoResize: 'HEIGHT', textAlignHorizontal: 'CENTER' },
        baseFill: { r: 0, g: 0, b: 0, a: 1 }, bg: { r: 255, g: 255, b: 255, a: 1 }, opacity: 1,
      }],
    };
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    const out = path.join(os.tmpdir(), `wrap-${Date.now()}.pptx`);
    const built = await buildPptx([s], {}, cfg, out, { title: 'wrap' });

    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'w-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const xml = fs.readFileSync(path.join(d, 'ppt', 'slides', 'slide1.xml'), 'utf8');
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });

    // the first <a:ext> in the file belongs to the empty group wrapper (cx="0"),
    // so read geometry from the shape that actually carries the text body
    const sp = xml.split('<p:sp>').find((chunk) => chunk.includes('<p:txBody>'));
    assert.ok(sp, '텍스트 도형을 찾지 못했습니다');
    const m = sp.match(/<a:off x="(-?\d+)" y="-?\d+"\/><a:ext cx="(\d+)"/);
    assert.ok(m, '텍스트 상자의 좌표를 찾지 못했습니다');
    const EMU = 6350;                                  // per Figma px at 1920 wide
    const gotX = Number(m[1]) / EMU, gotW = Number(m[2]) / EMU;

    assert.ok(Math.abs(gotW - cellW) <= 1, `줄바꿈 상자 폭이 ${Math.round(gotW)}px (설계 ${cellW}px)`);
    assert.ok(Math.abs(gotX - 1100) <= 1, `줄바꿈 상자가 ${Math.round(gotX)}px로 이동했습니다 (설계 1100px)`);
    assert.ok(!built.warnings.some((w) => w.includes('슬라이드 밖')), '슬라이드 밖으로 나갔습니다');
    console.log(`  ✅ 줄바꿈 텍스트는 설계한 폭·위치를 그대로 지킨다 (${Math.round(gotW)}px @ ${Math.round(gotX)}px)`);
    return true;
  } catch (e) {
    console.log(`  ❌ 줄바꿈 텍스트는 설계한 폭·위치를 그대로 지킨다\n     ${e.message}`);
    return false;
  }
}

// 14) a page-level percentage cannot separate "the table moved" from "the text
//     renders narrow" — finding that out took a hand-written script. Measure each
//     box's ink instead, and prove the measurement reacts to each fault alone.
function inkDiffTest() {
  try {
    const { PNG } = require('pngjs');
    const { compareText } = require('../src/inkdiff');

    // white page, one black bar standing in for a line of text
    const page = (x, width) => {
      const p = new PNG({ width: 400, height: 120 });
      p.data.fill(255);
      for (let y = 40; y < 60; y++) {
        for (let px = x; px < x + width; px++) {
          const i = (y * 400 + px) << 2;
          p.data[i] = p.data[i + 1] = p.data[i + 2] = 0; p.data[i + 3] = 255;
        }
      }
      return p;
    };
    const slide = {
      name: 'S', w: 400, h: 120,
      ops: [{ kind: 'text', x: 90, y: 30, w: 220, h: 40, runs: [{ text: '본문 한 줄' }] }],
    };
    const cfg = { elementTolerancePx: 2, elementToleranceRatio: 0.02 };
    const base = page(100, 200);
    const one = (b, s = slide) => {
      const r = compareText(base, b, s, 1, cfg);
      assert.strictEqual(r.length, 1, `측정 결과가 ${r.length}건`);
      return r[0];
    };

    const same = one(page(100, 200));
    assert.ok(Math.abs(same.moveX) < 1.5 && Math.abs(same.dw) < 1.5, '동일한데 차이로 보고');

    const moved = one(page(112, 200));
    assert.ok(Math.abs(moved.moveX - 12) < 1.5, `이동량 ${moved.moveX.toFixed(1)}px (실제 12px)`);
    assert.ok(Math.abs(moved.dw) < 1.5, `이동인데 폭 변화 ${moved.dw.toFixed(1)}px로 보고`);

    const narrow = one(page(100, 180));
    assert.ok(Math.abs(narrow.dw + 20) < 1.5, `폭 변화 ${narrow.dw.toFixed(1)}px (실제 -20px)`);
    assert.ok(Math.abs(narrow.moveX) < 1.5, `폭 변화인데 이동 ${narrow.moveX.toFixed(1)}px로 보고`);

    // centred text that renders narrower pulls its own left edge inward; that is
    // the width difference again, not a move, and must not be counted twice
    const centred = {
      ...slide,
      ops: [{ ...slide.ops[0], style: { textAlignHorizontal: 'CENTER' } }],
    };
    const shrunk = compareText(base, page(110, 180), centred, 1, cfg)[0];
    assert.ok(Math.abs(shrunk.dx - 10) < 1.5, `원시 이동량이 ${shrunk.dx.toFixed(1)}px`);
    assert.ok(Math.abs(shrunk.moveX) < 1.5,
      `가운데 정렬 축소인데 이동 ${shrunk.moveX.toFixed(1)}px로 보고 (0이어야 함)`);

    const blank = new PNG({ width: 400, height: 120 }); blank.data.fill(255);
    const missing = one(blank);
    assert.strictEqual(missing.kind, 'missing', `종류가 ${missing.kind}`);

    console.log('  ✅ 요소별로 이동·폭·누락을 구분해 잡는다 (가운데 정렬 축소는 이동으로 세지 않음)');
    return true;
  } catch (e) {
    console.log(`  ❌ 요소별로 이동·폭·누락을 구분해 잡는다\n     ${e.message}`);
    return false;
  }
}

driftTest().then(async (ok) => {
  if (ok) passed++; else process.exitCode = 1;
  if (scalingTest()) passed++; else process.exitCode = 1;
  if (await dedupeTest()) passed++; else process.exitCode = 1;
  if (await wrapWidthTest()) passed++; else process.exitCode = 1;
  if (inkDiffTest()) passed++; else process.exitCode = 1;
  if (await bulletTest()) passed++; else process.exitCode = 1;
  if (await decorationTest()) passed++; else process.exitCode = 1;
  if (imageFormatTest()) passed++; else process.exitCode = 1;
  if (await retryTest()) passed++; else process.exitCode = 1;
  if (frameFillTest()) passed++; else process.exitCode = 1;
  if (await shapeImageTest()) passed++; else process.exitCode = 1;
  if (await mixedLineHeightTest()) passed++; else process.exitCode = 1;
  if (mixedRunTest()) passed++; else process.exitCode = 1;
  if (localGuardTest()) passed++; else process.exitCode = 1;
  if (safeNameTest()) passed++; else process.exitCode = 1;
  console.log(`\n${passed}/24 통과`);
});

// 19) the blank-page bug, finally reproduced: a PDF page imported into Figma
//     arrives as one FRAME with an IMAGE fill and no children. walk() never runs
//     on the frame itself and the frame's own paint was only ever read for a
//     background colour, so the whole page converted to blank white.
function frameFillTest() {
  try {
    const mkFrame = (id, name, fills, children) => ({
      id, name, type: 'FRAME', clipsContent: true, fills,
      absoluteBoundingBox: { x: -21515, y: -6502, width: 1920, height: 1080 },
      children,
    });
    const imageFill = [{ type: 'IMAGE', scaleMode: 'FILL', imageRef: 'abc123' }];
    const solidFill = [{ type: 'SOLID', color: { r: 1, g: 1, b: 1, a: 1 } }];

    // no children — safe to render the frame itself
    let rasterIds = new Set(), imageRefs = new Set();
    let s = normalizeFrame(mkFrame('1:1', 'PDF Page 1', imageFill, []), { rasterIds, imageRefs });
    assert.strictEqual(s.ops.length, 1, `배경만 있는 프레임의 op가 ${s.ops.length}개`);
    assert.strictEqual(s.ops[0].kind, 'image', '이미지 op가 아닙니다');
    assert.ok(rasterIds.has('1:1'), '프레임 자체를 렌더 목록에 넣지 않았습니다');
    assert.ok(Math.abs(s.ops[0].w - 1920) < 1, '배경이 프레임 전체를 덮지 않습니다');

    // with children — rendering the frame would flatten them, so fetch the
    // original upload by imageRef and keep the children editable
    rasterIds = new Set(); imageRefs = new Set();
    const kid = {
      id: '1:3', name: 'label', type: 'TEXT', characters: '엽서',
      absoluteBoundingBox: { x: -21400, y: -6400, width: 200, height: 40 },
      style: { fontSize: 24 }, fills: [{ type: 'SOLID', color: { r: 0, g: 0, b: 0, a: 1 } }],
    };
    s = normalizeFrame(mkFrame('1:2', 'PDF Page 5', imageFill, [kid]), { rasterIds, imageRefs });
    assert.ok(imageRefs.has('abc123'), 'imageRef를 수집하지 않았습니다');
    assert.ok(!rasterIds.has('1:2'), '자식이 있는데 프레임을 통째로 래스터화했습니다');
    assert.strictEqual(s.ops[0].id, 'ref:abc123', '배경 op가 imageRef를 가리키지 않습니다');
    assert.ok(s.ops.some((o) => o.kind === 'text'), '자식 텍스트가 사라졌습니다');

    // a plain colour with nothing in it is a deliberately blank page
    s = normalizeFrame(mkFrame('1:4', '_재단', solidFill, []), {});
    assert.strictEqual(s.ops.length, 0, '단색 빈 프레임에 op가 생겼습니다');
    assert.ok(s.blank, '의도적으로 빈 페이지를 오류로 표시합니다');

    console.log('  ✅ 프레임 자체가 배경 이미지인 페이지를 살린다 (PDF 가져오기)');
    return true;
  } catch (e) {
    console.log(`  ❌ 프레임 자체가 배경 이미지인 페이지를 살린다\n     ${e.message}`);
    return false;
  }
}

// 18) a real 85-frame section needs 2605 image renders, and one 429 partway
//     through threw away minutes of work. 429 and 5xx are "wait", not "wrong";
//     403 and 404 are "wrong" and retrying them only wastes time.
async function retryTest() {
  const realFetch = global.fetch;
  try {
    let calls = 0;
    const reply = (status, retryAfter) => ({
      ok: status === 200,
      status,
      headers: { get: (k) => (k === 'retry-after' ? retryAfter : null) },
      text: async () => '{"err":"x"}',
      json: async () => ({ ok: true }),
    });

    delete require.cache[require.resolve('../src/figma')];
    const figma = require('../src/figma');
    figma.setRetryLogger(() => {});

    // transient: two refusals then success, honouring Retry-After
    calls = 0;
    global.fetch = async () => { calls++; return reply(calls < 3 ? 429 : 200, '1'); };
    const t0 = Date.now();
    await figma.getFile('K', 'T');
    const waited = Date.now() - t0;
    assert.strictEqual(calls, 3, `429 재시도 후 호출이 ${calls}번`);
    assert.ok(waited >= 1800, `Retry-After 1초를 두 번 기다려야 하는데 ${waited}ms`);

    // permanent: must not burn retries on an answer that will not change
    calls = 0;
    global.fetch = async () => { calls++; return reply(403); };
    await assert.rejects(() => figma.getFile('K', 'T'));
    assert.strictEqual(calls, 1, `403인데 ${calls}번 시도했습니다`);

    // downloads get the same treatment — a socket error among thousands is normal
    calls = 0;
    global.fetch = async () => {
      calls++;
      if (calls < 2) throw new Error('socket hang up');
      return { ok: true, status: 200, arrayBuffer: async () => new ArrayBuffer(4) };
    };
    const buf = await figma.download('http://x/y.png');
    assert.strictEqual(buf.length, 4, '재시도 후 이미지를 받지 못했습니다');

    console.log('  ✅ 일시적 오류는 기다렸다 다시 시도하고, 권한 오류는 즉시 알린다');
    return true;
  } catch (e) {
    console.log(`  ❌ 일시적 오류는 기다렸다 다시 시도하고, 권한 오류는 즉시 알린다\n     ${e.message}`);
    return false;
  } finally {
    global.fetch = realFetch;
    delete require.cache[require.resolve('../src/figma')];
  }
}

// 17) Figma renders everything as PNG. That is right for an icon with a soft
//     edge and wrong for a photograph — one real deck held 41 pictures in 56MB.
//     Transparency is the deciding fact, so it must be read from the pixels,
//     never guessed: dropping a used alpha channel puts a black box on the page.
function imageFormatTest() {
  try {
    const { PNG } = require('pngjs');
    const { encode, isOpaque } = require('../src/image');
    const mk = (w, h, alpha) => {
      const p = new PNG({ width: w, height: h });
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const i = (y * w + x) << 2;
          p.data[i] = (x * 7) % 256; p.data[i + 1] = (y * 5) % 256; p.data[i + 2] = ((x + y) * 3) % 256;
          p.data[i + 3] = (alpha && x < 3 && y < 3) ? 0 : 255;   // three corner pixels
        }
      }
      return PNG.sync.write(p);
    };
    const opaque = mk(900, 900, false), transparent = mk(900, 900, true);
    assert.ok(isOpaque(opaque), '불투명 이미지를 투명으로 판정');
    assert.ok(!isOpaque(transparent), '픽셀 3개의 투명도를 놓쳤습니다');

    const a = encode(opaque, {});
    const b = encode(transparent, {});
    assert.strictEqual(b.mime, 'image/png', '투명한 이미지를 JPEG로 바꿨습니다');
    assert.strictEqual(encode(mk(60, 60, false), {}).mime, 'image/png', '작은 이미지까지 변환했습니다');
    assert.strictEqual(encode(opaque, { jpeg: false }).mime, 'image/png', 'jpeg:false를 무시했습니다');

    // sips is macOS-only; where it is missing the PNG must simply survive
    if (a.mime === 'image/jpeg') {
      assert.ok(a.saved > 0, 'JPEG로 바꿨는데 용량이 줄지 않았습니다');
      console.log(`  ✅ 불투명 사진만 JPEG로 담는다 (${(opaque.length / 1024) | 0}KB → ${(Buffer.from(a.data, 'base64').length / 1024) | 0}KB)`);
    } else {
      console.log('  ✅ 불투명 사진만 JPEG로 담는다 (이 환경엔 sips가 없어 PNG 유지)');
    }
    return true;
  } catch (e) {
    console.log(`  ❌ 불투명 사진만 JPEG로 담는다\n     ${e.message}`);
    return false;
  }
}

// 16) textDecoration is ordinary character formatting in Figma and was simply
//     never read — one real file carried 196 underlines and 12 strikethroughs
//     that all vanished in conversion.
async function decorationTest() {
  try {
    const os = require('os');
    const cp = require('child_process');
    const mk = (id, y, deco) => ({
      kind: 'text', id, x: 100, y, w: 600, h: 60, clip: null,
      runs: [{ text: '밑줄 확인용 문장', style: { textDecoration: deco }, breakLine: false }],
      style: { fontSize: 28, textAutoResize: 'WIDTH_AND_HEIGHT', textAlignHorizontal: 'LEFT' },
      baseFill: { r: 0, g: 0, b: 0, a: 1 }, bg: { r: 255, g: 255, b: 255, a: 1 }, opacity: 1,
    });
    const s = {
      name: 'D', id: 'd', w: 1920, h: 1080, bg: { r: 255, g: 255, b: 255, a: 1 },
      ops: [mk('u', 100, 'UNDERLINE'), mk('k', 200, 'STRIKETHROUGH'), mk('n', 300, undefined)],
    };
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    const out = path.join(os.tmpdir(), `deco-${Date.now()}.pptx`);
    await buildPptx([s], {}, cfg, out, { title: 'deco' });

    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'd-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const xml = fs.readFileSync(path.join(d, 'ppt', 'slides', 'slide1.xml'), 'utf8');
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });

    assert.strictEqual((xml.match(/\bu="sng"/g) || []).length, 1, '밑줄이 1개여야 합니다');
    assert.strictEqual((xml.match(/strike="sngStrike"/g) || []).length, 1, '취소선이 1개여야 합니다');
    console.log('  ✅ 밑줄·취소선을 살린다');
    return true;
  } catch (e) {
    console.log(`  ❌ 밑줄·취소선을 살린다\n     ${e.message}`);
    return false;
  }
}

// 15) Figma keeps list formatting in per-paragraph arrays next to the runs, not
//     inside them. Ignoring those dropped every bullet and its indent, which the
//     element check surfaced as labels sitting ~17px left of the original.
async function bulletTest() {
  try {
    const os = require('os');
    const cp = require('child_process');
    const mkOp = (id, y, runs, lineTypes) => ({
      kind: 'text', id, x: 165, y, w: 400, h: 120, clip: null, runs,
      lineTypes, lineIndentations: lineTypes.map(() => 1),
      style: { fontSize: 28, textAutoResize: 'WIDTH_AND_HEIGHT', textAlignHorizontal: 'LEFT' },
      baseFill: { r: 0, g: 0, b: 0, a: 1 }, bg: { r: 255, g: 255, b: 255, a: 1 }, opacity: 1,
    });
    const s = {
      name: 'B', id: 'b', w: 1920, h: 1080, bg: { r: 255, g: 255, b: 255, a: 1 },
      ops: [
        mkOp('t1', 501, [{ text: '프로세스 흐름', style: {}, breakLine: false }], ['UNORDERED']),
        mkOp('t2', 601, [{ text: '첫째', style: {}, breakLine: true },
          { text: '둘째', style: {}, breakLine: false }], ['ORDERED', 'ORDERED']),
        mkOp('t3', 801, [{ text: '평범한 문장', style: {}, breakLine: false }], []),
      ],
    };
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    const out = path.join(os.tmpdir(), `bullet-${Date.now()}.pptx`);
    await buildPptx([s], {}, cfg, out, { title: 'bullet' });

    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'b-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const xml = fs.readFileSync(path.join(d, 'ppt', 'slides', 'slide1.xml'), 'utf8');
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });

    assert.strictEqual((xml.match(/<a:buChar/g) || []).length, 1, '글머리 기호가 1개여야 합니다');
    assert.strictEqual((xml.match(/<a:buAutoNum/g) || []).length, 2, '번호 문단이 2개여야 합니다');
    console.log('  ✅ Figma 글머리 기호·번호 목록을 살린다');
    return true;
  } catch (e) {
    console.log(`  ❌ Figma 글머리 기호·번호 목록을 살린다\n     ${e.message}`);
    return false;
  }
}

// 20) REGULAR_POLYGON fell through to a plain rectangle, a rotated picture was
//     stretched into its bounding box, and a line with an arrowhead has a
//     zero-height box, so the render disappeared into it. Figma exports the render
//     bounds before any frame clipping, so the picture must be placed at that size
//     and cropped, never squeezed into the clipped box.
async function shapeImageTest() {
  try {
    const os = require('os'), cp = require('child_process');
    const rb = (x, y, w, h) => ({ x: OX + x, y: OY + y, width: w, height: h });
    const frame = {
      id: 'f', name: 'f', type: 'FRAME', absoluteBoundingBox: bb(0, 0, 1000, 500), fills: solid(0, 0, 0),
      clipsContent: true,
      children: [
        { id: 'tri', name: 't', type: 'REGULAR_POLYGON', absoluteBoundingBox: bb(100, 100, 49, 43),
          absoluteRenderBounds: rb(106, 103, 37, 37), fills: solid(255, 255, 255) },
        { id: 'rot', name: 'r', type: 'RECTANGLE', rotation: 0.5, absoluteBoundingBox: bb(200, 100, 80, 80),
          absoluteRenderBounds: rb(200, 100, 80, 80), fills: solid(255, 0, 0) },
        { id: 'line', name: 'l', type: 'VECTOR', absoluteBoundingBox: bb(300, 200, 84, 0),
          absoluteRenderBounds: rb(300, 186, 84, 28), strokes: solid(255, 255, 255), strokeWeight: 5 },
        { id: 'pic', name: 'p', type: 'RECTANGLE', rotation: -0.7, absoluteBoundingBox: bb(-31, 100, 432, 440),
          absoluteRenderBounds: rb(0, 95, 406, 405), fills: [{ type: 'IMAGE', imageRef: 'x' }] },
      ],
    };
    const rasterIds = new Set();
    const s = normalizeFrame(frame, { rasterIds, imageRefs: new Set(), notes: [] });
    assert.ok(rasterIds.has('tri'), '정다각형이 이미지로 렌더되지 않았습니다');
    assert.ok(rasterIds.has('rot'), '기울어진 단색 도형이 이미지로 렌더되지 않았습니다');
    assert.ok(!s.ops.some((o) => o.kind === 'rect'), '다각형·기울어진 도형이 사각형으로 나갔습니다');

    const png = (w, h) => {
      const { PNG } = require('pngjs');
      const b = PNG.sync.write(new PNG({ width: w, height: h }));
      return { data: b.toString('base64'), mime: 'image/png', width: w, height: h };
    };
    const assets = { tri: png(37, 37), rot: png(80, 80), line: png(84, 28), pic: png(443, 450) };
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    cfg.raster.scale = 1;
    const out = path.join(os.tmpdir(), `shape-${Date.now()}.pptx`);
    await buildPptx([s], assets, cfg, out, { title: 'shape' });
    const d = fs.mkdtempSync(path.join(os.tmpdir(), 's-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const xml = fs.readFileSync(path.join(d, 'ppt', 'slides', 'slide1.xml'), 'utf8');
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });

    const pics = xml.split('<p:pic>').slice(1).map((c) => ({
      off: c.match(/<a:off x="(-?\d+)" y="(-?\d+)"/).slice(1).map(Number),
      ext: c.match(/<a:ext cx="(\d+)" cy="(\d+)"/).slice(1).map(Number),
      l: Number((c.match(/srcRect l="(-?\d+)"/) || [])[1] || 0),
    }));
    const unit = 13.3333 * 914400 / 1000;                 // EMU per design px
    const near = (a, b, tol = 2) => Math.abs(a / unit - b) <= tol;
    const line = pics[2];
    assert.ok(near(line.ext[1], 28), `화살표 선 높이 ${line.ext[1] / unit}`);
    assert.ok(near(line.off[1], 186), `화살표 선 위치 ${line.off[1] / unit}`);
    const pic = pics[3];
    assert.ok(near(pic.ext[0], 406), `잘린 그림 폭 ${pic.ext[0] / unit}`);
    assert.ok(near(pic.off[0], 0), `잘린 그림 위치 ${pic.off[0] / unit}`);
    assert.ok(Math.abs(pic.l - 8350) < 150, `왼쪽 자르기 ${pic.l}`);        // 37/443 of the picture
    console.log('  ✅ 다각형·기울어진 그림·화살표 선을 원래 크기로 놓고 자른다');
    return true;
  } catch (e) {
    console.log(`  ❌ 다각형·기울어진 그림·화살표 선을 원래 크기로 놓고 자른다\n     ${e.message}`);
    return false;
  }
}

// 21) One box with a small line above a big one got the big line's height on
//     every line, which pushed the whole block down by the difference.
async function mixedLineHeightTest() {
  try {
    const os = require('os'), cp = require('child_process');
    const node = text('t', 10, 10, 600, 169, '작은 줄\n큰 줄', {
      style: { fontSize: 80, lineHeightPx: 96, lineHeightUnit: 'FONT_SIZE_%', lineHeightPercentFontSize: 120 },
      characterStyleOverrides: [7, 7, 7, 7, 0, 0, 0, 0],
      styleOverrideTable: { 7: { fontSize: 56 } },
    });
    const frame = { id: 'f', name: 'f', type: 'FRAME', absoluteBoundingBox: bb(0, 0, 1000, 500), fills: solid(0, 0, 0), children: [node] };
    const s = normalizeFrame(frame, { rasterIds: new Set(), imageRefs: new Set(), notes: [] });
    const cfg = loadConfig();
    cfg.__resolver = makeResolver(cfg);
    cfg.text.useExactLineSpacing = true;
    const out = path.join(os.tmpdir(), `lh-${Date.now()}.pptx`);
    const r = await buildPptx([s], {}, cfg, out, { title: 'lh' });
    const d = fs.mkdtempSync(path.join(os.tmpdir(), 'l-'));
    cp.execFileSync('unzip', ['-q', '-o', out, '-d', d]);
    const xml = fs.readFileSync(path.join(d, 'ppt', 'slides', 'slide1.xml'), 'utf8');
    fs.rmSync(d, { recursive: true, force: true });
    fs.rmSync(out, { force: true });
    const pts = [...xml.matchAll(/<a:spcPts val="(\d+)"/g)].map((m) => Number(m[1]));
    const px2pt = 960 / 1000;
    assert.strictEqual(pts.length, 2, `줄 간격 ${pts.length}개`);
    assert.ok(Math.abs(pts[0] - 67.2 * px2pt * 100) < 2, `작은 줄 ${pts[0]}`);
    assert.ok(Math.abs(pts[1] - 96 * px2pt * 100) < 2, `큰 줄 ${pts[1]}`);
    console.log('  ✅ 크기가 다른 줄마다 자기 줄 높이를 쓴다');
    return true;
  } catch (e) {
    console.log(`  ❌ 크기가 다른 줄마다 자기 줄 높이를 쓴다\n     ${e.message}`);
    return false;
  }
}

// 22) Three ways one box with mixed runs went wrong: a small run inherited the big
//     run's tracking in px and collapsed, "\n\n" lost its empty line, and a gradient
//     run had no solid colour so it came out plain white.
function mixedRunTest() {
  try {
    const grad = [{ type: 'GRADIENT_LINEAR',
      gradientHandlePositions: [{ x: 0, y: 0.5 }, { x: 1, y: 0.5 }, { x: 0, y: 1 }],
      gradientStops: [{ position: 0, color: { r: 1, g: 0, b: 0, a: 1 } }, { position: 1, color: { r: 0, g: 0, b: 1, a: 1 } }] }];
    const chars = 'AB\n\nCD';
    const node = text('t', 10, 10, 100, 200, chars, {
      style: { fontSize: 40, letterSpacing: -2.4, textAlignHorizontal: 'LEFT' },
      characterStyleOverrides: [7, 7, 0, 0, 9, 9],
      styleOverrideTable: { 7: { fontSize: 18 }, 9: { fills: grad } },
    });
    const frame = { id: 'f', name: 'f', type: 'FRAME', absoluteBoundingBox: bb(0, 0, 1000, 500), fills: solid(0, 0, 0), children: [node] };
    const op = normalizeFrame(frame, { rasterIds: new Set(), imageRefs: new Set(), notes: [] }).ops.find((o) => o.kind === 'text');

    assert.ok(Math.abs(op.runs[0].style.letterSpacing + 1.08) < 0.01, `자간 ${op.runs[0].style.letterSpacing}`);
    assert.strictEqual(op.runs.filter((r) => r.breakLine).length, 2, '빈 줄이 사라졌습니다');
    const tail = op.runs.slice(-2);
    assert.deepStrictEqual(tail.map((r) => r.text), ['C', 'D'], '그라디언트 런이 글자별로 나뉘지 않았습니다');
    const red = (r) => r.style.fills[0].color.r;
    assert.ok(tail.every((r) => r.style.fills[0].type === 'SOLID'), '그라디언트가 단색으로 바뀌지 않았습니다');
    assert.ok(red(tail[0]) > red(tail[1]), '글자 위치에 따라 색이 바뀌지 않았습니다');
    console.log('  ✅ 섞인 글자의 자간·빈 줄·그라디언트를 살린다');
    return true;
  } catch (e) {
    console.log(`  ❌ 섞인 글자의 자간·빈 줄·그라디언트를 살린다\n     ${e.message}`);
    return false;
  }
}

// 23) The GUI is a local server, and any web page the user visits can send requests
//     to it. Without this check a page could swap the token, or start a conversion
//     writing to a path of its choosing.
function localGuardTest() {
  try {
    const { isLocalRequest } = require('../gui/guard');
    const ask = (headers) => isLocalRequest({ headers }, 4785);
    assert.ok(ask({ host: '127.0.0.1:4785' }), '주소창에서 연 요청이 막혔습니다');
    assert.ok(ask({ host: '[::1]:4785' }), 'IPv6 요청이 막혔습니다');
    assert.ok(ask({ host: '127.0.0.1:4785', origin: 'http://127.0.0.1:4785' }), '자기 화면의 요청이 막혔습니다');
    assert.ok(!ask({ host: '127.0.0.1:4785', origin: 'https://evil.example' }), '다른 사이트 요청이 통과했습니다');
    assert.ok(!ask({ host: 'evil.example:4785' }), 'DNS 리바인딩 요청이 통과했습니다');
    assert.ok(!ask({}), 'Host 없는 요청이 통과했습니다');
    // <img src=...> sends no Origin at all, only the fetch metadata
    assert.ok(!ask({ host: '127.0.0.1:4785', 'sec-fetch-site': 'cross-site' }), '이미지 태그 요청이 통과했습니다');
    assert.ok(!ask({ host: '127.0.0.1:4785', 'sec-fetch-site': 'same-site' }), '같은 사이트 요청이 통과했습니다');
    assert.ok(ask({ host: '127.0.0.1:4785', 'sec-fetch-site': 'same-origin' }), '자기 화면 요청이 막혔습니다');
    assert.ok(ask({ host: '127.0.0.1:4785', 'sec-fetch-site': 'none' }), '주소창 요청이 막혔습니다');
    console.log('  ✅ 다른 사이트에서 온 요청을 막는다');
    return true;
  } catch (e) {
    console.log(`  ❌ 다른 사이트에서 온 요청을 막는다\n     ${e.message}`);
    return false;
  }
}

// 24) The output folder is named after the Figma section, which anyone can name.
function safeNameTest() {
  try {
    const { safeName } = require('../src/pipeline');
    assert.strictEqual(safeName('Section 1'), 'Section 1');
    assert.strictEqual(safeName('a/b:c'), 'a_b_c');
    assert.ok(!safeName('..').startsWith('.'), `이름 ${safeName('..')}`);
    assert.ok(!safeName('../../x').includes('..'), `이름 ${safeName('../../x')}`);
    assert.strictEqual(safeName('   '), 'figma-deck');
    console.log('  ✅ 섹션 이름으로 폴더 밖에 쓰지 못한다');
    return true;
  } catch (e) {
    console.log(`  ❌ 섹션 이름으로 폴더 밖에 쓰지 못한다\n     ${e.message}`);
    return false;
  }
}

if (process.argv.includes('--pptx')) {
  const cfg = loadConfig();
  cfg.__resolver = makeResolver(cfg);
  const out = path.join(__dirname, 'selftest.pptx');
  buildPptx([slide], {}, cfg, out, { title: 'figma2pptx self test' }).then((r) => {
    console.log(`\nPPTX 저장: ${r.path}`);
    if (r.warnings.length) console.log('경고:', r.warnings);
  });
}
