'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const figma = require('./figma');
const { normalizeFrame } = require('./normalize');
const { buildPptx } = require('./pptx');
const { makeResolver } = require('./fonts');
const { verify } = require('./verify');
const { createMeasurer } = require('./metrics');
const { dedupeMedia } = require('./dedupe');
const { encode } = require('./image');
const cache = require('./cache');

const DEFAULTS = require('../config/default.json');

function loadConfig(overridePath) {
  const cfg = JSON.parse(JSON.stringify(DEFAULTS));
  if (overridePath && fs.existsSync(overridePath)) {
    const user = JSON.parse(fs.readFileSync(overridePath, 'utf8'));
    deepMerge(cfg, user);
  }
  return cfg;
}

function deepMerge(a, b) {
  for (const [k, v] of Object.entries(b)) {
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      a[k] = a[k] && typeof a[k] === 'object' ? a[k] : {};
      deepMerge(a[k], v);
    } else a[k] = v;
  }
  return a;
}

/** Turn per-element ink measurements into the part of the report you act on.
 *  The page percentage tells you a page is off; this tells you which box, which
 *  direction, and by how much — in the pixels the design was drawn in. */
function elementSection(pages) {
  const rows = [];
  for (const p of pages) for (const e of p.elements || []) rows.push({ page: p.index, ...e });
  const L = ['## 요소별 위치·폭', ''];
  if (!rows.length) return [...L, '비교할 텍스트를 찾지 못했습니다.', ''];

  const measured = rows.filter((r) => r.kind === 'measured');
  const gone = rows.filter((r) => r.kind !== 'measured');

  // A whole deck rendering a few percent narrow is one fact about the font, not
  // hundreds of faults. State it once, then judge each element against it — what
  // is left over is the part that is actually wrong with this page.
  const ratios = measured.map((r) => r.ratio).sort((a, b) => a - b);
  const trend = ratios.length ? ratios[Math.floor(ratios.length / 2)] : 0;

  L.push('텍스트 상자마다 원본 렌더와 변환 결과에서 글자가 실제로 차지한 범위를 재서 비교했습니다.', '');
  L.push(`- 비교한 텍스트 ${measured.length}개`);
  if (Math.abs(trend) >= 0.005) {
    L.push(`- **전체 경향: PowerPoint가 ${(Math.abs(trend) * 100).toFixed(1)}% ${trend < 0 ? '좁게' : '넓게'} 조판합니다** `
      + '(서체 조판 차이 — 개별 오류가 아닙니다)');
  }

  const moved = measured
    .filter((r) => Math.abs(r.moveX) > r.tolPx || Math.abs(r.dy) > r.tolPx)
    .sort((a, b) => b.severity - a.severity);
  const odd = measured
    .filter((r) => Math.abs(r.ratio - trend) > Math.max(r.tolRatio, 0.03) && r.aw > 20)
    .sort((a, b) => Math.abs(b.ratio - trend) - Math.abs(a.ratio - trend));

  L.push(`- 위치가 밀린 것 ${moved.length}개 · 전체 경향에서 벗어난 폭 ${odd.length}개`
    + (gone.length ? ` · 사라진 것 ${gone.length}개` : ''), '');

  if (gone.length) {
    L.push('**빠졌거나 더 생긴 것 (가장 심각)**', '');
    gone.slice(0, 8).forEach((r) => L.push(`- ${r.kind === 'missing' ? '❌ 안 나옴' : '❓ 원본에 없음'}  p${r.page}  «${r.label}»`));
    L.push('');
  }
  if (moved.length) {
    L.push('**위치가 밀린 것 — 대개 변환 오류입니다**', '');
    moved.slice(0, 8).forEach((r) => {
      const bits = [];
      if (Math.abs(r.moveX) >= 1) bits.push(`가로 ${r.moveX > 0 ? '+' : ''}${r.moveX.toFixed(1)}px`);
      if (Math.abs(r.dy) >= 1) bits.push(`세로 ${r.dy > 0 ? '+' : ''}${r.dy.toFixed(1)}px`);
      L.push(`- p${r.page}  «${r.label}»  ${bits.join(' · ')}`);
    });
    L.push('');
  }
  if (odd.length) {
    L.push('**전체 경향에서 벗어난 폭 — 이 요소만 다른 이유가 있습니다**', '');
    odd.slice(0, 8).forEach((r) => {
      L.push(`- p${r.page}  «${r.label}»  폭 ${(r.ratio * 100).toFixed(1)}% `
        + `(경향 ${(trend * 100).toFixed(1)}% 대비 ${((r.ratio - trend) * 100).toFixed(1)}%p)`);
    });
    L.push('');
  }
  if (!gone.length && !moved.length && !odd.length) {
    L.push('전체 경향을 벗어나거나 위치가 밀린 텍스트가 없습니다.', '');
  }
  return L;
}

/** One home for everything this tool makes, so it never scatters across the Desktop. */
function defaultOutRoot() {
  return path.join(require('os').homedir(), 'Desktop', 'figma2pptx');
}

/** Collect the frames to turn into slides, in reading order. */
function collectFrames(rootNode) {
  const isFrame = (n) => n.type === 'FRAME' || n.type === 'COMPONENT';
  let frames;
  if (isFrame(rootNode)) frames = [rootNode];
  else frames = (rootNode.children || []).filter(isFrame);

  // sort into rows, then left-to-right — matches how decks are laid out on a canvas
  return frames.slice().sort((a, b) => {
    const A = a.absoluteBoundingBox || { x: 0, y: 0 };
    const B = b.absoluteBoundingBox || { x: 0, y: 0 };
    if (Math.abs(A.y - B.y) > 200) return A.y - B.y;
    return A.x - B.x;
  });
}

async function run(opts) {
  const log = opts.log || (() => {});
  const cfg = loadConfig(opts.configPath);
  if (opts.noVerify) cfg.verify.enabled = false;

  const token = opts.token || process.env.FIGMA_TOKEN;
  if (!token) throw new Error('FIGMA_TOKEN이 없습니다. .env 파일 또는 --token 옵션으로 전달하세요.');

  // a long wait with no explanation looks like a hang; say why we are waiting
  figma.setRetryLogger(log);

  const { fileKey, nodeId } = figma.parseTarget(opts.target);
  log(`파일 ${fileKey}${nodeId ? ` / 노드 ${nodeId}` : ''}`);

  // --- fetch ---------------------------------------------------------------
  log('Figma에서 문서 구조를 가져오는 중…');
  let rootNode, version = null;
  if (nodeId) {
    const nodes = await figma.getNodes(fileKey, [nodeId], token);
    const entry = nodes[nodeId];
    if (!entry || !entry.document) {
      throw new Error(`노드를 찾지 못했습니다: ${nodeId}\n`
        + '방금 만들거나 옮긴 섹션이라면 Figma가 아직 저장을 끝내지 못했을 수 있습니다. '
        + '잠시 후 링크를 다시 복사해 주세요.');
    }
    rootNode = entry.document;
    version = nodes.__version;
  } else {
    const file = await figma.getFile(fileKey, token);
    rootNode = (file.document.children || [])[0];
    if (!rootNode) throw new Error('문서에 페이지가 없습니다.');
    version = file.version;
  }

  const frames = collectFrames(rootNode);
  if (!frames.length) throw new Error('변환할 프레임을 찾지 못했습니다. 섹션 또는 프레임을 지정해 주세요.');
  log(`프레임 ${frames.length}개를 찾았습니다.`);

  // --- normalize -----------------------------------------------------------
  const rasterIds = new Set();
  const imageRefs = new Set();
  const notes = [];
  const slides = frames.map((f) => normalizeFrame(f, { rasterIds, imageRefs, notes }));

  // A slide with nothing on it is never what the designer drew. This shipped
  // once — four blank pages whose Figma originals were full-bleed artwork — and
  // nothing in the report said so, because a blank page against a mostly-white
  // reference still scored as "close enough".
  const preWarnings = [];
  slides.forEach((s, i) => {
    if (!s.ops.length && !s.blank) {
      preWarnings.push(`⚠️ p${i + 1} «${s.name}» 이 비어 있습니다 — 원본에는 내용이 있는데 변환되지 않았을 수 있습니다.`);
    }
  });

  const sizes = new Set(slides.map((s) => `${Math.round(s.w)}x${Math.round(s.h)}`));
  if (sizes.size > 1) {
    notes.push(`프레임 크기가 서로 다릅니다(${[...sizes].join(', ')}). 첫 프레임 기준으로 슬라이드를 맞췄습니다.`);
  }

  // --- rasterize what cannot be native -------------------------------------
  const assets = {};
  const store = cache.open({ fileKey, version, enabled: cfg.raster.cache !== false && !!version });
  if (rasterIds.size) {
    // "why is it slow again when it was instant yesterday" has exactly one
    // answer, and it is worth saying out loud rather than letting the user guess.
    const held = cache.versionsOf(fileKey);
    if (held.length && !held.some((h) => h.v === String(version))) {
      log('Figma 파일이 수정되어 이전 캐시를 쓸 수 없습니다. 이번엔 다시 받습니다.');
    }
    const dropped = cache.prune(fileKey, version);
    if (dropped) log(`오래된 캐시 ${dropped}개를 정리했습니다.`);
    const ids = [...rasterIds];
    const missing = ids.filter((id) => !store.has(id, cfg.raster.scale));
    if (missing.length < ids.length) {
      log(`이미지 ${ids.length}개 중 ${ids.length - missing.length}개는 캐시에 있습니다.`);
    }

    // Fetch render URLs and download in the SAME loop, a chunk at a time.
    //
    // Asking for all 2605 URLs first and only then downloading meant the run had
    // to survive 66 consecutive rate-limited calls before a single byte reached
    // the cache. On a section that large it never got there: endless 429s, and
    // nothing saved to resume from. Chunking makes every completed chunk
    // permanent, so even a run that dies partway leaves the next one less to do.
    const urls = {};
    let fetched = 0;
    if (missing.length) {
      log(`남은 ${missing.length}개를 Figma에서 렌더링하는 중… (가장 오래 걸리는 단계입니다)`);
    } else {
      log('모든 이미지가 캐시에 있어 Figma 렌더링을 건너뜁니다.');
    }
    // Repeated artwork (the same cover shown on several pages) renders to
    // byte-identical PNGs. Downloading each one again and embedding it again is
    // how an 11-page deck reached 75MB with only 88 distinct images in it.
    // Figma hands out a distinct URL per node, so identical artwork can only be
    // spotted after it arrives — hash the bytes and let every node that renders
    // to the same picture share one string, which keeps the packer from writing
    // the same file over and over.
    const byHash = new Map();
    let reused = 0, saved = 0, toJpeg = 0, done = 0;

    // One at a time is fine for a deck with a dozen icons and hopeless for one
    // with thousands, where each round trip's latency lands end to end. Keep a
    // small pool: enough to hide the waiting, not so many that Figma refuses.
    const lanes = Math.max(1, Math.min(cfg.raster.concurrency || 5, 16));
    const take = async (id) => {
      try {
        let buf = store.read(id, cfg.raster.scale);
        if (!buf) {
          const u = urls[id];
          if (!u) { notes.push(`이미지 렌더 URL 없음: ${id}`); return; }
          buf = await figma.download(u);
          store.write(id, cfg.raster.scale, buf);
        }
        const h = crypto.createHash('sha1').update(buf).digest('hex');
        if (byHash.has(h)) { assets[id] = byHash.get(h); reused++; }
        else {
          const enc = encode(buf, cfg.raster);
          if (enc.saved) { saved += enc.saved; toJpeg++; }
          byHash.set(h, enc);
          assets[id] = enc;
        }
      } catch (e) {
        notes.push(`이미지 다운로드 실패 ${id}: ${e.message}`);
      }
      done++;
    };
    const drain = async (list) => {
      let cursor = 0;
      const worker = async () => { while (cursor < list.length) await take(list[cursor++]); };
      await Promise.all(Array.from({ length: lanes }, worker));
    };

    // whatever is already on disk costs nothing and needs no permission
    await drain(ids.filter((id) => store.has(id, cfg.raster.scale)));

    const chunk = Math.max(5, Math.min(cfg.raster.renderBatch || 30, 100));
    for (let i = 0; i < missing.length; i += chunk) {
      const part = missing.slice(i, i + chunk);
      let got;
      try {
        // A long run crosses network blips. Giving up after ~30s of backoff
        // abandons everything still unfetched, so be patient here specifically.
        got = await figma.getImageUrls(fileKey, part, token,
          { format: 'png', scale: cfg.raster.scale, tries: 8 });
      } catch (e) {
        // Everything fetched so far is already cached, so stopping here is a
        // pause rather than a loss — say so instead of throwing the run away.
        log(`⚠️ Figma 렌더 요청이 막혔습니다 (${e.message.split('\n')[0]})`);
        log(`   여기까지 ${fetched}개를 캐시에 저장했습니다. 다시 실행하면 이어서 받습니다.`);
        notes.push(`이미지 ${missing.length - fetched}개를 받지 못했습니다. 다시 실행하면 이어집니다.`);
        break;
      }
      Object.assign(urls, got);
      await drain(part);
      fetched += part.length;
      log(`  이미지 ${fetched}/${missing.length}개 받음 (캐시에 저장됨)`);
    }

    if (store.hits) log(`캐시에서 ${store.hits}개를 바로 읽었습니다 (다시 받지 않음).`);
    if (reused) log(`같은 그림 ${reused}개는 한 번만 담았습니다 (고유 ${byHash.size}개).`);
    if (toJpeg) log(`불투명한 사진 ${toJpeg}개를 JPEG로 담아 ${(saved / 1048576).toFixed(1)}MB 줄였습니다.`);
  }

  // Frame backgrounds that sit behind editable children: fetch the original
  // upload rather than a render of the frame, which would flatten the children.
  if (imageRefs.size) {
    log(`프레임 배경 이미지 ${imageRefs.size}개를 가져오는 중…`);
    try {
      const fills = await figma.getImageFills(fileKey, token);
      for (const ref of imageRefs) {
        const u = fills && fills[ref];
        if (!u) { notes.push(`프레임 배경 이미지를 찾지 못했습니다: ${ref}`); continue; }
        try {
          let buf = store.read(`ref:${ref}`, 0);
          if (!buf) { buf = await figma.download(u); store.write(`ref:${ref}`, 0, buf); }
          assets[`ref:${ref}`] = encode(buf, cfg.raster);
        } catch (e) {
          notes.push(`프레임 배경 이미지 다운로드 실패: ${e.message}`);
        }
      }
    } catch (e) {
      notes.push(`프레임 배경 이미지 목록을 가져오지 못했습니다: ${e.message}`);
    }
  }

  // --- build ---------------------------------------------------------------
  const resolver = makeResolver(cfg);
  cfg.__resolver = resolver;
  cfg.__measurer = createMeasurer(log);

  // Everything a conversion produces goes in one folder named after the deck:
  // the pptx, its report, and the verify images. Loose files at the top level of
  // the Desktop is how the last run ended up with a stale "verify/" that no
  // longer matched the report sitting next to it.
  const baseName = safeName(opts.name || rootNode.name);
  const outDir = path.join(opts.outDir || defaultOutRoot(), baseName);
  fs.mkdirSync(outDir, { recursive: true });
  const pptxPath = path.join(outDir, `${baseName}.pptx`);

  // An interrupted run must never quietly replace a better one. A partial deck
  // (930 of 2605 images) overwrote a nearly complete one (2430) and the only
  // sign was the file shrinking from 80MB to 3MB. Keep the previous result.
  const missingImages = rasterIds.size - Object.keys(assets).length;
  if (missingImages > 0 && fs.existsSync(pptxPath)) {
    const kept = path.join(outDir, `${baseName} (직전).pptx`);
    try {
      fs.renameSync(pptxPath, kept);
      log(`이번 변환은 이미지 ${missingImages}개가 빠져 있어, 직전 결과를 «${path.basename(kept)}»로 남겨 두었습니다.`);
    } catch { /* if it cannot be moved it will simply be overwritten */ }
  }

  log('PPTX를 만드는 중…');
  const built = await buildPptx(slides, assets, cfg, pptxPath, { title: baseName });

  const dd = await dedupeMedia(built.path).catch((e) => {
    notes.push(`이미지 중복 정리를 건너뛰었습니다: ${e.message}`);
    return null;
  });
  if (dd && dd.removed) {
    const mb = (n) => (n / 1048576).toFixed(1);
    log(`같은 그림 ${dd.removed}개를 합쳤습니다 — ${mb(dd.before)}MB → ${mb(dd.after)}MB`);
  }
  log(`저장 완료: ${built.path}`);

  const report = {
    file: built.path,
    slides: slides.length,
    texts: slides.reduce((n, s) => n + s.ops.filter((o) => o.kind === 'text').length, 0),
    shapes: slides.reduce((n, s) => n + s.ops.filter((o) => o.kind !== 'text' && o.kind !== 'image').length, 0),
    images: Object.keys(assets).length,
    missingFonts: [...resolver.missing],
    notes,
    warnings: [...preWarnings, ...built.warnings],
    drift: built.drift || [],
    verify: null,
  };

  // --- verify --------------------------------------------------------------
  if (cfg.verify.enabled) {
    try {
      report.verify = await verify({ fileKey, token, slides, pptxPath: built.path, outDir, baseName, cfg, log });
    } catch (e) {
      report.verify = { ran: false, reason: e.message, pages: [] };
    }
  }

  writeReport(report, outDir, baseName);
  return report;
}

function writeReport(r, outDir, baseName) {
  const L = [];
  L.push(`# 변환 리포트 — ${baseName}`, '');
  L.push(`- 결과 파일: \`${r.file}\``);
  L.push(`- 슬라이드 ${r.slides}장 · 텍스트 ${r.texts}개 · 도형/선 ${r.shapes}개 · 이미지 ${r.images}개`);
  L.push('');

  if (r.missingFonts.length) {
    L.push('## 설치되지 않은 서체', '');
    L.push('아래 서체가 이 Mac에 없어 PowerPoint가 다른 서체로 대체합니다. 설치 후 다시 열면 정상 표시됩니다.', '');
    r.missingFonts.forEach((f) => L.push(`- ${f}`));
    L.push('');
  }

  if (r.drift.length) {
    const worst = r.drift.slice().sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
    const wide = worst.filter((d) => d.delta > 0).length;
    L.push('## 서체 폭 차이', '');
    L.push('디자인이 측정한 글자 폭과 PowerPoint가 실제로 조판하는 폭이 다릅니다.');
    L.push('글자에 딱 맞춰 그린 요소(칩·박스·구분선)나 가운데/오른쪽 정렬이 어긋나 보일 수 있습니다.', '');
    L.push(`- 차이가 난 텍스트 ${r.drift.length}개 (더 넓게 조판 ${wide}개)`);
    L.push(`- 최대 차이 ${worst[0].delta.toFixed(1)}px`, '');
    L.push('가장 큰 것들:', '');
    for (const d of worst.slice(0, 8))
      L.push(`  - ${d.delta > 0 ? '+' : ''}${d.delta.toFixed(1)}px  ${d.page} «${d.text}»`);
    L.push('');
  }

  if (r.verify && r.verify.ran) {
    const bad = r.verify.pages.filter((p) => p.status === 'check');
    L.push('## 렌더 비교', '');
    L.push(`엔진: ${r.verify.engine} · 확인 필요 ${bad.length} / ${r.verify.pages.length}장`, '');
    for (const p of r.verify.pages) {
      const d = p.diff === null ? '—' : (p.diff * 100).toFixed(2) + '%';
      L.push(`- ${p.status === 'ok' ? '✅' : p.status === 'skip' ? '⏭️' : '⚠️'} p${p.index} ${p.name} — 차이 ${d}`);
    }
    // only point at the folder when tidy() actually left one behind
    if (bad.length && r.verify.dir && fs.existsSync(r.verify.dir)) {
      L.push('', `⚠️ 표시된 페이지의 이미지는 \`${path.basename(r.verify.dir)}/\` 폴더에 있습니다.`,
        '`figma-*` 원본 · `ppt-*` 변환 결과 · `diff-*` 다른 부분(빨간색)', '');
    } else {
      L.push('', '모든 페이지가 기준을 통과해 검수 이미지는 남기지 않았습니다.', '');
    }
    L.push('> 두 프로그램이 글자 외곽선을 다르게 그리므로 0~3%는 정상입니다.', '');
    L.push(...elementSection(r.verify.pages));
  } else if (r.verify) {
    L.push('## 렌더 비교', '', `건너뜀: ${r.verify.reason}`, '');
  }

  if (r.notes.length) {
    L.push('## 참고', '');
    [...new Set(r.notes)].forEach((n) => L.push(`- ${n}`));
    L.push('');
  }
  if (r.warnings.length) {
    L.push('## 경고', '');
    [...new Set(r.warnings)].forEach((n) => L.push(`- ${n}`));
    L.push('');
  }

  const p = path.join(outDir, `${baseName}.report.md`);
  fs.writeFileSync(p, L.join('\n'), 'utf8');
  r.reportPath = p;
}

/** The output folder is named after a Figma section, and a section can be called anything. */
function safeName(name) {
  const clean = String(name || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\.{2,}/g, '_').replace(/^\.+/, '').trim();
  return clean || 'figma-deck';
}

module.exports = { run, loadConfig, safeName };
