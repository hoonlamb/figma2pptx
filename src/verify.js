'use strict';
// Renders the produced .pptx through PowerPoint (the renderer that actually matters)
// and compares each page against Figma's own render of the same frame.
// The point is to catch silent breakage — dropped text, re-wrapped lines, shifted
// tables — instead of trusting the conversion.

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');
const { PNG } = require('pngjs');
const pixelmatch = require('pixelmatch');
const figma = require('./figma');
const { compareText } = require('./inkdiff');

function which(bin) {
  try {
    return execFileSync('/usr/bin/which', [bin], { encoding: 'utf8' }).trim() || null;
  } catch {
    return null;
  }
}

function devToolsReady() {
  try {
    execFileSync('/usr/bin/xcode-select', ['-p'], { stdio: 'pipe' });
    return true;
  } catch {
    return false;
  }
}

function pptxToPdf(pptxPath, pdfPath, log) {
  if (process.platform === 'darwin') {
    // Touch only our own file. The first version bracketed this with
    // "close every presentation saving no", which throws away unsaved changes in
    // whatever the user happens to have open — a verification step must never be
    // able to destroy the work it is verifying alongside.
    //
    // PowerPoint is sandboxed: writing a PDF next to a deck anywhere else pops a
    // "Grant File Access" dialog nobody is watching, and the script hangs until the
    // AppleEvent times out. Its own container needs no grant, so render there.
    const box = path.join(os.homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/tmp', `figma2pptx-${process.pid}`);
    const boxPptx = path.join(box, 'render.pptx');
    const boxPdf = path.join(box, 'render.pdf');
    const script = `
tell application "Microsoft PowerPoint"
  open POSIX file "${boxPptx}"
  set deck to active presentation
  save deck in POSIX file "${boxPdf}" as save as PDF
  close deck saving no
end tell`;
    try {
      fs.mkdirSync(box, { recursive: true });
      fs.copyFileSync(pptxPath, boxPptx);
      execFileSync('/usr/bin/osascript', ['-e', script], { timeout: 240000, stdio: 'pipe' });
      if (fs.existsSync(boxPdf)) {
        fs.copyFileSync(boxPdf, pdfPath);
        return 'powerpoint';
      }
      log('PowerPoint가 PDF를 만들지 못했습니다 (파일이 생성되지 않음).');
    } catch (e) {
      // stderr carries the actual AppleScript complaint; e.message alone is
      // "Command failed: /usr/bin/osascript -e" and says nothing.
      const detail = String((e.stderr && e.stderr.toString()) || e.message || e).trim();
      log(`PowerPoint 렌더 실패: ${detail.split('\n').filter(Boolean).slice(0, 2).join(' / ')}`);
    } finally {
      fs.rmSync(box, { recursive: true, force: true });
    }
  }
  const soffice = which('soffice') || which('libreoffice');
  if (soffice) {
    execFileSync(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', path.dirname(pdfPath), pptxPath], {
      timeout: 240000,
      stdio: 'pipe',
    });
    const produced = path.join(path.dirname(pdfPath), path.basename(pptxPath).replace(/\.pptx$/i, '.pdf'));
    if (produced !== pdfPath && fs.existsSync(produced)) fs.renameSync(produced, pdfPath);
    if (fs.existsSync(pdfPath)) return 'libreoffice';
  }
  return null;
}

function pdfToPngs(pdfPath, outDir, targetW, log) {
  const prefix = path.join(outDir, 'ppt');
  if (process.platform === 'darwin') {
    // /usr/bin/swift exists on every Mac as a stub that pops an installer dialog,
    // so finding it says nothing. xcode-select says whether it would really run.
    const swift = devToolsReady() ? which('swift') : null;
    if (swift) {
      try {
        execFileSync(swift, [path.join(__dirname, '..', 'tools', 'pdf2png.swift'), pdfPath, prefix, String(targetW)], {
          timeout: 300000,
          stdio: 'pipe',
        });
        const files = fs.readdirSync(outDir).filter((f) => f.startsWith('ppt-') && f.endsWith('.png')).sort();
        if (files.length) return files.map((f) => path.join(outDir, f));
      } catch (e) {
        log(`PDF 렌더 실패(swift): ${String(e.message || e).split('\n')[0]}`);
      }
    }
  }
  const pdftoppm = which('pdftoppm');
  if (pdftoppm) {
    execFileSync(pdftoppm, ['-png', '-scale-to-x', String(targetW), '-scale-to-y', '-1', pdfPath, prefix], {
      timeout: 300000,
      stdio: 'pipe',
    });
    const files = fs.readdirSync(outDir).filter((f) => f.startsWith('ppt-') && f.endsWith('.png')).sort();
    return files.map((f) => path.join(outDir, f));
  }
  return [];
}

function readPng(p) {
  return PNG.sync.read(fs.readFileSync(p));
}

/** Bilinear resample. pngjs has no scaler and the two renderers pick their own
 *  raster sizes, so we need one before any pixel can be compared. */
function resample(img, w, h) {
  if (img.width === w && img.height === h) return img;
  const out = new PNG({ width: w, height: h });
  const sx = img.width / w, sy = img.height / h;
  for (let y = 0; y < h; y++) {
    const fy = Math.min((y + 0.5) * sy - 0.5, img.height - 1);
    const y0 = Math.max(Math.floor(fy), 0), y1 = Math.min(y0 + 1, img.height - 1);
    const wy = fy - y0;
    for (let x = 0; x < w; x++) {
      const fx = Math.min((x + 0.5) * sx - 0.5, img.width - 1);
      const x0 = Math.max(Math.floor(fx), 0), x1 = Math.min(x0 + 1, img.width - 1);
      const wx = fx - x0;
      const i = (y * w + x) << 2;
      for (let c = 0; c < 4; c++) {
        const p00 = img.data[((y0 * img.width + x0) << 2) + c];
        const p10 = img.data[((y0 * img.width + x1) << 2) + c];
        const p01 = img.data[((y1 * img.width + x0) << 2) + c];
        const p11 = img.data[((y1 * img.width + x1) << 2) + c];
        out.data[i + c] =
          (p00 * (1 - wx) + p10 * wx) * (1 - wy) + (p01 * (1 - wx) + p11 * wx) * wy;
      }
    }
  }
  return out;
}

/** Bring both renders onto one canvas so pixelmatch compares like with like.
 *
 *  These two images are the SAME slide drawn by two programs that each chose
 *  their own raster size — PowerPoint's PDF came out at 3200x1800 while Figma
 *  rendered 1600x900. Padding the smaller one (the old behaviour) parks it in
 *  the top-left quarter and leaves the rest blank, so nearly every pixel counts
 *  as different and a perfect conversion reports a 60%+ mismatch. Scale first,
 *  and only pad whatever aspect-ratio remainder is left. */
function align(a, b) {
  const w = Math.min(a.width, b.width);
  const ha = Math.round(a.height * (w / a.width));
  const hb = Math.round(b.height * (w / b.width));
  const A = resample(a, w, ha);
  const B = resample(b, w, hb);
  const h = Math.max(ha, hb);
  const pad = (img) => {
    if (img.height === h) return img;
    const out = new PNG({ width: w, height: h });
    out.data.fill(255);
    PNG.bitblt(img, out, 0, 0, w, img.height, 0, 0);
    return out;
  };
  return [pad(A), pad(B), w, h];
}

async function verify({ fileKey, token, slides, pptxPath, outDir, baseName, cfg, log }) {
  // One folder per deck, named after it. A single shared "verify/" meant the next
  // conversion silently overwrote the last one's evidence — and worse, pdfToPngs
  // globs ppt-*.png out of this folder, so a 5-page deck run after an 11-page one
  // would pick up six stale pages and compare the wrong slides. Wiping a
  // deck-specific folder makes both impossible.
  const vDir = path.join(outDir, `${baseName} 검수`);
  fs.rmSync(vDir, { recursive: true, force: true });
  fs.mkdirSync(vDir, { recursive: true });

  const targetW = 1600;
  const frameW = slides[0].w;
  const scale = Math.max(0.01, Math.min(4, targetW / frameW));

  log('Figma 원본 렌더를 받는 중…');
  const urls = await figma.getImageUrls(fileKey, slides.map((s) => s.id), token, { format: 'png', scale });
  const refPaths = [];
  for (let i = 0; i < slides.length; i++) {
    const u = urls[slides[i].id];
    if (!u) { refPaths.push(null); continue; }
    const buf = await figma.download(u);
    const p = path.join(vDir, `figma-${String(i + 1).padStart(3, '0')}.png`);
    fs.writeFileSync(p, buf);
    refPaths.push(p);
  }

  log('PowerPoint로 PDF 렌더 중…');
  const pdfPath = path.join(vDir, 'render.pdf');
  const engine = pptxToPdf(pptxPath, pdfPath, log);
  if (!engine) {
    return {
      ran: false,
      reason: 'PowerPoint 렌더가 실패했거나 PowerPoint·LibreOffice가 없어 렌더 비교를 건너뛰었습니다.',
      pages: [],
    };
  }
  log(`PDF 렌더 완료 (${engine}) → 페이지 이미지 변환 중…`);
  const pptPngs = pdfToPngs(pdfPath, vDir, targetW, log);
  if (!pptPngs.length) {
    return { ran: false, reason: 'PDF를 이미지로 변환하지 못했습니다.', pages: [] };
  }

  const pages = [];
  for (let i = 0; i < slides.length; i++) {
    const ref = refPaths[i];
    const got = pptPngs[i];
    if (!ref || !got) {
      pages.push({ index: i + 1, name: slides[i].name, status: 'skip', diff: null });
      continue;
    }
    const [A, B, w, h] = align(readPng(ref), readPng(got));
    const diff = new PNG({ width: w, height: h });
    const mismatched = pixelmatch(A.data, B.data, diff.data, w, h, { threshold: 0.12, includeAA: false });
    const ratio = mismatched / (w * h);
    const diffPath = path.join(vDir, `diff-${String(i + 1).padStart(3, '0')}.png`);
    fs.writeFileSync(diffPath, PNG.sync.write(diff));

    // the page percentage says something is wrong; this says which element
    const elements = compareText(A, B, slides[i], w / slides[i].w, cfg.verify);

    pages.push({
      index: i + 1,
      name: slides[i].name,
      status: ratio <= cfg.verify.diffThreshold ? 'ok' : 'check',
      diff: ratio,
      elements,
      refPath: ref,
      gotPath: got,
      diffPath,
    });
  }

  tidy(vDir, pages, cfg, log);
  return { ran: true, engine, pages, dir: vDir };
}

/** Keep the evidence that is worth looking at and throw away the rest.
 *  A clean run is ~50MB of images that all say "identical" — on the Desktop
 *  that is clutter, not a diagnostic. Pages that need attention keep all three
 *  images so there is something to compare. */
function tidy(vDir, pages, cfg, log) {
  const keep = cfg.verify.keepImages || 'flagged';
  if (keep === 'all') return;

  const rm = (p) => { if (p) fs.rmSync(p, { force: true }); };
  rm(path.join(vDir, 'render.pdf'));

  for (const pg of pages) {
    if (keep === 'flagged' && pg.status === 'check') continue;
    rm(pg.refPath); rm(pg.gotPath); rm(pg.diffPath);
    pg.refPath = pg.gotPath = pg.diffPath = null;
  }

  // an empty folder is just another thing to tidy up by hand
  if (!fs.readdirSync(vDir).length) {
    fs.rmSync(vDir, { recursive: true, force: true });
    log('모든 페이지가 기준을 통과해 검수 이미지는 남기지 않았습니다.');
  } else {
    log(`확인이 필요한 페이지의 이미지만 남겼습니다 → ${path.basename(vDir)}/`);
  }
}

module.exports = { verify };
