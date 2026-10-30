'use strict';
// Keep photographs out of PNG.
//
// Figma renders everything as PNG, which is the right format for icons and
// anything with a transparent edge and the wrong one for a photograph: one real
// deck carried 41 pictures in 56MB, and 27 of them were fully opaque artwork
// that JPEG stores at a fifth of the size. A 56MB deck is slow to open, awkward
// to email, and makes PowerPoint itself sluggish.
//
// The choice is made by looking at the actual pixels rather than guessing from
// the node, because "does this image use its alpha channel" is the only question
// that matters and it has an exact answer.

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { PNG } = require('pngjs');

/** True when every pixel is fully opaque, so nothing is lost by dropping alpha. */
function isOpaque(buf) {
  let img;
  try { img = PNG.sync.read(buf); } catch { return false; }
  for (let i = 3; i < img.data.length; i += 4) if (img.data[i] < 250) return false;
  return true;
}

let sipsChecked = false, sipsOk = false;
function haveSips() {
  if (!sipsChecked) {
    sipsChecked = true;
    try { execFileSync('/usr/bin/sips', ['--help'], { stdio: 'ignore' }); sipsOk = true; } catch { sipsOk = false; }
  }
  return sipsOk;
}

/**
 * @returns { data: base64, mime: 'image/png'|'image/jpeg', saved: bytes }
 * Falls back to the original PNG whenever anything is uncertain — a slightly
 * large deck beats a deck with a hole where a picture used to be.
 */
function encode(buf, cfg = {}) {
  const minBytes = cfg.jpegMinBytes === undefined ? 120 * 1024 : cfg.jpegMinBytes;
  const quality = cfg.jpegQuality === undefined ? 90 : cfg.jpegQuality;
  // Pixel size survives the JPEG step, and pptx.js needs it to place the picture.
  const size = buf.length > 24 ? { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) } : {};
  const keep = { data: buf.toString('base64'), mime: 'image/png', saved: 0, ...size };

  if (cfg.jpeg === false || buf.length < minBytes || !haveSips()) return keep;
  if (!isOpaque(buf)) return keep;             // transparency is load-bearing

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'f2p-img-'));
  const src = path.join(dir, 'a.png'), out = path.join(dir, 'a.jpg');
  try {
    fs.writeFileSync(src, buf);
    execFileSync('/usr/bin/sips', ['-s', 'format', 'jpeg', '-s', 'formatOptions', String(quality), src, '--out', out],
      { stdio: 'ignore' });
    const jpg = fs.readFileSync(out);
    if (jpg.length >= buf.length) return keep;   // nothing gained
    return { data: jpg.toString('base64'), mime: 'image/jpeg', saved: buf.length - jpg.length, ...size };
  } catch {
    return keep;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

module.exports = { encode, isOpaque };
