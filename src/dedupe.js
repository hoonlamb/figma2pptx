'use strict';
// Collapse duplicate images inside a finished .pptx.
//
// pptxgenjs writes one media file per addImage call, so a cover that appears on
// six pages is stored six times. A real deck came out at 75MB holding only 88
// distinct pictures across 114 embeds — big enough that PowerPoint struggles to
// open it. Nothing upstream can prevent this (the packer assigns a fresh
// relationship id per call), so the fix is a pass over the built archive:
// hash every media entry, keep one file per hash, and repoint the relationships
// of the rest at the survivor.

const fs = require('fs');
const crypto = require('crypto');
const JSZip = require('jszip');

async function dedupeMedia(pptxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));

  const canonical = new Map();          // hash  -> kept media path
  const replace = new Map();            // dropped media path -> kept media path
  // zip.files also carries directory entries, which have no content to hash
  const names = Object.keys(zip.files)
    .filter((n) => n.startsWith('ppt/media/') && !zip.files[n].dir);

  for (const name of names.sort()) {
    const buf = await zip.file(name).async('nodebuffer');
    const h = crypto.createHash('sha1').update(buf).digest('hex');
    if (canonical.has(h)) replace.set(name, canonical.get(h));
    else canonical.set(h, name);
  }
  if (!replace.size) return { removed: 0, before: fs.statSync(pptxPath).size, after: fs.statSync(pptxPath).size };

  // Relationship targets are written relative to the .rels file's own folder,
  // e.g. "../media/image3.png" from ppt/slides/_rels — so match on the file name
  // rather than trying to resolve paths.
  const shortMap = new Map();
  for (const [from, to] of replace) {
    shortMap.set(from.split('/').pop(), to.split('/').pop());
  }
  const rels = Object.keys(zip.files).filter((n) => n.endsWith('.rels') && !zip.files[n].dir);
  for (const name of rels) {
    let xml = await zip.file(name).async('string');
    let touched = false;
    for (const [from, to] of shortMap) {
      if (!xml.includes(from)) continue;
      // guard the boundary: image1.png must not match inside image10.png
      const re = new RegExp(`([/"])${from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=")`, 'g');
      const next = xml.replace(re, `$1${to}`);
      if (next !== xml) { xml = next; touched = true; }
    }
    if (touched) zip.file(name, xml);
  }

  for (const name of replace.keys()) zip.remove(name);

  const before = fs.statSync(pptxPath).size;
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(pptxPath, out);
  return { removed: replace.size, kept: canonical.size, before, after: out.length };
}

module.exports = { dedupeMedia };
