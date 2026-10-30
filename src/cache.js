'use strict';
// Disk cache for rendered images.
//
// The slow part of a conversion is not our code — it is Figma rasterising each
// node on its servers, about two seconds apiece, and refusing us with 429 if we
// ask faster. A section with 2605 images therefore takes hours, and every retry
// after a failure started again from zero.
//
// So stop asking twice. Entries are keyed by the file's `version`, which Figma
// bumps on every edit: while the design is untouched a re-run is served entirely
// from disk, and the moment it changes the old entries simply stop matching.
// That makes a failed run resumable, which is what turns a 2605-image section
// from impossible into merely slow.

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

function cacheRoot() {
  return process.env.F2P_CACHE_DIR
    || path.join(os.homedir(), 'Library', 'Caches', 'figma2pptx');
}

function keyFor(parts) {
  return crypto.createHash('sha1').update(parts.join('|')).digest('hex');
}

function open({ fileKey, version, enabled = true }) {
  const dir = path.join(cacheRoot(), `${fileKey}-${version}`);
  let ready = false;
  const ensure = () => {
    if (ready || !enabled) return enabled;
    try { fs.mkdirSync(dir, { recursive: true }); ready = true; } catch { return false; }
    return true;
  };

  let hits = 0, writes = 0;
  return {
    dir,
    get hits() { return hits; },
    get writes() { return writes; },

    /** Is it already on disk? Asking this BEFORE requesting render URLs is the
     *  whole point: the URL call is what makes Figma rasterise the nodes, and
     *  that — not the download — is the slow, rate-limited step. */
    has(id, scale) {
      if (!enabled || !ensure()) return false;
      return fs.existsSync(path.join(dir, keyFor([id, scale])));
    },

    read(id, scale) {
      if (!enabled || !ensure()) return null;
      const p = path.join(dir, keyFor([id, scale]));
      try {
        const buf = fs.readFileSync(p);
        hits++;
        return buf;
      } catch { return null; }
    },

    write(id, scale, buf) {
      if (!enabled || !ensure()) return;
      const p = path.join(dir, keyFor([id, scale]));
      try {
        // write-then-rename so an interrupted run never leaves a half file that
        // would later be served as if it were a complete image
        const tmp = `${p}.${process.pid}.tmp`;
        fs.writeFileSync(tmp, buf);
        fs.renameSync(tmp, p);
        writes++;
      } catch { /* a cache that cannot write is still a working converter */ }
    },
  };
}

/** Retire old versions, keeping the most recent few.
 *
 *  Deleting every other version on sight looked tidy and was destructive: an
 *  edit to one frame in Figma bumps the file version, which threw away 2735
 *  perfectly good renders of the 2604 nodes that had not changed. Keeping a few
 *  generations costs disk and nothing else, and means undoing an edit — or
 *  converting an older version again — is still instant. */
function prune(fileKey, keepVersion, keep = 3) {
  const root = cacheRoot();
  let removed = 0;
  try {
    const mine = fs.readdirSync(root)
      .filter((n) => n.startsWith(`${fileKey}-`) && n !== `${fileKey}-${keepVersion}`)
      .map((n) => ({ n, t: fs.statSync(path.join(root, n)).mtimeMs }))
      .sort((a, b) => b.t - a.t);
    for (const { n } of mine.slice(Math.max(0, keep - 1))) {
      fs.rmSync(path.join(root, n), { recursive: true, force: true });
      removed++;
    }
  } catch { /* nothing cached yet */ }
  return removed;
}

/** Versions of this file we already hold renders for, newest first. */
function versionsOf(fileKey) {
  try {
    return fs.readdirSync(cacheRoot())
      .filter((n) => n.startsWith(`${fileKey}-`))
      .map((n) => ({ v: n.slice(fileKey.length + 1), t: fs.statSync(path.join(cacheRoot(), n)).mtimeMs }))
      .sort((a, b) => b.t - a.t);
  } catch { return []; }
}

module.exports = { open, prune, versionsOf, cacheRoot };
