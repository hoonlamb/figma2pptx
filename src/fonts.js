'use strict';
// Figma reports a family plus a numeric weight; PowerPoint addresses weights by
// typeface name (e.g. "Pretendard Medium") except for regular/bold, which use the
// bold attribute. This resolves one to the other and reports what it could not find.

const { execFileSync } = require('child_process');

let installedCache = null;

/** Font family/full names available to the OS (macOS). Empty set = unknown, skip checks. */
function installedFonts() {
  if (installedCache) return installedCache;
  installedCache = new Set();
  if (process.platform !== 'darwin') return installedCache;
  try {
    const out = execFileSync('system_profiler', ['-json', 'SPFontsDataType'], {
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
      timeout: 60000,
    });
    const data = JSON.parse(out).SPFontsDataType || [];
    for (const f of data) {
      for (const t of f.typefaces || []) {
        if (t._name) installedCache.add(t._name);
        if (t.family) installedCache.add(t.family);
        if (t.full_name) installedCache.add(t.full_name);
      }
    }
  } catch {
    /* font checking is advisory only */
  }
  return installedCache;
}

function makeResolver(cfg) {
  const familyMap = (cfg.fonts && cfg.fonts.familyMap) || {};
  const weights = (cfg.fonts && cfg.fonts.weights) || {};
  const installed = installedFonts();
  const missing = new Set();

  return {
    missing,
    /** @returns {{face:string, bold:boolean}} */
    resolve(style) {
      const rawFamily = (style && style.fontFamily) || 'Pretendard';
      const family = familyMap[rawFamily] || rawFamily;
      const weight = String(style && style.fontWeight ? style.fontWeight : 400);
      const rule = weights[weight] || weights['400'] || { suffix: '' };
      const face = (family + (rule.suffix || '')).trim();
      const bold = !!rule.bold;

      if (installed.size && !installed.has(face) && !installed.has(family)) missing.add(face);
      return { face, bold };
    },
  };
}

module.exports = { makeResolver, installedFonts };
