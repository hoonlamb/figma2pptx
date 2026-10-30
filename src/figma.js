'use strict';
// Figma REST API client.
// REST is used instead of the plugin/MCP path on purpose: node payloads arrive whole
// (no chunked reads) and every node carries absoluteBoundingBox, so group nesting
// never has to be reconstructed by hand.

const API = 'https://api.figma.com/v1';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// 429 and 5xx are Figma asking us to wait, not telling us the request was wrong.
// A big deck needs dozens of calls, and losing minutes of work to one transient
// refusal is the difference between a tool that works on real files and one that
// only works on small ones. Retry those; fail fast on 403/404, which retrying
// cannot fix.
const RETRY_STATUS = new Set([429, 500, 502, 503, 504]);
let onRetry = () => {};
function setRetryLogger(fn) { onRetry = fn || (() => {}); }

async function api(path, token, opts = {}) {
  const tries = opts.tries || 5;
  for (let attempt = 1; ; attempt++) {
    let res;
    try {
      res = await fetch(API + path, { headers: { 'X-Figma-Token': token } });
    } catch (e) {
      if (attempt >= tries) throw new Error(`Figma에 연결하지 못했습니다: ${e.message}`);
      await sleep(backoff(attempt));
      continue;
    }
    if (res.ok) return res.json();

    if (RETRY_STATUS.has(res.status) && attempt < tries) {
      // the server's own advice beats our guess when it gives one
      const ra = Number(res.headers.get('retry-after'));
      const wait = Number.isFinite(ra) && ra > 0 ? ra * 1000 : backoff(attempt);
      onRetry(`Figma가 ${res.status}로 응답해 ${Math.round(wait / 1000)}초 후 다시 시도합니다 (${attempt}/${tries - 1})`);
      await sleep(wait);
      continue;
    }

    const body = await res.text().catch(() => '');
    const hint =
      res.status === 403 && /expired/i.test(body) ? ' (토큰이 만료되었습니다. Figma → Settings → Security 에서 새로 만들어 주세요)'
      : res.status === 403 ? ' (토큰 권한을 확인하세요 — file_content:read 스코프 필요)'
      : res.status === 404 ? ' (파일 키 또는 노드 ID가 잘못되었습니다)'
      : res.status === 429 ? ' (요청 한도에 걸렸습니다. 잠시 후 다시 시도하세요)'
      : '';
    throw new Error(`Figma API ${res.status}${hint}\n${body.slice(0, 400)}`);
  }
}

/** Exponential with jitter, so parallel workers don't retry in lockstep. */
function backoff(attempt) {
  return Math.min(30000, 1000 * 2 ** (attempt - 1)) + Math.floor(Math.random() * 500);
}

/** Accepts a full figma.com URL or a bare file key. */
function parseTarget(input) {
  const s = String(input || '').trim();
  const m = s.match(/figma\.com\/(?:file|design|board|slides)\/([A-Za-z0-9]+)/);
  const fileKey = m ? m[1] : (/^[A-Za-z0-9]{10,}$/.test(s) ? s : null);
  if (!fileKey) throw new Error(`피그마 파일 URL 또는 파일 키를 인식하지 못했습니다: ${input}`);

  let nodeId = null;
  const q = s.match(/[?&]node-id=([^&]+)/);
  if (q) nodeId = decodeURIComponent(q[1]).replace('-', ':');
  return { fileKey, nodeId };
}

const getFile = (fileKey, token) => api(`/files/${fileKey}`, token);

async function getNodes(fileKey, ids, token) {
  const out = {};
  // Figma bumps `version` on every edit, which is exactly what the image cache
  // needs to know whether yesterday's renders still describe today's design.
  let version = null;
  // URL length is the real limit here, so request in modest batches.
  for (let i = 0; i < ids.length; i += 20) {
    const batch = ids.slice(i, i + 20);
    const j = await api(`/files/${fileKey}/nodes?ids=${encodeURIComponent(batch.join(','))}`, token);
    Object.assign(out, j.nodes || {});
    if (j.version) version = j.version;
  }
  // non-enumerable so callers that iterate node ids are unaffected
  Object.defineProperty(out, '__version', { value: version, enumerable: false });
  return out;
}

/** Rendered PNG/SVG URLs for nodes we cannot express as native PPT shapes. */
async function getImageUrls(fileKey, ids, token, { format = 'png', scale = 3, tries } = {}) {
  const out = {};
  for (let i = 0; i < ids.length; i += 40) {
    const batch = ids.slice(i, i + 40);
    const j = await api(
      `/images/${fileKey}?ids=${encodeURIComponent(batch.join(','))}&format=${format}&scale=${scale}`,
      token,
      { tries }
    );
    Object.assign(out, j.images || {});
  }
  return out;
}

/** imageRef -> CDN url, for nodes painted with an IMAGE fill. */
async function getImageFills(fileKey, token) {
  const j = await api(`/files/${fileKey}/images`, token);
  return (j.meta && j.meta.images) || {};
}

async function download(url, tries = 4) {
  for (let attempt = 1; ; attempt++) {
    try {
      const res = await fetch(url);
      if (res.ok) return Buffer.from(await res.arrayBuffer());
      if (!RETRY_STATUS.has(res.status) || attempt >= tries) {
        throw new Error(`이미지 다운로드 실패 ${res.status}: ${url}`);
      }
    } catch (e) {
      // a socket error midway through thousands of images is normal, not fatal
      if (attempt >= tries) throw e;
    }
    await sleep(backoff(attempt));
  }
}

module.exports = { parseTarget, getFile, getNodes, getImageUrls, getImageFills, download, setRetryLogger };
