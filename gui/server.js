#!/usr/bin/env node
'use strict';
// Tiny local GUI: paste a Figma link, press a button, watch the log.
// Plain Node http + SSE so there is nothing extra to install or keep updated.

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');

(function loadEnv() {
  for (const dir of [process.cwd(), path.join(__dirname, '..')]) {
    const p = path.join(dir, '.env');
    if (!fs.existsSync(p)) continue;
    for (const line of fs.readFileSync(p, 'utf8').split('\n')) {
      const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/i);
      if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
    }
  }
})();

const { run } = require('../src/pipeline');
const { isLocalRequest } = require('./guard');
const PORT = Number(process.env.PORT || 4785);
const ENV_PATH = path.join(__dirname, '..', '.env');

// Storing a token without trying it just moves the failure to conversion time,
// where it reads as "the converter is broken". Figma has no token-introspection
// endpoint, so ask for a syntactically valid but nonexistent file: a token that
// can read file content gets 404 (no such file), anything else gets 403.
//
// The 403 body is always {"err":"Invalid token"} — Figma does NOT distinguish a
// revoked/expired token from one that is merely missing file_content:read
// (verified against a real comments-only token). So the message has to name
// every cause rather than guess one.
// Returns a message to show the user, or null when the token is good.
const PROBE_KEY = 'AAAAAAAAAAAAAAAAAAAAAA';
async function tokenProblem(token) {
  let r;
  try {
    r = await fetch(`https://api.figma.com/v1/files/${PROBE_KEY}`, {
      headers: { 'X-Figma-Token': token },
    });
  } catch {
    return '인터넷에 연결되지 않아 토큰을 확인하지 못했습니다.';
  }
  if (r.status === 404) return null;                    // reached the file API
  if (r.status === 403) {
    return '이 토큰으로는 파일을 읽을 수 없습니다. Figma가 원인을 알려주지 않으니 셋 다 확인해 주세요.\n'
         + '• 토큰을 만들 때 Scopes에서 file_content:read 를 체크했는지 '
         + '(file_comments:read는 댓글만 읽는 권한이라 변환에 쓸 수 없습니다)\n'
         + '• 만료되지 않았는지 (Expiration을 1 day로 두면 다음 날 죽습니다)\n'
         + '• 토큰 문자열이 잘리거나 공백이 섞이지 않았는지';
  }
  if (r.status === 429) return 'Figma 요청 한도에 걸렸습니다. 잠시 후 다시 시도해 주세요.';
  return null;                                          // unexpected: don't block
}

const STARTED = Date.now();
const SRC_DIR = path.join(__dirname, '..', 'src');

/** Source files modified since this process loaded them. */
function staleSources() {
  try {
    return fs.readdirSync(SRC_DIR)
      .filter((f) => f.endsWith('.js'))
      .filter((f) => fs.statSync(path.join(SRC_DIR, f)).mtimeMs > STARTED)
      .sort();
  } catch {
    return [];
  }
}

const handler = async (req, res) => {
  if (!isLocalRequest(req, PORT)) { res.writeHead(403); return res.end('forbidden'); }
  const url = new URL(req.url, `http://127.0.0.1:${PORT}`);

  if (url.pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    return res.end(fs.readFileSync(path.join(__dirname, 'index.html')));
  }

  if (url.pathname === '/status') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({
      hasToken: !!process.env.FIGMA_TOKEN,
      managed: process.env.F2P_MANAGED === '1',
      defaultOut: path.join(os.homedir(), 'Desktop', 'figma2pptx'),
    }));
  }

  if (url.pathname === '/token' && req.method === 'POST') {
    const body = await readBody(req);
    const token = (JSON.parse(body || '{}').token || '').trim();
    if (!token) { res.writeHead(400); return res.end('토큰을 입력해 주세요.'); }

    const bad = await tokenProblem(token);
    if (bad) { res.writeHead(400); return res.end(bad); }

    // stored locally only, never transmitted anywhere but api.figma.com
    let existing = fs.existsSync(ENV_PATH) ? fs.readFileSync(ENV_PATH, 'utf8') : '';
    existing = existing.replace(/^FIGMA_TOKEN=.*$/m, '').trim();
    fs.writeFileSync(ENV_PATH, (existing ? existing + '\n' : '') + `FIGMA_TOKEN=${token}\n`, { mode: 0o600 });
    fs.chmodSync(ENV_PATH, 0o600);    // mode above only applies to a new file
    process.env.FIGMA_TOKEN = token;
    res.writeHead(200); return res.end('ok');
  }

  // let a wrong token be corrected instead of being a dead end
  if (url.pathname === '/token' && req.method === 'DELETE') {
    if (fs.existsSync(ENV_PATH)) {
      const kept = fs.readFileSync(ENV_PATH, 'utf8').replace(/^FIGMA_TOKEN=.*$/m, '').trim();
      fs.writeFileSync(ENV_PATH, kept ? kept + '\n' : '', { mode: 0o600 });
    }
    delete process.env.FIGMA_TOKEN;
    res.writeHead(200); return res.end('ok');
  }

  if (url.pathname === '/convert') {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
    });
    const send = (type, data) => res.write(`data: ${JSON.stringify({ type, data })}\n\n`);

    // Node reads a module once, at startup. A server left running for a day
    // keeps converting with the code it was born with, so fixes made since then
    // silently do nothing — this exact confusion has now cost two debugging
    // sessions. Compare the sources on disk against our own start time and say so.
    const stale = staleSources();
    if (stale.length) {
      send('log', `⚠️ 이 서버는 ${new Date(STARTED).toLocaleString('ko-KR')}에 시작됐고, 그 뒤 코드가 바뀌었습니다 (${stale.join(', ')}).`);
      send('log', '   최신 코드로 변환하려면 이 창을 닫고 figma2pptx를 다시 실행해 주세요.');
    }

    try {
      const report = await run({
        target: url.searchParams.get('target'),
        outDir: url.searchParams.get('out') || path.join(os.homedir(), 'Desktop', 'figma2pptx'),
        name: url.searchParams.get('name') || undefined,
        noVerify: url.searchParams.get('verify') === '0',
        log: (m) => send('log', m),
      });
      send('done', report);
    } catch (e) {
      send('error', e.message);
    }
    return res.end();
  }

  if (url.pathname === '/quit') {
    res.writeHead(200); res.end('bye');
    setTimeout(() => process.exit(0), 200);
    return;
  }

  if (url.pathname === '/reveal') {
    const target = url.searchParams.get('path');
    if (target && fs.existsSync(target)) execFile('/usr/bin/open', ['-R', target]);
    res.writeHead(200); return res.end('ok');
  }

  res.writeHead(404); res.end('not found');
};

function readBody(req) {
  return new Promise((resolve) => {
    let b = '';
    req.on('data', (c) => (b += c));
    req.on('end', () => resolve(b));
  });
}

// macOS resolves "localhost" to ::1 before 127.0.0.1, and Chrome does not retry
// on the other family — a v4-only bind therefore looks completely dead in Chrome
// while curl (which does fall back) reports everything fine. Bind both loopback
// families, and never hand out a "localhost" URL.
const HOSTS = ['127.0.0.1', '::1'];
let ready = 0, failed = 0;

for (const host of HOSTS) {
  const s = http.createServer(handler);
  s.on('error', (e) => {
    if (e.code === 'EADDRINUSE') {
      console.error(`포트 ${PORT}이 이미 사용 중입니다.`);
      console.error(`이미 실행 중이라면 http://127.0.0.1:${PORT} 를 여세요.`);
      console.error(`다른 프로그램이 쓰고 있다면: PORT=4786 npm run gui`);
      process.exit(1);
    }
    // ::1 is absent on IPv6-disabled machines; one live socket is enough
    if (++failed === HOSTS.length) { console.error('서버를 열지 못했습니다:', e.message); process.exit(1); }
  });
  s.listen(PORT, host, () => {
    if (++ready > 1) return;
    const url = `http://127.0.0.1:${PORT}`;
    console.log(`figma2pptx GUI → ${url}`);
    // The .app launcher opens the browser itself (it also handles the
    // already-running case), so it suppresses this to avoid a duplicate tab.
    if (process.env.F2P_NO_OPEN !== '1') execFile('/usr/bin/open', [url]);
  });
}
