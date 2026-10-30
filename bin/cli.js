#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');

// minimal .env loader (no dependency)
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

function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--no-verify') out.noVerify = true;
    else if (a === '--out' || a === '-o') out.outDir = argv[++i];
    else if (a === '--name' || a === '-n') out.name = argv[++i];
    else if (a === '--token') out.token = argv[++i];
    else if (a === '--config' || a === '-c') out.configPath = argv[++i];
    else if (a === '--help' || a === '-h') out.help = true;
    else out._.push(a);
  }
  return out;
}

const HELP = `
figma2pptx — Figma 프레임을 편집 가능한 PowerPoint로 변환합니다.

사용법
  figma2pptx <피그마 URL 또는 파일키> [옵션]

옵션
  -o, --out <폴더>      결과 저장 폴더 (기본: 현재 폴더)
  -n, --name <이름>     파일 이름 (기본: 섹션/페이지 이름)
  -c, --config <파일>   설정 JSON (서체 매핑 등)
      --token <토큰>    Figma 토큰 (기본: .env의 FIGMA_TOKEN)
      --no-verify       변환 후 렌더 비교를 건너뜁니다
  -h, --help            도움말

예시
  figma2pptx "https://www.figma.com/design/AbC123/Deck?node-id=1234-5678" -o ~/Desktop

준비
  1) Figma → Settings → Security → Personal access tokens 에서 토큰 생성
     (스코프: File content = Read)
  2) 프로젝트 폴더의 .env 파일에 FIGMA_TOKEN=... 저장
`;

(async () => {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args._.length) {
    console.log(HELP);
    process.exit(args.help ? 0 : 1);
  }

  const log = (m) => console.log(`  ${m}`);
  try {
    const r = await run({
      target: args._[0],
      outDir: args.outDir,
      name: args.name,
      token: args.token,
      configPath: args.configPath,
      noVerify: args.noVerify,
      log,
    });

    console.log('\n완료');
    console.log(`  파일   ${r.file}`);
    console.log(`  리포트 ${r.reportPath}`);
    console.log(`  구성   슬라이드 ${r.slides} · 텍스트 ${r.texts} · 도형 ${r.shapes} · 이미지 ${r.images}`);
    if (r.missingFonts.length) console.log(`  ⚠️ 미설치 서체: ${r.missingFonts.join(', ')}`);
    if (r.verify && r.verify.ran) {
      const bad = r.verify.pages.filter((p) => p.status === 'check');
      console.log(bad.length ? `  ⚠️ 확인 필요 ${bad.length}장: ${bad.map((p) => 'p' + p.index).join(', ')}` : '  ✅ 렌더 비교 통과');
    } else if (r.verify) {
      console.log(`  ⏭️ 렌더 비교 건너뜀 — ${r.verify.reason}`);
    }
  } catch (e) {
    console.error('\n실패:', e.message);
    process.exit(1);
  }
})();
