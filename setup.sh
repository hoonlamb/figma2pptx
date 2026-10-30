#!/bin/bash
# One-line setup for people who have never opened a terminal before:
#
#   curl -fsSL https://raw.githubusercontent.com/hoonlamb/figma2pptx/main/setup.sh | bash
#
# Needs no git and no admin password. Everything lands in ~/figma2pptx, and Node is
# fetched into that folder only when the Mac does not already have 18 or newer.
set -euo pipefail

NODE_VERSION="v24.22.2"
DEST="${F2P_DIR:-$HOME/figma2pptx}"
TARBALL="${F2P_TARBALL:-https://codeload.github.com/hoonlamb/figma2pptx/tar.gz/refs/heads/main}"

say() { printf '  %s\n' "$1"; }

[ "$(uname -s)" = "Darwin" ] || { say "지금은 맥에서만 돌아갑니다."; exit 1; }

say "내려받는 중…"
mkdir -p "$DEST"
curl -fsSL "$TARBALL" | tar -xz --strip-components=1 -C "$DEST"

node_ok() {
  command -v node >/dev/null 2>&1 && [ "$(node -p 'process.versions.node.split(".")[0]')" -ge 18 ]
}
if ! node_ok; then
  case "$(uname -m)" in
    arm64) ARCH=arm64 ;;
    *) ARCH=x64 ;;
  esac
  NAME="node-$NODE_VERSION-darwin-$ARCH.tar.gz"
  BASE="https://nodejs.org/dist/$NODE_VERSION"
  TMP="$(mktemp -d)"
  trap 'rm -rf "$TMP"' EXIT
  say "Node 내려받는 중… (이 폴더 안에만 들어갑니다)"
  curl -fsSL "$BASE/$NAME" -o "$TMP/$NAME"
  WANT="$(curl -fsSL "$BASE/SHASUMS256.txt" | awk -v f="$NAME" '$2 == f {print $1}')"
  GOT="$(shasum -a 256 "$TMP/$NAME" | awk '{print $1}')"
  [ -n "$WANT" ] && [ "$WANT" = "$GOT" ] || { say "Node 파일 검증에 실패했습니다. 다시 시도해 주세요."; exit 1; }
  mkdir -p "$DEST/.node"
  tar -xzf "$TMP/$NAME" --strip-components=1 -C "$DEST/.node"
fi

say "설치하는 중…"
bash "$DEST/install.sh"

if ! xcode-select -p >/dev/null 2>&1; then
  say "변환은 바로 됩니다. PowerPoint 자동 비교까지 쓰려면 터미널에서 xcode-select --install 을 한 번 실행해 주세요."
fi
say "바탕화면의 figma2pptx.command 를 더블클릭하면 시작합니다."
