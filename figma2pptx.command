#!/bin/bash
# Double-clickable launcher.
#
# A .command file is used rather than an AppleScript .app on purpose: Gatekeeper
# refuses locally-built, ad-hoc-signed bundles, and an applet that fails does so
# silently. Here the Terminal window IS the program — every step is visible, and
# closing the window stops the server.

cd "$(dirname "$0")" || exit 1
PORT="${PORT:-4785}"
URL="http://127.0.0.1:$PORT"          # never "localhost": macOS resolves that to
                                      # ::1 first, and Chrome will not fall back

printf '\n  figma2pptx\n  ─────────────────────────────────────────\n\n'

# --- node ---------------------------------------------------------------------
NODE=""
for c in "$PWD/.node/bin/node" /usr/local/bin/node /opt/homebrew/bin/node /usr/bin/node "$(command -v node 2>/dev/null)"; do
  [ -x "$c" ] && { NODE="$c"; break; }
done
if [ -z "$NODE" ]; then
  echo "  ❌ Node.js를 찾지 못했습니다."
  echo "     https://nodejs.org 에서 설치한 뒤 이 파일을 다시 실행해 주세요."
  echo; read -r -p "  엔터를 누르면 창이 닫힙니다..." _; exit 1
fi

# --- already running? ---------------------------------------------------------
if curl -s -m 1 -o /dev/null "$URL/status" 2>/dev/null; then
  echo "  이미 실행 중입니다 → $URL"
  open "$URL"
  echo; read -r -p "  엔터를 누르면 이 창이 닫힙니다..." _; exit 0
fi

# --- go -----------------------------------------------------------------------
echo "  주소   $URL"
echo "  끄려면 이 창에서 Control+C 를 누르거나 창을 닫으세요."
echo

( for _ in $(seq 1 40); do
    sleep 0.4
    curl -s -m 1 -o /dev/null "$URL/status" 2>/dev/null && { open "$URL"; break; }
  done ) &

# foreground so the window's lifetime is the server's lifetime
F2P_NO_OPEN=1 PORT="$PORT" exec "$NODE" gui/server.js
