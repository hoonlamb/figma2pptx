#!/bin/bash
# Set up the two ways to launch figma2pptx.
#
#   bash install.sh
#
# No .app bundle: Gatekeeper refuses locally-built, ad-hoc-signed bundles and an
# AppleScript applet that fails does so silently, which is exactly the "I
# double-clicked and nothing happened" case. A .command file opens Terminal, so
# every step — including any failure — is visible.
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"

chmod +x "$DIR/figma2pptx.command" "$DIR/bin/figma2pptx" "$DIR/bin/cli.js" 2>/dev/null || true

# packages: a fresh download has none
if [ ! -d "$DIR/node_modules/pptxgenjs" ]; then
  PATH="$DIR/.node/bin:$PATH" npm --prefix "$DIR" ci --omit=dev --no-audit --no-fund --loglevel=error
fi

# 1) terminal command: figma2pptx
mkdir -p "$HOME/bin"
ln -sf "$DIR/bin/figma2pptx" "$HOME/bin/figma2pptx"
if ! grep -q 'HOME/bin' "$HOME/.zshrc" 2>/dev/null; then
  printf '\n# 개인 실행 파일 (figma2pptx 등)\nexport PATH="$HOME/bin:$PATH"\n' >> "$HOME/.zshrc"
  ADDED_PATH=1
fi

# 2) double-clickable file on the Desktop, pinned to this project folder.
# plain bash: python3 is only a stub on a Mac without developer tools.
DESK="$HOME/Desktop/figma2pptx.command"
LINE='cd "$(dirname "$0")" || exit 1'
PINNED="cd \"$DIR\" || { echo \"  $DIR 폴더를 찾지 못했습니다.\"; read -r; exit 1; }"
while IFS= read -r row || [ -n "$row" ]; do
  if [ "$row" = "$LINE" ]; then printf '%s\n' "$PINNED"; else printf '%s\n' "$row"; fi
done < "$DIR/figma2pptx.command" > "$DESK"
chmod +x "$DESK"

# 3) optional: keep the server running at login, so the only step is a bookmark
PLIST="$HOME/Library/LaunchAgents/com.hoonlamb.figma2pptx.plist"
if [ "$1" = "--autostart" ]; then
  NODE=""
  for c in /usr/local/bin/node /opt/homebrew/bin/node /usr/bin/node "$(command -v node 2>/dev/null)"; do
    [ -x "$c" ] && { NODE="$c"; break; }
  done
  [ -z "$NODE" ] && { echo "Node.js를 찾지 못했습니다."; exit 1; }
  mkdir -p "$HOME/Library/LaunchAgents"
  sed -e "s|__NODE__|$NODE|g" -e "s|__DIR__|$DIR|g" \
      "$DIR/tools/com.hoonlamb.figma2pptx.plist" > "$PLIST"
  launchctl bootout "gui/$(id -u)/com.hoonlamb.figma2pptx" 2>/dev/null || true
  launchctl bootstrap "gui/$(id -u)" "$PLIST"
  AUTOSTART=1
elif [ "$1" = "--no-autostart" ]; then
  launchctl bootout "gui/$(id -u)/com.hoonlamb.figma2pptx" 2>/dev/null || true
  rm -f "$PLIST"
  echo "자동 실행을 해제했습니다."
fi

echo "완료"
if [ -n "$AUTOSTART" ]; then
  echo "  로그인할 때 자동으로 켜집니다. 북마크만 누르세요 → http://127.0.0.1:4785"
  echo "  끄려면: bash install.sh --no-autostart"
else
  echo "  바탕화면   figma2pptx.command  ← 더블클릭"
  echo "  터미널     figma2pptx"
  echo "  터미널 없이 쓰려면: bash install.sh --autostart"
fi
[ -n "$ADDED_PATH" ] && echo "  (터미널 명령은 새 터미널 창부터 인식됩니다)"
exit 0
