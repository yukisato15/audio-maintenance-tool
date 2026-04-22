#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

TKDND_DIR="$(python3 -c 'import pathlib, tkinterdnd2; print((pathlib.Path(tkinterdnd2.__file__).resolve().parent / "tkdnd").as_posix())')"

rm -rf build dist
python3 generate_app_icon.py >/dev/null
iconutil -c icns "assets/icon.iconset" -o "assets/audio_maintenance_tool.icns"

python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "AudioMaintenanceTool" \
  --osx-bundle-identifier "com.yukisato.audio-maintenance-tool" \
  --icon "assets/audio_maintenance_tool.icns" \
  --collect-data customtkinter \
  --hidden-import tkinterdnd2 \
  --add-data "${TKDND_DIR}:tkinterdnd2/tkdnd" \
  main.py

APP_PATH="dist/AudioMaintenanceTool.app"
PLIST_PATH="${APP_PATH}/Contents/Info.plist"

/usr/libexec/PlistBuddy -c "Set :CFBundleDisplayName 音声整備ツール" "$PLIST_PATH"
/usr/libexec/PlistBuddy -c "Set :CFBundleName 音声整備ツール" "$PLIST_PATH"

xattr -cr "$APP_PATH" || true
codesign --force --deep --sign - "$APP_PATH"

echo ""
echo "Build complete:"
echo "  ${ROOT_DIR}/${APP_PATH}"
