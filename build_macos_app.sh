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
  --name "音声整備ツール" \
  --icon "assets/audio_maintenance_tool.icns" \
  --collect-data customtkinter \
  --hidden-import tkinterdnd2 \
  --add-data "${TKDND_DIR}:tkinterdnd2/tkdnd" \
  main.py

echo ""
echo "Build complete:"
echo "  ${ROOT_DIR}/dist/音声整備ツール.app"
