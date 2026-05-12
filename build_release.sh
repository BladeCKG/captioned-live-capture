#!/usr/bin/env bash
set -euo pipefail

APP_NAME="${1:-CaptionedLiveCapture}"

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DIST_DIR="$ROOT_DIR/dist"
BUILD_DIR="$ROOT_DIR/build"
RELEASE_DIR="$ROOT_DIR/release"
APP_BUNDLE="$DIST_DIR/$APP_NAME.app"
ZIP_PATH="$RELEASE_DIR/$APP_NAME-macos.zip"

find_python() {
  local candidate
  for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done
  echo "Python was not found in PATH. Install Python and try again." >&2
  exit 1
}

PYTHON_CMD="$(find_python)"

cd "$ROOT_DIR"

echo "Installing build requirements..."
PYTHONIOENCODING="utf-8" "$PYTHON_CMD" -m pip install -r requirements.txt pyinstaller

echo "Cleaning old build outputs..."
rm -rf "$BUILD_DIR" "$DIST_DIR"
mkdir -p "$RELEASE_DIR"
rm -f "$ZIP_PATH"

echo "Building macOS app bundle..."
PYTHONIOENCODING="utf-8" "$PYTHON_CMD" -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --hidden-import ApplicationServices \
  --hidden-import Quartz \
  capture_text_app.py

if [[ ! -d "$APP_BUNDLE" ]]; then
  echo "Build output was not created: $APP_BUNDLE" >&2
  exit 1
fi

echo "Creating portable zip..."
ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE" "$ZIP_PATH"

echo "Release created:"
echo "$ZIP_PATH"
