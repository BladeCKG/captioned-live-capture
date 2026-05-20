#!/usr/bin/env bash
set -euo pipefail

APP_NAME="${1:-CaptionedLiveCapture}"

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DIST_DIR="$ROOT_DIR/dist"
BUILD_DIR="$ROOT_DIR/build"
RELEASE_DIR="$ROOT_DIR/release"
VENV_DIR="$ROOT_DIR/.build-venv"
APP_BUNDLE="$DIST_DIR/$APP_NAME.app"
ZIP_PATH="$RELEASE_DIR/$APP_NAME-macos.zip"

find_python() {
  if [[ -n "${PYTHON:-}" ]]; then
    if command -v "$PYTHON" >/dev/null 2>&1; then
      if ensure_python_has_working_tk "$PYTHON"; then
        printf '%s\n' "$PYTHON"
        return 0
      else
        echo "PYTHON is set but Tkinter cannot start with it: $PYTHON" >&2
        exit 1
      fi
    fi
    echo "PYTHON is set but was not found: $PYTHON" >&2
    exit 1
  fi

  local candidate
  for candidate in /usr/local/bin/python3 /opt/homebrew/bin/python3 python3 python /usr/bin/python3; do
    if command -v "$candidate" >/dev/null 2>&1 && ensure_python_has_working_tk "$candidate"; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done
  echo "No Python with working Tkinter was found." >&2
  echo "Install Python with Tk support, then rerun this script." >&2
  exit 1
}

python_has_working_tk() {
  "$1" -c "import tkinter as tk; root = tk.Tk(); root.withdraw(); root.destroy()" >/dev/null 2>&1
}

ensure_python_has_working_tk() {
  local python_cmd="$1"
  if python_has_working_tk "$python_cmd"; then
    return 0
  fi

  if ! command -v brew >/dev/null 2>&1; then
    return 1
  fi

  local resolved_python version formula
  resolved_python="$(python_realpath "$python_cmd")"
  case "$resolved_python" in
    /usr/local/*|/opt/homebrew/*) ;;
    *) return 1 ;;
  esac

  version="$("$python_cmd" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null || true)"
  if [[ -z "$version" ]]; then
    return 1
  fi

  formula="python-tk@$version"
  echo "Tkinter is not available for $python_cmd. Installing $formula with Homebrew..." >&2
  brew install "$formula" >&2
  python_has_working_tk "$python_cmd"
}

python_realpath() {
  if command -v realpath >/dev/null 2>&1; then
    realpath "$1"
  else
    python3 -c "import os, sys; print(os.path.realpath(sys.argv[1]))" "$1"
  fi
}

PYTHON_CMD="$(find_python)"

cd "$ROOT_DIR"

echo "Creating isolated build environment..."
rm -rf "$VENV_DIR"
"$PYTHON_CMD" -m venv "$VENV_DIR"
VENV_PYTHON="$VENV_DIR/bin/python"

echo "Installing build requirements..."
PYTHONIOENCODING="utf-8" "$VENV_PYTHON" -m pip install --upgrade pip setuptools wheel
PYTHONIOENCODING="utf-8" "$VENV_PYTHON" -m pip install -r requirements.txt pyinstaller

echo "Cleaning old build outputs..."
rm -rf "$BUILD_DIR" "$DIST_DIR"
mkdir -p "$RELEASE_DIR"
rm -f "$ZIP_PATH"

echo "Building macOS app bundle..."
PYTHONIOENCODING="utf-8" "$VENV_PYTHON" -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --hidden-import ApplicationServices \
  --hidden-import Quartz \
  --hidden-import captioned_windows \
  --hidden-import captioned_macos \
  --hidden-import chrome_live_caption_windows \
  --hidden-import chrome_live_caption_macos \
  --hidden-import macos_accessibility \
  --collect-all ApplicationServices \
  --collect-all Quartz \
  --collect-all objc \
  capture_text_app.py

if [[ ! -d "$APP_BUNDLE" ]]; then
  echo "Build output was not created: $APP_BUNDLE" >&2
  exit 1
fi

echo "Creating portable zip..."
ditto -c -k --sequesterRsrc --keepParent "$APP_BUNDLE" "$ZIP_PATH"

echo "Release created:"
echo "$ZIP_PATH"
