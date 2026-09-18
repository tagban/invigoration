#!/bin/bash
# Rebuilds the Warcraft III theme's wood-grain overlays (see make-wood-grain.swift) into the app's
# Assets/Themes. macOS only: needs swiftc from the Xcode command line tools.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
OUT="$(cd "$HERE/../src/Invigoration.App/Assets" && pwd)/Themes"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

mkdir -p "$OUT"
swiftc -O "$HERE/make-wood-grain.swift" -o "$WORK/make-wood-grain"
"$WORK/make-wood-grain" "$OUT/wood-grain.png" "$OUT/wood-grain-vertical.png"
