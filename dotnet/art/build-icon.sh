#!/bin/bash
# Rebuilds packaging/AppIcon.icns from src/Invigoration.App/Assets/flag.ico — the flag on drawn
# wood planks (see make-icon.swift for the design and why macOS 26 needs the backing). macOS only:
# uses swiftc, sips and iconutil, all part of the Xcode command line tools.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
PACKAGING="$(cd "$HERE/../packaging" && pwd)"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

sips -s format png "$HERE/../src/Invigoration.App/Assets/flag.ico" --out "$WORK/flag.png" >/dev/null
swiftc -O "$HERE/make-icon.swift" -o "$WORK/make-icon"
"$WORK/make-icon" "$WORK/flag.png" "$WORK/icon-1024.png"

# Every size from the one 1024 render: shrinking smooths fine, it's only enlarging pixel art that blurs.
# No 16 or 32 point sizes, on purpose: macOS 26 puts those two sizes of any old-style icns on its
# plate whatever they look like (checked against installed apps too), and without them it scales
# the 128 down instead — which it leaves alone. Rendered through NSWorkspace, the same icon with
# them still plated at 32px in list views; without them, clean at every size.
mkdir "$WORK/AppIcon.iconset"
for size in 128 256 512; do
    sips -z "$size" "$size" "$WORK/icon-1024.png" --out "$WORK/AppIcon.iconset/icon_${size}x${size}.png" >/dev/null
    sips -z $((size * 2)) $((size * 2)) "$WORK/icon-1024.png" --out "$WORK/AppIcon.iconset/icon_${size}x${size}@2x.png" >/dev/null
done

iconutil -c icns "$WORK/AppIcon.iconset" -o "$PACKAGING/AppIcon.icns"
echo "Wrote $PACKAGING/AppIcon.icns"
