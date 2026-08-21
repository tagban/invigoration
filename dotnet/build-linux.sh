#!/bin/bash
# Builds a self-contained Linux release for Invigoration.
#
# Usage:
#   ./build-linux.sh [--rid linux-x64|linux-arm64]

set -euo pipefail

RID="linux-x64"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --rid) RID="$2"; shift 2 ;;
    *) echo "Unknown argument: $1"; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VERSION=$(grep -o '"[0-9][^"]*"' src/Invigoration.Core/AppVersion.cs | tr -d '"')

DIST_DIR="$SCRIPT_DIR/dist/$RID"
PUBLISH_DIR="$DIST_DIR/Invigoration-$VERSION-$RID"

echo "==> Publishing self-contained $RID build (version $VERSION)"
rm -rf "$DIST_DIR"
dotnet publish src/Invigoration.App/Invigoration.App.csproj \
  -c Release -r "$RID" --self-contained true \
  -p:PublishSingleFile=false \
  -o "$PUBLISH_DIR"

chmod +x "$PUBLISH_DIR/Invigoration.App"

cat > "$PUBLISH_DIR/Invigoration.sh" <<'EOF'
#!/bin/bash
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec "$DIR/Invigoration.App" "$@"
EOF
chmod +x "$PUBLISH_DIR/Invigoration.sh"

echo "==> Packaging tarball"
TARBALL="$DIST_DIR/Invigoration-$VERSION-$RID.tar.gz"
tar -czf "$TARBALL" -C "$DIST_DIR" "Invigoration-$VERSION-$RID"

echo "==> Done: $TARBALL"
