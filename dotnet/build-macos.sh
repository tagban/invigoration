#!/bin/bash
# Builds, signs, and (optionally) notarizes a macOS .app bundle for Invigoration.
#
# Usage:
#   ./build-macos.sh [--rid osx-arm64|osx-x64] [--notarize-profile PROFILE_NAME] [--no-sign]
#
# First-time notarization setup (run once, stores credentials in Keychain):
#   xcrun notarytool store-credentials "invigoration-notary" \
#     --apple-id "you@example.com" --team-id "YOURTEAMID" --password "app-specific-password"
#
# Then build + sign + notarize with:
#   ./build-macos.sh --notarize-profile invigoration-notary

set -euo pipefail

RID="osx-arm64"
NOTARIZE_PROFILE=""
SIGN=true

while [[ $# -gt 0 ]]; do
  case "$1" in
    --rid) RID="$2"; shift 2 ;;
    --notarize-profile) NOTARIZE_PROFILE="$2"; shift 2 ;;
    --no-sign) SIGN=false; shift ;;
    *) echo "Unknown argument: $1"; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VERSION=$(grep -o '"[0-9][^"]*"' src/Invigoration.Core/AppVersion.cs | tr -d '"')
NUMERIC_VERSION=$(echo "$VERSION" | grep -o '^[0-9]*\.[0-9]*\.[0-9]*')

DIST_DIR="$SCRIPT_DIR/dist/macos-$RID"
APP_BUNDLE="$DIST_DIR/Invigoration.app"
PUBLISH_DIR="$SCRIPT_DIR/src/Invigoration.App/bin/Release/net10.0/$RID/publish"

echo "==> Publishing self-contained $RID build (version $VERSION)"
rm -rf "$DIST_DIR"
dotnet publish src/Invigoration.App/Invigoration.App.csproj \
  -c Release -r "$RID" --self-contained true \
  -p:PublishSingleFile=false \
  -o "$PUBLISH_DIR"

echo "==> Assembling .app bundle"
mkdir -p "$APP_BUNDLE/Contents/MacOS" "$APP_BUNDLE/Contents/Resources"
cp -R "$PUBLISH_DIR"/* "$APP_BUNDLE/Contents/MacOS/"
cp "$SCRIPT_DIR/packaging/AppIcon.icns" "$APP_BUNDLE/Contents/Resources/AppIcon.icns"
sed "s/__VERSION__/$NUMERIC_VERSION/g" "$SCRIPT_DIR/packaging/Info.plist.template" > "$APP_BUNDLE/Contents/Info.plist"
chmod +x "$APP_BUNDLE/Contents/MacOS/Invigoration.App"

if [ "$SIGN" = true ]; then
  IDENTITY=$(security find-identity -v -p codesigning | grep "Developer ID Application" | head -1 | awk -F'"' '{print $2}' || true)
  if [ -z "$IDENTITY" ]; then
    echo "!! No 'Developer ID Application' certificate found in Keychain."
    echo "!! Falling back to ad-hoc signing (may still be killed by macOS on launch)."
    IDENTITY="-"
  else
    echo "==> Signing with identity: $IDENTITY"
  fi

  codesign --deep --force --timestamp --options runtime \
    --entitlements "$SCRIPT_DIR/packaging/entitlements.plist" \
    --sign "$IDENTITY" "$APP_BUNDLE"

  echo "==> Verifying signature"
  codesign --verify --deep --strict --verbose=2 "$APP_BUNDLE"
fi

if [ -n "$NOTARIZE_PROFILE" ]; then
  echo "==> Zipping for notarization"
  ZIP_PATH="$DIST_DIR/Invigoration-notarize.zip"
  ditto -c -k --keepParent "$APP_BUNDLE" "$ZIP_PATH"

  echo "==> Submitting to Apple notary service (this can take a few minutes)"
  xcrun notarytool submit "$ZIP_PATH" --keychain-profile "$NOTARIZE_PROFILE" --wait

  echo "==> Stapling notarization ticket"
  xcrun stapler staple "$APP_BUNDLE"
  rm -f "$ZIP_PATH"
fi

echo "==> Done: $APP_BUNDLE"
