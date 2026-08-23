#!/bin/bash
# Applies every local patch (see patches/) to the Stimpak submodule and
# builds its native library + auth-window helper. Run this once before
# building Invigoration; Stimpak.csproj's own MSBuild targets
# (StimpakBuildNative=false, see Invigoration.Core.csproj) then just copy
# whatever lands in native/superiority/target/release/ — no separate copy
# step needed here.
#
# Requires a Rust toolchain (cargo). On macOS, building stimpak-auth-window
# also needs Xcode command line tools (for the WebKit/AppKit frameworks wry
# links against); on Linux it needs webkit2gtk (see wry's own build docs).
#
# Build/publish Invigoration by pointing dotnet at a specific project
# (dotnet/src/Invigoration.App/Invigoration.App.csproj — see build-macos.sh
# and build-linux.sh), not at the .slnx solution file. A solution-level build
# uses a different project-graph traversal that doesn't consistently respect
# the AdditionalProperties (AssemblyName=Stimpak.Managed) Core/App reference
# Stimpak.csproj with, and can rebuild it with the wrong (colliding) name —
# see Invigoration.Core.csproj's comment on why that override exists at all.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SUBMODULE_DIR="$SCRIPT_DIR/superiority"
PATCH_DIR="$SCRIPT_DIR/patches"

cd "$SUBMODULE_DIR"

cleanup() {
  # Whatever patches applied cleanly, revert them so the submodule stays a
  # plain, unmodified checkout of the pinned commit between runs.
  git -C "$SUBMODULE_DIR" checkout -- . 2>/dev/null || true
}
trap cleanup EXIT

for patch in "$PATCH_DIR"/*.patch; do
  [[ -e "$patch" ]] || continue
  if git apply --check "$patch" 2>/dev/null; then
    echo "==> Applying $(basename "$patch")"
    git apply "$patch"
  else
    echo "==> Skipping $(basename "$patch") (already applied, or no longer applies cleanly — check it against the current submodule commit)"
  fi
done

echo "==> cargo build -p stimpak -p stimpak-auth-window --release"
cargo build -p stimpak -p stimpak-auth-window --release

# Warms Stimpak.csproj's own build state (its ref-assembly cache) with the
# exact AssemblyName=Stimpak.Managed;StimpakBuildNative=false property set
# every Invigoration project references it with. Without this, the *first*
# `dotnet build`/`dotnet publish` of the whole solution on a fresh checkout
# fails with "Could not find file ...obj\...\Stimpak.Managed.dll" — an
# MSBuild quirk where a ProjectReference's AdditionalProperties-driven build
# needs that project already built once before multi-hop reference
# resolution (Core -> Stimpak, App -> Core -> Stimpak) picks it up cleanly.
echo "==> Warming Stimpak.csproj's managed build (AssemblyName=Stimpak.Managed)"
dotnet build "$SCRIPT_DIR/superiority/stimpak/csharp/Stimpak/Stimpak.csproj" \
  -p:StimpakBuildNative=false -p:AssemblyName=Stimpak.Managed -v quiet

echo "==> Done: $SUBMODULE_DIR/target/release/"
