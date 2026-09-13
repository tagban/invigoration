# Plan: dressed, animated Diablo II characters in the lobby strip

**Status:** planned, not started (2026-09-13). The strip already shows Battle.net's chat avatars
(Moderator, StarCraft marine, Unknown…) on grass; live D2 characters still fall back to their
28×14 Battle.net portrait tile, which needs replacing with the real thing.

## Goal

Every Diablo II character in the Diablo II theme's character dock is drawn the way D2's own
lobby drew them: full body, animated, wearing the armor and weapons their statstring describes,
in the item colors it describes — on any server Invigoration connects to.

## Shape of the solution

Two halves, split along what each project is good at:

| | Command Center (`bnet_command_center`, Rust) | Invigoration (C#) |
|---|---|---|
| Owns | Turning the operator's D2 install into a **character art pack** | Downloading, **keeping**, and drawing from that pack |
| Needs D2 files? | Yes — the operator's own MPQs, which its D2 server already requires | No |
| Delivers | The pack over **BNFTP**, like `icons.bni` | Characters on every server, including ones that don't serve a pack |

Why a pack (the component layers) rather than the server rendering each character: a stored pack
lets Invigoration dress **any** character it ever sees, on **any** server, offline. Per-character
renders would only cover characters already seen on a server that renders.

Why BNFTP: Command Center already speaks it (`bnetcc-proto::bnftp`) and has no HTTP server;
BNFTP's file-time field gives cache validation for free; and Invigoration needs a BNFTP client
anyway for the deferred server ad banner.

Legal alignment: Command Center's `docs/LEGAL.md` says to ship no Blizzard assets and have BNFTP
serve files the operator provides. The pack is generated at runtime from the operator's install
and never committed to either repo.

## Phases

### 0. Research (both projects need this)
- **Gear bytes → animation tokens.** The statstring's 11 equipment bytes (head, torso, legs,
  right/left arm, right/left hand, shield, both shoulder pads, special) name armor-class tokens
  (`lit`/`med`/`hvy`, helm and weapon codes…) that pick a DCC file per body part. Pin down the
  exact mapping. Command Center's realm needs the same mapping to send real gear in its own
  portraits (`realm.rs` currently sends class/status/level/progression only).
- **Color bytes → color transforms**: the 11 color bytes against `colors.txt` and the item
  palette-shift tables.
- **Weapon class**: which of `hth`/`1hs`/`2hs`/`stf`/`bow`/… a character's weapons select.
- **Lobby animation**: which mode (likely town-neutral) and which of the 8/16 directions the chat
  screen uses, and its frame count per class.
- **Check against the real client:** screenshot a few known characters in a 1.14d lobby on
  Command Center; those become the golden references for everything below.

### 1. Command Center — decoders (`d2-formats`)
- `dcc.rs`: DCC decoding (the hard part — variable-bit frame streams, pixel buffers, equal-cell
  optimization). Test against frames exported by an established extraction tool.
- `cof.rs`: layer lists, per-direction/per-frame draw order.
- Palette and colormap loaders.
- Dependency-free, like the rest of the core crates.

### 2. Command Center — pack builder
- `bnetcc d2art build` (and/or on `bnetccd` start when MPQs are present): for each class ×
  weapon class × component slot × token that exists in the MPQs, decode only the lobby mode and
  direction, and write palette-indexed frames with their offsets.
- Pack file: magic + format version + source game version + content hash; palette; colormaps;
  per-class COF draw order; the gear-byte/color-byte mapping tables from phase 0; the frames.
  Target a few MB.
- Deterministic output, so the same install always produces the same hash.

### 3. Command Center — serving
- Serve the pack over BNFTP under a reserved name (e.g. `d2-lobby-art.pak`) with a correct file
  time; a node without D2 files simply doesn't have it.

### 4. Invigoration — fetch and keep
- Minimal BNFTP client (also unblocks the ad banner).
- After connecting, ask the server for the pack (sending the cached copy's file time). Store
  packs in the config folder by content hash; keep the newest valid one and use it on **every**
  server, including ones that don't serve a pack.
- Validate before use (magic, version, hash, bounds on every count and offset — it's data from
  the network).
- Nothing is sent but the file request; no identity beyond the normal connection.

### 5. Invigoration — draw
- Pack reader and compositor: for a character, look up tokens and colors from their statstring,
  stack layers in COF order, apply colormaps, produce RGBA frames, and play them with the existing
  `SpriteAnimationView`.
- Cache composed strips by the 33-byte character struct (identical gear shares one strip).
  Bounded, since a mass-join can bring thousands of characters.
- Dock tile: dressed figure when a pack exists → portrait tile otherwise. Chat avatars (Moderator
  etc.) keep priority exactly as today (`ChatAvatar`).

### 6. Polish
- Idle-animation timing, name plates with titles (already done), render performance under
  load-test floods, and a small Manage Themes/Config note showing which pack is in use and where
  it came from.

## Open decisions
- Whether other (non-Command Center) servers should be able to host a pack too — the format
  would allow it.
- Whether Invigoration should also build a pack from a local D2 install for people who never
  connect to a Command Center node.
