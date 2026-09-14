# Diablo II characters in your bot: `d2-equipment.json` and `d2-characters.zip`

Diablo II's lobby drew every character standing in the gear they wore. A chat bot can do the same:
a Command Center server (bnet.cc) builds two files from its own Diablo II 1.14d install and serves
them over **BNFTP**, the same way Battle.net serves `icons.bni`. Download them once, keep them, and
use them on any server.

| File | Size | What it's for | `format` / `version` |
|---|---|---|---|
| `d2-equipment.json` | ~90 KB | **Naming** what a character wears ("Shako", "Dusk Shroud", "Crystal Blue") | `bnetcc-d2-equipment` / `1` |
| `d2-characters.zip` | ~7 MB | **Drawing** the character: layered, animated GIFs plus the rules to stack them | `bnetcc-d2-characters` / `1` |

You can use either on its own. A runnable Python example that downloads both and renders a
character is in [`examples/d2_characters.py`](examples/d2_characters.py).

---

## 1. Download the files once, over BNFTP

BNFTP is a one-shot TCP exchange on the server's normal Battle.net port (**6112**): connect, send
one request, read one reply, and the server closes the connection. One connection per file. No
logon, no CD key.

### Request (everything little-endian)

| Bytes | Field | Value |
|---|---|---|
| 1 | Protocol selector | `0x02` (BNFTP) |
| 2 | Request length | length of everything after the selector, **including this field** |
| 2 | Protocol version | `0x0100` |
| 4 | Platform ID | `IX86` as a little-endian DWORD, so the ASCII bytes `68XI` |
| 4 | Product ID | `STAR` as a little-endian DWORD, so the ASCII bytes `RATS` |
| 4 | Banner ID | `0` |
| 4 | Banner file extension | `0` |
| 4 | Start position | `0` (the whole file) |
| 8 | Local file time | `0` |
| n | File name | ASCII, null-terminated: `d2-equipment.json` or `d2-characters.zip` |

### Reply

| Bytes | Field |
|---|---|
| 2 | Header length (including this field) |
| 2 | Type (ignore) |
| 4 | File size in bytes |
| 4 | Banner ID (ignore) |
| 4 | Banner file extension (ignore) |
| 8 | File time (a Windows FILETIME: 100 ns ticks since 1601) |
| n | File name, null-terminated |
| *size* | The file |

**If the server doesn't have the file, it closes the connection without replying** — treat "closed
before any bytes arrived" as "not available here", not as an error. After sending the file the
server closes the connection; the simple version below reads until then, but reading the header
and then exactly *size* bytes works too.

### Python

```python
import socket, struct

def bnftp_download(host, name, port=6112, timeout=30):
    """Returns (file_time, bytes), or None if the server doesn't serve the file."""
    body = (struct.pack("<H", 0x0100) + b"68XI" + b"RATS" + struct.pack("<IIIQ", 0, 0, 0, 0)
            + name.encode("ascii") + b"\0")
    with socket.create_connection((host, port), timeout=timeout) as s:
        s.sendall(b"\x02" + struct.pack("<H", len(body) + 2) + body)
        data = b""
        while chunk := s.recv(65536):
            data += chunk
    if len(data) < 2:
        return None                                    # closed without a reply: not served here
    header_len = struct.unpack_from("<H", data, 0)[0]
    size = struct.unpack_from("<I", data, 4)[0]
    file_time = struct.unpack_from("<Q", data, 16)[0]
    payload = data[header_len:header_len + size]
    if len(payload) != size:
        raise IOError("transfer cut short")
    return file_time, payload
```

### Please be a good neighbour

- **Ask for these files only from servers that serve them** (bnet.cc). Don't try every server you
  connect to — an older or different server implementation shouldn't have to handle an unfamiliar
  file request from your bot.
- **Download once and keep the copy.** Store the file time from the reply alongside it.
- **Check before trusting.** Refuse anything implausibly large (the art pack is a few MB), confirm
  `format` and `version` before using a file, and treat an unknown `version` as "can't use this
  yet" rather than guessing.
- **Only read zip entries the manifest names.** Never extract the archive to disk by the names
  inside it.

## 2. Optional: check for updates with `SID_GETFILETIME`

On a bot that's already logged on to a server that serves the files, you can ask for a file's time
without downloading it — and re-download only when the server's copy is newer. The server rewrites
the files only when their contents change, so the time is stable.

`SID_GETFILETIME` is BNCS packet **`0x33`** (the usual `FF 33 <WORD length>` header):

| Direction | Layout |
|---|---|
| Client → server | `(DWORD)` request ID (anything; echoed back) · `(DWORD)` unknown, send `0` · `(STRING)` file name |
| Server → client | `(DWORD)` request ID · `(DWORD)` unknown · `(FILETIME)` file time · `(STRING)` file name |

A file time of **0** means the server doesn't have that file. Otherwise, if it's newer than the
time you stored, fetch the file again over BNFTP.

## 3. Read the character from its statstring

A Diablo II **realm** character's chat statstring is:

```
<product><realm>,<character name>,<33-byte portrait>
```

`<product>` is `VD2D` (Diablo II) or `PX2D` (Lord of Destruction). A statstring that's just the
product, with no commas, is an Open Battle.net character: there's nothing to name or draw.

**Keep the statstring as raw bytes.** The portrait is binary; decoding it as UTF-8 will corrupt
it. (If your bot has already made it a string, decode as Latin-1 so each byte stays one character.)

The portrait bytes you need:

| Offset | Meaning |
|---|---|
| 2 – 12 | Graphics value per equipment slot: head, torso, legs, right arm, left arm, right hand, left hand, shield, right shoulder, left shoulder, special. **255 = nothing.** |
| 13 | Class + 1 (1 Amazon, 2 Sorceress, 3 Necromancer, 4 Paladin, 5 Barbarian, 6 Druid, 7 Assassin) |
| 14 – 24 | Tint for each of those slots, same order. **255 = untinted.** |
| 26 | Status flags: `0x04` hardcore, `0x08` dead, `0x20` expansion |

A server that doesn't track items sends 255 in every equipment slot; the character then has nothing
to name and is drawn in their base look.

## 4. Name the gear with `d2-equipment.json`

The equipment bytes aren't item IDs. Each names a **look** that several items share (a Cap, a War
Hat and a Shako look identical), so show every name or the first one.

1. For each entry in `slots`, `value = portrait[slot.offset]`. Skip 255.
2. Find the entry in `slot.values` with that `value`. It has `items` (each with a `name`) — or, for
   the six body-armour slots, a `weight` (`light` / `medium` / `heavy`).
3. Body armour is six slots written together. Take the values of torso, legs, right arm, left arm,
   right shoulder and left shoulder, in that order, and look for a `body_armor.sets` entry whose
   `parts` match: its `items` are the armours with that look. No match → just say the weight.
4. A tint byte `x` (not 255): `colour = (x - 1) & 0x1F`, and `tints.colors[colour]` is its name.
5. When both hands hold the same value (a crossbow fills both, or two matching one-handers), it's
   one look — list it once.

Example: `Wearing: Cap / War Hat / Shako (Crystal Blue), Quilted Armor / Ghost Armor / Dusk Shroud, Eagle Orb / Sacred Globe / …`

`not_drawn.items` lists things that are worn but never drawn (circlets, arrows): a 255 head can
still be a circlet.

## 5. Draw the character with `d2-characters.zip`

The zip holds:

- `manifest.json` — every animation, every part, the lookup tables, and **`rules`**: the exact
  recipe, as a list of sentences. The steps below explain them; the manifest's own `rules` are the
  authority if they ever differ.
- `palette.bin` — 256 RGB triples (768 bytes). Every GIF uses this palette.
- `tints.bin` — 168 × 256 bytes of palette-index remaps, one per (transform, colour).
- `parts/<class>/<part>.gif` — one GIF per body part per look per stance, frames in animation
  order, **palette index 0 = transparent**.

**Decode the GIFs to palette indices, not colours** — the tints work on indices. Watch for image
libraries that convert frames to RGB behind your back: Pillow does it to every frame after the
first unless told otherwise (the example sets `GifImagePlugin.LOADING_STRATEGY` for this), which
leaves the first frame right and the rest garbled.

Everything faces **direction 0**, and one animation tick is `tick_ms` (40 ms).

### Step by step

Let `g[c] = portrait[2 + c]` and `t[c] = portrait[14 + c]` for components `c = 0..10`
(`HD TR LG RA LA RH LH SH S1 S2 S3`); components 11–15 are 255. `class = portrait[13] - 1`
(0–6); `status = portrait[26]`.

1. **Stance.** Hardcore *and* dead (`status & 0x0C == 0x0C`): not in the pack — draw nothing (or
   your own ghost). Hardcore: mode `NU`. Otherwise: mode `TN`.

2. **Hands → weapon class.** `rh = g[5]`, `lh = g[6]`, `sh = g[7]`; `both = rh != 255 and lh != 255`.
   For a hand value `v`, with `s = slots[v]`:
   - its hand class is `s.two_handed` if `both`; for the **right** hand also when `lh == 255`,
     `sh == 255` and `s.two_handed != s.hand`; otherwise `s.hand`.
   - if that class is 13 or 14 and the character isn't an Assassin, or `s.armor` is true, use
     `s.reserved_hand` instead; if it's *still* 13 or 14 and not an Assassin, use 0.
   - a missing hand (255) is 0.

   The weapon class is `w = hand_pairs[right][left]`. **If `w` is 0, draw nothing** (the game draws
   a fallback figure there, which isn't in the pack).

3. **Animation.** `animations[classes[class] + mode + weapon_classes[w]]`, upper case — for
   example `SOTN1HS` for a softcore Sorceress holding a one-handed swinging weapon.

4. **Parts.** For each `[c, L]` in the animation's `layers`: `v = g[c]`, and
   - `code = "lit"` if `v` is 0 or 255, if `c` is 0 (head) and `slots[v].helm` is false, or if
     `slots[v].code` is null; otherwise `code = slots[v].code`.
   - the part is `parts[classes[class] + components[c] + code + mode + L]`, upper case (e.g.
     `SOHDCAPTNHTH`). **If it isn't in `parts`, that component isn't drawn.**
   - each part gives `file`, `frames`, `width`, `height`, and `left` / `top`: where its top-left
     corner goes relative to the character's base point (the feet).

5. **Tints.** For component `c`, `x = t[c]`; no tint if `x` is 0 or 255. Otherwise
   `s = x - 1`, `colour = s & 31`, `transform = s >> 5` (and a transform of 0 means 8). No tint if
   `transform` is 3 or 4 or `colour > 20`. Otherwise the remap table starts at byte
   `((transform - 1) * 21 + colour) * 256` of `tints.bin`: replace every **non-zero** palette index
   `p` in that part with `table[p]`. (A tint byte is `transform * 32 + colour + 1`.)

6. **Draw.** The animation's `sequence` is a list of `[frame f, ticks]`: show frame `f` for
   `ticks × tick_ms`, then the next, looping. For frame `f`, draw the components in `order[f]`,
   **back to front**: part GIF frame `min(f, part.frames - 1)` at `(left, top)` from the base
   point, skipping index 0, colouring through `palette.bin`.

That's the whole renderer; the example script is about 150 lines including the download.
Invigoration's C# version (`D2CharacterPack` in `Invigoration.Core`) produces frames identical to
the example's, pixel for pixel. A drawn character is small — roughly 30–40 × 70–80 pixels once
cropped to the figure — so it sits comfortably next to the classic 78 × 88 chat avatars.

## 6. Checklist

- [ ] Only request the files from servers that serve them; download once; store the file time.
- [ ] Check `format` and `version` (both files), and bound sizes before allocating.
- [ ] Handle the statstring as bytes.
- [ ] "Nothing to draw" is a normal outcome: Open characters, dead hardcore characters, weapon class
      0, a missing animation. Fall back to whatever your bot shows today.
- [ ] Read only the zip entries the manifest names.
- [ ] The art comes from Blizzard's game via the server operator's own install. Serve and cache it
      for your users; don't bundle it into a public repository.
