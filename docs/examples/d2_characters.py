#!/usr/bin/env python3
"""
Example: download Command Center's Diablo II character data over BNFTP and draw a character.

    pip install pillow
    python3 d2_characters.py us.bnet.cc out/

Downloads d2-equipment.json and d2-characters.zip once (kept next to this script), then renders a
few sample characters to out/<name>.gif. See ../D2-CHARACTER-DATA-FOR-BOTS.md for the formats.
"""
import io, json, os, socket, struct, sys, zipfile
from PIL import GifImagePlugin, Image

# Pillow turns every GIF frame after the first into RGB(A) by default, which loses the palette
# indices the tint tables work on. All pack GIFs share one palette, so keep them as indices.
GifImagePlugin.LOADING_STRATEGY = GifImagePlugin.LoadingStrategy.RGB_AFTER_DIFFERENT_PALETTE_ONLY


def bnftp_download(host, name, port=6112, timeout=60):
    """Returns (file_time, bytes), or None if the server doesn't serve the file."""
    body = (struct.pack("<H", 0x0100) + b"68XI" + b"RATS" + struct.pack("<IIIQ", 0, 0, 0, 0)
            + name.encode("ascii") + b"\0")
    with socket.create_connection((host, port), timeout=timeout) as s:
        s.sendall(b"\x02" + struct.pack("<H", len(body) + 2) + body)
        data = b""
        while chunk := s.recv(65536):
            data += chunk
    if len(data) < 2:
        return None
    header_len = struct.unpack_from("<H", data, 0)[0]
    size = struct.unpack_from("<I", data, 4)[0]
    if size > 64 * 1024 * 1024:
        raise IOError("implausibly large file")
    file_time = struct.unpack_from("<Q", data, 16)[0]
    payload = data[header_len:header_len + size]
    if len(payload) != size:
        raise IOError("transfer cut short")
    return file_time, payload


def fetch_once(host, name, folder):
    """Downloads a file the first time; afterwards uses the kept copy."""
    path = os.path.join(folder, name)
    if not os.path.exists(path):
        result = bnftp_download(host, name)
        if result is None:
            raise SystemExit(f"{host} doesn't serve {name}")
        file_time, data = result
        with open(path, "wb") as f:
            f.write(data)
        with open(path + ".filetime", "w") as f:
            f.write(str(file_time))
    return path


class CharacterPack:
    def __init__(self, zip_path):
        self.zip = zipfile.ZipFile(zip_path)
        self.m = json.loads(self.zip.read("manifest.json"))
        if self.m.get("format") != "bnetcc-d2-characters" or self.m.get("version") != 1:
            raise ValueError("not a version 1 bnetcc-d2-characters pack")
        pal = self.zip.read("palette.bin")
        self.palette = [tuple(pal[i*3:i*3+3]) for i in range(256)]
        self.tints = self.zip.read("tints.bin")

    def portrait(self, statstring: bytes):
        # <product><realm>,<character>,<33-byte portrait>
        parts = statstring.split(b",", 2)
        if len(parts) < 3 or len(parts[2]) < 33:
            return None
        return parts[2][:33]

    def plan(self, statstring: bytes):
        """Returns (animation, [(frame_order...)]) or None when nothing should be drawn."""
        m = self.m
        p = self.portrait(statstring)
        if p is None:
            return None
        cls = p[13] - 1
        if not 0 <= cls <= 6:
            return None
        status = p[26]
        g = [p[2 + c] for c in range(11)] + [255] * 5
        t = [p[14 + c] for c in range(11)] + [255] * 5

        # Stance
        if status & 4 and status & 8:
            return None
        mode = "NU" if status & 4 else "TN"

        # Hands -> weapon class
        slots = m["slots"]
        def slot(v):
            return slots[v] if 0 <= v < len(slots) else None
        rh, lh, sh = g[5], g[6], g[7]
        both = rh != 255 and lh != 255
        assassin = cls == 6
        def hand_class(v, right):
            s = slot(v)
            if v == 255 or s is None:
                return 0
            if both:
                k = s["two_handed"]
            elif right and lh == 255 and sh == 255 and s["two_handed"] != s["hand"]:
                k = s["two_handed"]
            else:
                k = s["hand"]
            if (k in (13, 14) and not assassin) or s["armor"]:
                k = s["reserved_hand"]
            if k in (13, 14) and not assassin:
                k = 0
            return k
        right, left = hand_class(rh, True), hand_class(lh, False)
        w = m["hand_pairs"][right][left]
        if w == 0:
            return None

        cls_tok = m["classes"][cls]
        anim_key = (cls_tok + mode + m["weapon_classes"][w]).upper()
        anim = m["animations"].get(anim_key)
        if anim is None:
            return None

        layers = {}
        for c, layer_wc in anim["layers"]:
            v = g[c]
            s = slot(v) if v not in (0, 255) else None
            if v in (0, 255) or s is None or s.get("code") is None or (c == 0 and not s["helm"]):
                code = "lit"
            else:
                code = s["code"]
            key = (cls_tok + m["components"][c] + code + mode + layer_wc).upper()
            part = m["parts"].get(key)
            if part is not None:
                layers[c] = (part, t[c])
        return anim_key, anim, layers

    def tint_map(self, x):
        if x in (0, 255):
            return None
        s = x - 1
        colour, transform = s & 31, s >> 5
        if transform == 0:
            transform = 8
        if transform in (3, 4) or colour > 20:
            return None
        base = ((transform - 1) * 21 + colour) * 256
        return self.tints[base:base + 256]

    def frames(self, statstring: bytes):
        planned = self.plan(statstring)
        if planned is None:
            return None
        anim_key, anim, layers = planned
        gifs = {}
        for c, (part, _) in layers.items():
            img = Image.open(io.BytesIO(self.zip.read(part["file"])))
            frames = []
            for i in range(part["frames"]):
                img.seek(min(i, getattr(img, "n_frames", 1) - 1))
                if img.mode != "P":
                    raise ValueError(f"{part['file']} frame {i} isn't palette indices")
                frames.append(img.copy())  # palette indices, 0 = transparent
            gifs[c] = frames
        # Canvas big enough for every part around the base point
        xs = [pt["left"] for pt, _ in layers.values()] + [pt["left"] + pt["width"] for pt, _ in layers.values()]
        ys = [pt["top"] for pt, _ in layers.values()] + [pt["top"] + pt["height"] for pt, _ in layers.values()]
        ox, oy = -min(xs), -min(ys)
        W, H = max(xs) + ox, max(ys) + oy
        out = []
        for f, ticks in anim["sequence"]:
            canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            for c in anim["order"][f]:
                if c not in layers:
                    continue
                part, tint = layers[c]
                src = gifs[c][min(f, part["frames"] - 1)]
                idx = src.tobytes()
                tmap = self.tint_map(tint)
                rgba = bytearray()
                for px in idx:
                    if px == 0:
                        rgba += b"\x00\x00\x00\x00"
                    else:
                        q = tmap[px] if tmap else px
                        r, gg, b = self.palette[q]
                        rgba += bytes((r, gg, b, 255))
                layer_img = Image.frombytes("RGBA", src.size, bytes(rgba))
                canvas.alpha_composite(layer_img, (part["left"] + ox, part["top"] + oy))
            out.append((canvas, ticks * self.m["tick_ms"]))
        return anim_key, out

def realm_statstring(cls_raw, gear, tints=None, status=0xA0, act=0x9E):
    p = bytearray([0xFF] * 33)
    p[0], p[1], p[13], p[25], p[26], p[27] = 0x84, 0x80, cls_raw, 80, status, act
    for i in range(11):
        p[2 + i] = gear[i]
        p[14 + i] = (tints or [0xFF] * 11)[i]
    return b"PX2DUSEast,Kilua," + bytes(p)

if __name__ == "__main__":
    host = sys.argv[1] if len(sys.argv) > 1 else "us.bnet.cc"
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "out"
    here = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(out_dir, exist_ok=True)

    fetch_once(host, "d2-equipment.json", here)
    pack = CharacterPack(fetch_once(host, "d2-characters.zip", here))

    samples = {
        # Sorceress: Cap/War Hat/Shako look, light armour with medium shoulders, an orb, crystal-blue helm
        "sorceress": realm_statstring(2, [57, 1, 1, 1, 1, 0x33, 0xFF, 0xFF, 2, 2, 0xFF], [7] + [0xFF] * 10),
        # Paladin: heavy armour, a sword and a shield
        "paladin": realm_statstring(4, [59, 3, 3, 3, 3, 17, 0xFF, 81, 3, 3, 0xFF]),
        # Amazon with a bow in the left hand
        "amazon_bow": realm_statstring(1, [0xFF, 2, 2, 2, 2, 0xFF, 41, 0xFF, 2, 2, 0xFF]),
        # Hardcore Barbarian wearing nothing (hardcore stands in the NU stance)
        "barbarian_hc": realm_statstring(5, [0xFF] * 11, status=0xA4),
    }
    for name, statstring in samples.items():
        result = pack.frames(statstring)
        if result is None:
            print(f"{name}: nothing to draw")
            continue
        animation, frames = result
        frames[0][0].save(os.path.join(out_dir, f"{name}.gif"), save_all=True,
                          append_images=[frame for frame, _ in frames[1:]],
                          duration=[ms for _, ms in frames], loop=0, disposal=2)
        print(f"{name}: {animation}, {len(frames)} frames -> {out_dir}/{name}.gif")
