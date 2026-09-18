// Renders the Warcraft III theme's wood grain: the same grain as the app icon's planks
// (make-icon.swift — the noise and ring shaping are copied from there so the two match), as a
// light/dark OVERLAY rather than a coloured image. The theme lays it over its own band colour, so a
// custom theme that changes that colour keeps its colour and still gets the grain.
//
// One plank-length strip, grain running along it, with a butt joint at each tile edge (x = 0) and
// one part-way along, so tiled end to end it reads as boards laid in a row and needs no seamless
// blend. Written at 2x for Retina; the XAML sizes it in half its pixels.
//
// usage: make-wood-grain <out-horizontal.png> <out-vertical.png>
import AppKit
import CoreGraphics
import Foundation

let args = CommandLine.arguments
let outH = args[1], outV = args[2]

let scale = 2.0              // pixels per DIP
let W = Int(512 * scale)     // 512 DIP of plank before the pattern repeats
let H = Int(32 * scale)      // tall enough for the 24 DIP name sign without tiling vertically
let jointAt = 0.58           // the joint part-way along, as a fraction of the width

// --- value noise / fbm (as in make-icon.swift) --------------------------------------------------
func hash(_ x: Int, _ y: Int, _ seed: Int) -> Double {
    var h = UInt64(bitPattern: Int64(x &* 374761393 &+ y &* 668265263 &+ seed &* 1442695040888963407))
    h = (h ^ (h >> 13)) &* 1274126177
    h = h ^ (h >> 16)
    return Double(h & 0xFFFF) / 65535.0
}
func smooth(_ t: Double) -> Double { t * t * (3 - 2 * t) }
func noise(_ x: Double, _ y: Double, _ seed: Int) -> Double {
    let xi = Int(floor(x)), yi = Int(floor(y))
    let xf = x - floor(x), yf = y - floor(y)
    let a = hash(xi, yi, seed), b = hash(xi + 1, yi, seed)
    let c = hash(xi, yi + 1, seed), d = hash(xi + 1, yi + 1, seed)
    let u = smooth(xf), v = smooth(yf)
    return (a * (1 - u) + b * u) * (1 - v) + (c * (1 - u) + d * u) * v
}
func fbm(_ x: Double, _ y: Double, _ seed: Int, octaves: Int = 5) -> Double {
    var sum = 0.0, amp = 0.5, fx = x, fy = y
    for o in 0..<octaves {
        sum += amp * noise(fx, fy, seed + o * 31)
        fx *= 2.03; fy *= 2.03; amp *= 0.5
    }
    return sum
}
func clamp(_ v: Double, _ lo: Double = 0, _ hi: Double = 1) -> Double { min(hi, max(lo, v)) }

// How dark the icon's wood is at a point, 0 (lightest) to 1 (darkest) — the icon's own formula,
// in DIPs so the grain is the size it looks on the icon.
func woodDarkness(_ x: Double, _ y: Double, seed: Int, shift: Double) -> Double {
    let warp = fbm(x / 260 + shift, y / 55, seed) * 38
    let ringPos = (y + warp + fbm(x / 90, y / 18, seed + 7) * 6) / 7.5
    var ring = ringPos - floor(ringPos)
    ring = pow(abs(sin(ring * .pi)), 5)
    let fiber = fbm(x / 6, y / 1.2, seed + 3, octaves: 3)
    let broad = fbm(x / 140, y / 40, seed + 11, octaves: 3)
    return clamp(0.25 + ring * 0.55 + (fiber - 0.5) * 0.35 + (broad - 0.5) * 0.6)
}

// The icon's two wood colours; the overlay is how far each point sits above or below their middle.
func luminance(_ c: (Double, Double, Double)) -> Double { 0.2126 * c.0 + 0.7152 * c.1 + 0.0722 * c.2 }
let light: (Double, Double, Double) = (0.62, 0.43, 0.26)
let dark: (Double, Double, Double) = (0.36, 0.22, 0.12)
let midLum = luminance((light.0 * 0.55 + dark.0 * 0.45, light.1 * 0.55 + dark.1 * 0.45, light.2 * 0.55 + dark.2 * 0.45))

var pixels = [UInt8](repeating: 0, count: W * H * 4)
let jointX = Double(W) * jointAt
for py in 0..<H {
    for px in 0..<W {
        let x = Double(px) / scale, y = Double(py) / scale // DIPs
        // Each board between joints has its own grain.
        let second = Double(px) >= jointX
        let d = woodDarkness(x + (second ? 1000 : 0), y + 11, seed: second ? 137 : 59, shift: second ? 137 : 59)
        let lum = luminance((light.0 + (dark.0 - light.0) * d, light.1 + (dark.1 - light.1) * d, light.2 + (dark.2 - light.2) * d))

        // Darker than the middle → black at some opacity; lighter → white.
        var r = 0.0, a = 0.0
        if lum < midLum {
            a = min(0.62, 1 - lum / midLum)
        } else {
            r = 1
            a = min(0.30, (lum - midLum) / (1 - midLum))
        }

        // Butt joints: a dark gap with a lit edge after it (light from the left), at the tile edge and part-way along.
        for jx in [0.0, jointX, Double(W)] {
            let dx = Double(px) + 0.5 - jx
            if abs(dx) < 1.5 * scale { r = 0; a = 0.85 }
            else if dx >= 1.5 * scale && dx < 2.5 * scale { r = 1; a = 0.22 }
        }

        let i = (py * W + px) * 4
        let alpha = UInt8(clamp(a) * 255)
        // Premultiplied for the bitmap context below.
        let v = UInt8(clamp(r * a) * 255)
        pixels[i] = v; pixels[i + 1] = v; pixels[i + 2] = v; pixels[i + 3] = alpha
    }
}

let colorSpace = CGColorSpace(name: CGColorSpace.sRGB)!
func write(_ image: CGImage, _ path: String) {
    let rep = NSBitmapImageRep(cgImage: image)
    rep.size = NSSize(width: Double(image.width) / scale, height: Double(image.height) / scale) // 144 dpi: Avalonia sizes it in DIPs
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: path))
    print("wrote \(path) (\(image.width)x\(image.height))")
}

let ctx = CGContext(data: &pixels, width: W, height: H, bitsPerComponent: 8, bytesPerRow: W * 4,
                    space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
let horizontal = ctx.makeImage()!
write(horizontal, outH)

// The same strip turned upright for the frame's sides, so their grain runs along them too.
let vctx = CGContext(data: nil, width: H, height: W, bitsPerComponent: 8, bytesPerRow: H * 4,
                     space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
vctx.translateBy(x: CGFloat(H), y: 0)
vctx.rotate(by: .pi / 2)
vctx.draw(horizontal, in: CGRect(x: 0, y: 0, width: W, height: H))
write(vctx.makeImage()!, outV)
