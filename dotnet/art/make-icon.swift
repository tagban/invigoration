// Renders Invigoration's macOS app icon at 1024px: the 64x64 pirate flag from Assets/flag.ico,
// scaled up in whole multiples (nearest-neighbour, so every pixel keeps its colour) and lightly
// softened, on a wood-plank rounded square. Everything but the flag is drawn here.
//
// Why a backing at all: macOS 26 no longer shows a free-form icon as-is. One that doesn't fill
// the standard rounded-square shape gets shrunk onto a light plate of the system's own, which is
// what the bare transparent flag used to get. Filling Apple's icon grid (an 824px body inset
// 100px in the 1024 canvas, circular 185.4px corners) with our own wood means the system masks
// ours instead of adding its plate, and it still looks right on macOS 11-15. (A home-made
// superellipse did get plated at the large sizes, so the corners stay plain circular arcs.)
//
// Run build-icon.sh rather than this directly; it makes every icns size from the one render.
// usage: make-icon <flag.png> <out.png> [inset=100] [cornerRadius=185.4] [flagScale=10] [soften=2.2]
import AppKit
import CoreGraphics
import CoreImage
import Foundation

let args = CommandLine.arguments
let flagPath = args[1]
let outPath = args[2]
let inset = args.count > 3 ? Double(args[3])! : 100
let corner = args.count > 4 ? Double(args[4])! : 185.4 // Apple's grid; macOS 26 is picky about the corner shape it will mask
let flagScale = args.count > 5 ? Int(args[5])! : 10
let soften = args.count > 6 ? Double(args[6])! : 2.2 // blur applied to the scaled-up flag; 0 = razor-crisp pixels

let size = 1024
let W = size, H = size

// --- value noise / fbm -----------------------------------------------------------------------
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
func mix(_ a: Double, _ b: Double, _ t: Double) -> Double { a + (b - a) * t }
func clamp(_ v: Double, _ lo: Double = 0, _ hi: Double = 1) -> Double { min(hi, max(lo, v)) }

// --- the planks --------------------------------------------------------------------------------
// Horizontal planks across the body, like a ship's hull or a tavern sign.
let bodyX0 = inset, bodyY0 = inset, bodyW = Double(size) - 2 * inset, bodyH = Double(size) - 2 * inset
let plankCount = 4
let plankH = bodyH / Double(plankCount)
// Warm weathered oak: light and dark grain colors, varied a little per plank.
let light: (Double, Double, Double) = (0.62, 0.43, 0.26)
let dark: (Double, Double, Double) = (0.36, 0.22, 0.12)
let plankTint: [Double] = [1.00, 0.93, 1.05, 0.96]
let plankShift: [Double] = [0, 137, 59, 211] // seeds/offsets so no two planks share a grain
// A couple of knots: (plank, x within body 0..1, radius px)
let knots: [(Int, Double, Double)] = [(0, 0.78, 20), (2, 0.18, 26), (3, 0.64, 14)]

var pixels = [UInt8](repeating: 0, count: W * H * 4)
for py in 0..<H {
    for px in 0..<W {
        let x = Double(px) + 0.5, y = Double(py) + 0.5
        // Everything outside the body stays transparent; the rounded mask is applied later.
        guard x >= bodyX0, x < bodyX0 + bodyW, y >= bodyY0, y < bodyY0 + bodyH else { continue }
        let bx = x - bodyX0, by = y - bodyY0
        let plank = min(plankCount - 1, Int(by / plankH))
        let ly = by - Double(plank) * plankH // 0..plankH within this plank
        let seed = Int(plankShift[plank])

        // Grain: rings stretched along the plank, bent by low-frequency noise.
        var gx = bx, gy = ly
        for (kp, kx, kr) in knots where kp == plank {
            // Pull the grain around a knot so the rings flow round it.
            let cx = kx * bodyW, cy = plankH * 0.5
            let dx = bx - cx, dy = (ly - cy) * 2.2
            let d = sqrt(dx * dx + dy * dy)
            let influence = exp(-d / (kr * 3.2))
            gy += influence * (ly - cy) * 0.9
            gx += influence * 8
        }
        let warp = fbm(gx / 260 + plankShift[plank], gy / 55, seed) * 38
        let ringPos = (gy + warp + fbm(gx / 90, gy / 18, seed + 7) * 6) / 7.5
        var ring = ringPos - floor(ringPos)
        ring = pow(abs(sin(ring * .pi)), 5) // thin dark lines
        let fiber = fbm(gx / 6, gy / 1.2, seed + 3, octaves: 3) // fine fibres along the plank
        let broad = fbm(gx / 140, gy / 40, seed + 11, octaves: 3) // broad tonal variation
        var t = clamp(0.25 + ring * 0.55 + (fiber - 0.5) * 0.35 + (broad - 0.5) * 0.6)

        // Knot centres: darker concentric rings.
        for (kp, kx, kr) in knots where kp == plank {
            let cx = kx * bodyW, cy = plankH * 0.5
            let dx = bx - cx, dy = (ly - cy) * 1.6
            let d = sqrt(dx * dx + dy * dy)
            if d < kr * 1.6 {
                let k = 1 - d / (kr * 1.6)
                t = clamp(t + k * 0.55 + pow(abs(sin(d / 3.2)), 6) * k * 0.4)
            }
        }

        let tint = plankTint[plank]
        var r = mix(light.0, dark.0, t) * tint
        var g = mix(light.1, dark.1, t) * tint
        var b = mix(light.2, dark.2, t) * tint

        // Bevel: a lit top edge and a shaded bottom edge on each plank, and a dark seam between.
        let edgeTop = ly, edgeBottom = plankH - ly
        if edgeTop < 7 { let k = 1 - edgeTop / 7; r += k * 0.10; g += k * 0.08; b += k * 0.05 }
        if edgeBottom < 9 { let k = 1 - edgeBottom / 9; r *= 1 - k * 0.45; g *= 1 - k * 0.45; b *= 1 - k * 0.45 }
        if (plank > 0 && edgeTop < 2.5) || (plank < plankCount - 1 && edgeBottom < 2.5) {
            r *= 0.25; g *= 0.22; b *= 0.2
        }

        // Nails near each end of every plank — far enough in that macOS 26's own rounded mask
        // (slightly tighter than this shape) doesn't clip the ones by the corners.
        for nx in [0.095, 0.905] {
            let cx = nx * bodyW, cy = plankH * 0.5
            let dx = bx - cx, dy = ly - cy
            let d = sqrt(dx * dx + dy * dy)
            if d < 10 {
                let shade = 0.30 + 0.35 * clamp(1 - (dx + dy + 10) / 20) // lit from the top left
                r = shade; g = shade * 0.97; b = shade * 0.93
                if d > 8 { r *= 0.6; g *= 0.6; b *= 0.6 }
            } else if d < 13 {
                let k = 1 - (d - 10) / 3
                r *= 1 - k * 0.35; g *= 1 - k * 0.35; b *= 1 - k * 0.35
            }
        }

        // Soft vignette and a little light from the top left, for depth.
        let ux = bx / bodyW - 0.5, uy = by / bodyH - 0.5
        let vignette = 1 - clamp((ux * ux + uy * uy) * 1.1) * 0.55
        let lightFall = 1.08 - (bx / bodyW + by / bodyH) * 0.08
        r *= vignette * lightFall; g *= vignette * lightFall; b *= vignette * lightFall

        let i = (py * W + px) * 4
        pixels[i] = UInt8(clamp(r) * 255)
        pixels[i + 1] = UInt8(clamp(g) * 255)
        pixels[i + 2] = UInt8(clamp(b) * 255)
        pixels[i + 3] = 255
    }
}

let colorSpace = CGColorSpace(name: CGColorSpace.sRGB)!
let woodContext = CGContext(data: &pixels, width: W, height: H, bitsPerComponent: 8, bytesPerRow: W * 4,
                            space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
let wood = woodContext.makeImage()!

// --- compose -----------------------------------------------------------------------------------
let ctx = CGContext(data: nil, width: W, height: H, bitsPerComponent: 8, bytesPerRow: W * 4,
                    space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
let body = CGRect(x: bodyX0, y: bodyY0, width: bodyW, height: bodyH)
let shape = CGPath(roundedRect: body, cornerWidth: corner, cornerHeight: corner, transform: nil)

// The icon's own drop shadow (Apple's grid leaves room for it below the body).
if inset > 0 {
    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -10), blur: 28, color: NSColor(white: 0, alpha: 0.5).cgColor)
    ctx.addPath(shape); ctx.setFillColor(NSColor(white: 0.2, alpha: 1).cgColor); ctx.fillPath()
    ctx.restoreGState()
}

ctx.saveGState()
ctx.addPath(shape); ctx.clip()
ctx.draw(wood, in: CGRect(x: 0, y: 0, width: W, height: H))

// The flag, crisp: nearest-neighbour at a whole multiple, centred, with a soft shadow on the wood.
let flagImage = NSImage(contentsOfFile: flagPath)!
var flagRect = CGRect(x: 0, y: 0, width: 64, height: 64)
let flag = flagImage.cgImage(forProposedRect: &flagRect, context: nil, hints: nil)!
let fw = Double(flag.width * flagScale), fh = Double(flag.height * flagScale)
let flagDest = CGRect(x: (Double(W) - fw) / 2, y: (Double(H) - fh) / 2 - 6, width: fw, height: fh)
// Scale up nearest-neighbour first (keeps every pixel's colour true), then optionally take the
// staircase off the edges with a small blur — at icon sizes the pixels still read, just not jagged.
let big = CGContext(data: nil, width: Int(fw), height: Int(fh), bitsPerComponent: 8, bytesPerRow: Int(fw) * 4,
                    space: colorSpace, bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
big.interpolationQuality = .none
big.draw(flag, in: CGRect(x: 0, y: 0, width: fw, height: fh))
var scaledFlag = big.makeImage()!
if soften > 0 {
    let ci = CIImage(cgImage: scaledFlag).clampedToExtent()
    let blurred = ci.applyingGaussianBlur(sigma: soften).cropped(to: CGRect(x: 0, y: 0, width: fw, height: fh))
    scaledFlag = CIContext().createCGImage(blurred, from: CGRect(x: 0, y: 0, width: fw, height: fh))!
}
ctx.setShadow(offset: CGSize(width: 6, height: -12), blur: 22, color: NSColor(white: 0, alpha: 0.65).cgColor)
ctx.draw(scaledFlag, in: flagDest)
ctx.restoreGState()

// A thin inner rim so the edge reads as a board's edge rather than a cut-off texture.
ctx.saveGState()
ctx.addPath(shape); ctx.clip()
ctx.addPath(shape)
ctx.setStrokeColor(NSColor(calibratedRed: 0.12, green: 0.07, blue: 0.03, alpha: 0.55).cgColor)
ctx.setLineWidth(10)
ctx.strokePath()
ctx.restoreGState()

let out = ctx.makeImage()!
let rep = NSBitmapImageRep(cgImage: out)
try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: outPath))
print("wrote \(outPath)")
