// Renders PDF pages to PNG using macOS PDFKit — no external binaries required.
// usage: swift pdf2png.swift <input.pdf> <outPrefix> <targetWidthPx>
import Foundation
import PDFKit
import AppKit

let args = CommandLine.arguments
guard args.count >= 4, let doc = PDFDocument(url: URL(fileURLWithPath: args[1])) else {
    FileHandle.standardError.write("pdf2png: cannot open PDF\n".data(using: .utf8)!)
    exit(1)
}
let prefix = args[2]
let targetW = CGFloat(Double(args[3]) ?? 1600)

for p in 0..<doc.pageCount {
    guard let page = doc.page(at: p) else { continue }
    let rect = page.bounds(for: .mediaBox)
    let scale = targetW / rect.width
    let size = NSSize(width: (rect.width * scale).rounded(), height: (rect.height * scale).rounded())
    let img = NSImage(size: size)
    img.lockFocus()
    NSColor.white.setFill()
    NSRect(origin: .zero, size: size).fill()
    if let ctx = NSGraphicsContext.current?.cgContext {
        ctx.scaleBy(x: scale, y: scale)
        page.draw(with: .mediaBox, to: ctx)
    }
    img.unlockFocus()
    guard let tiff = img.tiffRepresentation,
          let rep = NSBitmapImageRep(data: tiff),
          let png = rep.representation(using: .png, properties: [:]) else { continue }
    let out = String(format: "%@-%03d.png", prefix, p + 1)
    try? png.write(to: URL(fileURLWithPath: out))
    print(out)
}
