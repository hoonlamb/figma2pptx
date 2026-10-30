// Advance widths in the font PowerPoint will actually use.
//
// Figma stores each text node's width as measured in whatever font the design
// used. When the deck renders in a different face — or the same family with
// different shaping — those widths are wrong, and anything sized or centred
// from them lands off. This measures the real thing with CoreText.
//
//   echo '[{"t":"공통사항","f":"Pretendard","b":true,"s":16,"k":-0.8}]' \
//     | swift textwidth.swift
//   -> [52.9]                      // widths in the same px unit as `s`
import Foundation
import AppKit

struct Item: Decodable {
    let t: String            // the string
    let f: String            // family, as written in the pptx
    let b: Bool?             // bold attribute
    let i: Bool?             // italic
    let s: Double            // font size, px
    let k: Double?           // letter spacing, px
}

let data = FileHandle.standardInput.readDataToEndOfFile()
guard let items = try? JSONDecoder().decode([Item].self, from: data) else {
    FileHandle.standardError.write("textwidth: bad input\n".data(using: .utf8)!)
    exit(1)
}

var cache: [String: NSFont] = [:]
func font(_ family: String, _ size: Double, _ bold: Bool, _ italic: Bool) -> NSFont {
    let key = "\(family)|\(size)|\(bold)|\(italic)"
    if let f = cache[key] { return f }
    // PowerPoint addresses weights by face name ("Pretendard Medium"); the bold
    // attribute is a separate switch. Mirror both, then fall back gracefully.
    var f = NSFont(name: family, size: CGFloat(size))
    if f == nil { f = NSFont(name: family.replacingOccurrences(of: " ", with: "-"), size: CGFloat(size)) }
    if f == nil { f = NSFont.systemFont(ofSize: CGFloat(size)) }
    var out = f!
    var traits: NSFontTraitMask = []
    if bold { traits.insert(.boldFontMask) }
    if italic { traits.insert(.italicFontMask) }
    if !traits.isEmpty {
        if let t = NSFontManager.shared.font(withFamily: out.familyName ?? family,
                                             traits: traits, weight: bold ? 9 : 5,
                                             size: CGFloat(size)) { out = t }
    }
    cache[key] = out
    return out
}

var widths: [Double] = []
for it in items {
    if it.t.isEmpty { widths.append(0); continue }
    let fnt = font(it.f, it.s, it.b ?? false, it.i ?? false)
    var attrs: [NSAttributedString.Key: Any] = [.font: fnt]
    if let k = it.k, k != 0 { attrs[.kern] = CGFloat(k) }
    let s = NSMutableAttributedString(string: it.t, attributes: attrs)
    // kern is applied after every glyph; drop it on the last one so the width
    // matches the visible run rather than trailing into empty space
    if it.k != nil && it.k != 0 {
        s.removeAttribute(.kern, range: NSRange(location: it.t.utf16.count - 1, length: 1))
    }
    widths.append(Double(s.size().width))
}

let out = try! JSONSerialization.data(withJSONObject: widths.map { round($0 * 100) / 100 })
FileHandle.standardOutput.write(out)
