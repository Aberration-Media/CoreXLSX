//
//  Styles+Standard.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation
import XMLCoder

public extension Styles {
  /// default styles
  static var standard: Styles = {
    // disabled lint rule to allow hard coded XML entries
    // swiftlint:disable line_length
    let stylesXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"2\"><numFmt numFmtId=\"0\" formatCode=\"General\"/><numFmt numFmtId=\"59\" formatCode=\"d&quot;-&quot;mmm&quot;-&quot;yyyy\"/></numFmts><fonts count=\"3\"><font><sz val=\"11\"/><color indexed=\"8\"/><name val=\"ＭＳ Ｐゴシック\"/></font><font><sz val=\"12\"/><color indexed=\"8\"/><name val=\"Helvetica Neue\"/></font><font><sz val=\"14\"/><color indexed=\"8\"/><name val=\"ＭＳ Ｐゴシック\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor indexed=\"9\"/><bgColor auto=\"1\"/></patternFill></fill></fills><borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=\"thin\"><color indexed=\"10\"/></left><right style=\"thin\"><color indexed=\"10\"/></right><top style=\"thin\"><color indexed=\"10\"/></top><bottom style=\"thin\"><color indexed=\"10\"/></bottom><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" applyNumberFormat=\"0\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf></cellStyleXfs><cellXfs count=\"11\"><xf numFmtId=\"0\" fontId=\"0\" applyNumberFormat=\"0\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"0\" fontId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"49\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"top\" wrapText=\"1\"/></xf><xf numFmtId=\"49\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"49\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"59\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"1\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"0\" fontId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"0\" fontId=\"0\" borderId=\"1\" applyNumberFormat=\"0\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"1\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf><xf numFmtId=\"0\" fontId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"1\" applyProtection=\"0\"><alignment vertical=\"bottom\"/></xf></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\"/><colors><indexedColors><rgbColor rgb=\"ff000000\"/><rgbColor rgb=\"ffffffff\"/><rgbColor rgb=\"ffff0000\"/><rgbColor rgb=\"ff00ff00\"/><rgbColor rgb=\"ff0000ff\"/><rgbColor rgb=\"ffffff00\"/><rgbColor rgb=\"ffff00ff\"/><rgbColor rgb=\"ff00ffff\"/><rgbColor rgb=\"ff000000\"/><rgbColor rgb=\"ffffffff\"/><rgbColor rgb=\"ffaaaaaa\"/></indexedColors></colors></styleSheet>"
    // swiftlint:enable line_length

    let decoder = XMLDecoder()
    decoder.trimValueWhitespaces = false
    decoder.shouldProcessNamespaces = true
    do {
      return try decoder.decode(Styles.self, from: stylesXML.data(using: .utf8) ?? Data())

    } catch {
      fatalError("Could not create standard Styles: \(error)")
    }

  }()
} // end extension Styles
