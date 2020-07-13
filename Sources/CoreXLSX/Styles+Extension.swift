//
//  Styles+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension Styles: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs", "colors", "cellStyles", "tableStyles"]
  }
}

extension Fills: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["fill"]
  }
}

extension NumberFormats: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["numFmt"]
  }
}

extension NumberFormat: AttributeType {}
extension Fonts: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["font"]
  }
}

extension Borders: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["border"]
  }
}

extension Font.Size: AttributeType {}
extension Font.Name: AttributeType {}
extension Font.Bold: AttributeType {}
extension Font.Italic: AttributeType {}
extension Font.Strike: AttributeType {}
extension Color: AttributeType {}
extension PatternFill: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["fgColor", "bgColor"]
  }
}

extension Border.Value: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["color"]
  }
}

extension Format: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["alignment"]
  }
}

extension Format.Alignment: AttributeType {}
extension DifferentialFormats: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["dxf"]
  }
}

extension CellStyles: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["cellStyle"]
  }
}

extension CellStyle: AttributeType {}
extension CellStyleFormats: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["xf"]
  }
}

extension CellFormats: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["xf"]
  }
}

extension TableStyles: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["tableStyle"]
  }
}
