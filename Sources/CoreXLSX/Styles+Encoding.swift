//
//  Styles+Encoding.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

// enable model attributes to be encoded correctly

// MARK: - Styles

extension Styles: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs", "colors", "cellStyles", "tableStyles"]
  }
}

// MARK: - Fills

extension Fills: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["fill"]
  }
}

// MARK: - NumberFormats

extension NumberFormats: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["numFmt"]
  }
}

// MARK: - NumberFormat

extension NumberFormat: AttributeType {}

// MARK: - Fonts

extension Fonts: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["font"]
  }
}

// MARK: - Borders

extension Borders: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["border"]
  }
}

// MARK: - Font.Size

extension Font.Size: AttributeType {}

// MARK: - Font.Name

extension Font.Name: AttributeType {}

// MARK: - Font.Bold

extension Font.Bold: AttributeType {}

// MARK: - Font.Italic

extension Font.Italic: AttributeType {}

// MARK: - Font.Strike

extension Font.Strike: AttributeType {}

// MARK: - Color

extension Color: AttributeType {}

// MARK: - PatternFill

extension PatternFill: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["fgColor", "bgColor"]
  }
}

// MARK: - Border.Value

extension Border.Value: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["color"]
  }
}

// MARK: - Format

extension Format: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["alignment"]
  }

  public func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: CodingKeys.self)
      try container.encodeIfPresent(numberFormatId, forKey: .numberFormatId)
      try container.encodeIfPresent(borderId, forKey: .borderId)
      try container.encodeIfPresent(fillId, forKey: .fillId)
      try container.encodeIfPresent(fontId, forKey: .fontId)
      try container.encodeIfPresent(BoolConverter(applyNumberFormat).intValue, forKey: .applyNumberFormat)
      try container.encodeIfPresent(BoolConverter(applyFont).intValue, forKey: .applyFont)
      try container.encodeIfPresent(BoolConverter(applyFill).intValue, forKey: .applyFill)
      try container.encodeIfPresent(BoolConverter(applyBorder).intValue, forKey: .applyBorder)
      try container.encodeIfPresent(BoolConverter(applyAlignment).intValue, forKey: .applyAlignment)
      try container.encodeIfPresent(BoolConverter(applyProtection).intValue, forKey: .applyProtection)
      try container.encodeIfPresent(alignment, forKey: .alignment)
  } //end encode()

} //end extension Format

// MARK: - Format.Alignment

extension Format.Alignment: AttributeType {

  enum CodingKeys: String, CodingKey {
    case vertical
    case horizontal
    case wrapText
  }

  public func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: CodingKeys.self)
      try container.encodeIfPresent(vertical, forKey: .vertical)
      try container.encodeIfPresent(horizontal, forKey: .horizontal)
      try container.encodeIfPresent(BoolConverter(wrapText).intValue, forKey: .wrapText)
  } //end encode()
}

// MARK: - DifferentialFormats

extension DifferentialFormats: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["dxf"]
  }
}

// MARK: - CellStyles

extension CellStyles: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["cellStyle"]
  }
}

// MARK: - CellStyle

extension CellStyle: AttributeType {}

// MARK: - CellStyleFormats

extension CellStyleFormats: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["xf"]
  }
}

// MARK: - CellFormats

extension CellFormats: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["xf"]
  }
}

// MARK: - TableStyles

extension TableStyles: ExcludeAttributeType {

  /// list of keys to encode as a sub element
  public static var nonAttributeKeys: [String]? {
    return ["tableStyle"]
  }
}
