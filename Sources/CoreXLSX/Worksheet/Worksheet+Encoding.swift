//
//  Worksheet+Encoding.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension Worksheet.Dimension: AttributeType {}
extension Worksheet.FormatProperties: AttributeType {}
// extension Worksheet.Properties: AttributeType {}
extension Pane: AttributeType {}
extension PageSetUpProperties: AttributeType {}
extension Column: AttributeType {}
extension Row: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["c"]
  }
}

extension Cell: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["v", "f"]
  }
}

extension SheetView: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["pane"]
  }
}

public extension Worksheet.Data {

  /// computed list of rows
  var unorderedRows: [Row] { [Row](rowsByReference.values) }

  //rows listed in order of reference
  var rows: [Row] { self.unorderedRows.sorted(by: { $0.reference < $1.reference }) }

  init(rows: [Row]) {
    self.rowsByReference = rows.dictionary { return ($1.reference, $1) }
  } //end constructor()

  init(from decoder: Decoder) throws {
    let values = try decoder.container(keyedBy: CodingKeys.self)
    let decodedRows = try values.decode([Row].self, forKey: .rows)
    self.rowsByReference = decodedRows.dictionary { return ($1.reference, $1) }
  } //end decoder()

  func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: CodingKeys.self)
      try container.encode(rows, forKey: .rows)
  } //end encode()

} //end extension Worksheet.Data

public extension Row {

  /// computed list of cells
  var unorderedCells: [Cell] { [Cell](cellsByReference.values) }

  //cells listed in order of reference
  var cells: [Cell] { self.unorderedCells.sorted(by: { $0.reference.column < $1.reference.column }) }

  init(reference: UInt, height: Double?, customHeight: String?, cells: [Cell]) {
    self.reference = reference
    self.height = height
    self.customHeight = customHeight
    self.cellsByReference = cells.dictionary({ return ($1.reference.column, $1) })
  } //end constructor()

  init(from decoder: Decoder) throws {
    let values = try decoder.container(keyedBy: CodingKeys.self)
    let decodedCells = try values.decode([Cell].self, forKey: .cells)
    self.cellsByReference = decodedCells.dictionary({ return ($1.reference.column, $1) })
    self.reference = try values.decode(UInt.self, forKey: .reference)
    self.height = try values.decodeIfPresent(Double.self, forKey: .height)
    self.customHeight = try values.decodeIfPresent(String.self, forKey: .customHeight)
  } //end decoder()

  func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: CodingKeys.self)
      try container.encode(cells, forKey: .cells)
      try container.encode(reference, forKey: .reference)
      try container.encode(height, forKey: .height)
      try container.encode(customHeight, forKey: .customHeight)
  } //end encode()

} //end extension Row
