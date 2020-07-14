//
//  Worksheet+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

public extension Worksheet {
  /// create empty worksheet
  init() {
    properties = nil
    formatProperties = nil
    dimension = nil
    sheetViews = nil
    columns = nil
    data = nil
    mergeCells = nil
  }
} // end extension Worksheet

public extension Worksheet.Data {

  /// computed list of rows
  var rows: [Row] { [Row](rowsByReference.values) }

  //rows listed in order of reference
  var orderedRows: [Row] { self.rows.sorted(by: { $0.reference < $1.reference }) }

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
      try container.encode(orderedRows, forKey: .rows)
  } //end encode()

} //end extension Worksheet.Data

public extension Row {

  /// computed list of cells
  var cells: [Cell] { [Cell](cellsByReference.values) }

  //cells listed in order of reference
  var orderedCells: [Cell] { self.cells.sorted(by: { $0.reference.column < $1.reference.column }) }

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
      try container.encode(orderedCells, forKey: .cells)
      try container.encode(reference, forKey: .reference)
      try container.encode(height, forKey: .height)
      try container.encode(customHeight, forKey: .customHeight)
  } //end encode()

} //end extension Row

public extension Worksheet {

  // MARK: Convenience Properties

  /// number of rows in worksheet
  var numberOfRows: Int {
    if let last = self.lastRow {
      return Int(last.reference + 1)
    }
    return 0
  }

  /// the first row in the spread sheet
  var firstRow: Row? { self.data?.rows.sorted(by: { $0.reference < $1.reference }).first}

  /// the last row in the spread sheet
  var lastRow: Row? { self.data?.rows.sorted(by: { $0.reference < $1.reference }).last }

  // MARK: - Editing Functions

  mutating func updateCell(
    at rowIndex: UInt,
    column: ColumnReference,
    with value: String,
    sharedStrings: inout SharedStrings
  ) {

    //valid row index
    if self.data?.rowsByReference[rowIndex] != nil {
      //update existing cell (updating in place as Cells are passed by value not reference)
      if self.data?.rowsByReference[rowIndex]?.cellsByReference[column] != nil {
        self.data?.rowsByReference[rowIndex]?.cellsByReference[column]?.value = value
      }
      //create new cell
      else {
//        //create cell
//        let cellReference = CellReference(columnReference, rowReference)
//        let cell = Cell(reference: cellReference, type: .sharedString, s: style, inlineString: nil, formula: nil, value: String(valueIndex))
      }
    }
    //invalid row index
    else {
      print("No row found at index: \(rowIndex)")
    }

  } //end updateCell()

  mutating func updateRow(
    at rowIndex: UInt,
    from column: Int = 1,
    with values: [String],
    sharedStrings: inout SharedStrings
  ) {

    //valid row index
    if self.data?.rowsByReference[rowIndex] != nil {

      //update cells
      let firstColumn: Int = column < 1 ? 1 : column
      for (index, value) in values.enumerated() {
        if let columnReference = ColumnReference(index + firstColumn) {
          self.updateCell(at: rowIndex, column: columnReference, with: value, sharedStrings: &sharedStrings)
        }
      } //end for (values)

    }
    //invalid row index
    else {
      print("No row found at index: \(rowIndex)")
    }

  } //end updateRow()

  /**
    Add a row at the end of the worksheet with the specified values
   */
  mutating func addRow(
    with values: [String] = [],
    sharedStrings: inout SharedStrings,
    height: Double? = nil,
    styles: [String?]? = nil
  ) {

    //insert row at end of worksheet
    let rowIndex: UInt = UInt(self.numberOfRows)
    self.insertRow(at: rowIndex, with: values, sharedStrings: &sharedStrings, height: height, styles: styles)

  } //end addRow()

  /**
    Insert a row at the specified row index in the worksheet with the specified values
   */
  mutating func insertRow(
    at rowIndex: UInt,
    with values: [String] = [],
    sharedStrings: inout SharedStrings,
    height: Double? = nil,
    styles: [String?]? = nil
  ) {

    //add string values and create cells
    var cells: [Cell] = []
    for (index, value) in values.enumerated() {
      if let columnReference = ColumnReference(index + 1) {

        //add shared string
        let valueIndex: Int = sharedStrings.addString(value)

        //get cell style
        var style: String?
        if let styleReferences = styles, index < styleReferences.count {
          style = styleReferences[index]
        }

        //create cell
        let cellReference = CellReference(columnReference, rowIndex)
        let cell = Cell(reference: cellReference, type: .sharedString, s: style, inlineString: nil, formula: nil, value: String(valueIndex))
        cells.append(cell)

      }
    } //end for (values)

    //create new row
    let row = Row(reference: rowIndex, height: height, customHeight: nil, cells: cells)
    self.data?.rowsByReference[rowIndex] = row

  } //end insertRow()

  /**
    Retrieve the styles for the specified row index
   */
  func stylesForRow(_ index: Int, range: Range<Int>? = nil) -> [String?] {
    var styles: [String?] = []

    //found rows
    if let rows = self.data?.orderedRows {

      //valid row index
      if index >= 0 && index < rows.count {

        //get styles
        styles = rows[index].styles(range: range)

      } //end if (valid row index)

    } //end if found rows

    return styles

  } //end stylesForRow()

} //end extension Worksheet

extension Row {

  /**
   Retrieve the cell styles for this row
  */
  func styles(range: Range<Int>? = nil) -> [String?] {
    var styles: [String?] = []

    //process cells
    let allCells: [Cell] = self.orderedCells
    for cellIndex in range ?? 0..<allCells.count {
      if cellIndex >= 0 && cellIndex < allCells.count {
        styles.append(allCells[cellIndex].s)
      }
    } //end for (cells)

    return styles

  } //end styles()

} //end extension Row

extension ColumnReference: Hashable {
  public func hash(into hasher: inout Hasher) {
    hasher.combine(self.value)
  }
} //end extension ColumnReference
