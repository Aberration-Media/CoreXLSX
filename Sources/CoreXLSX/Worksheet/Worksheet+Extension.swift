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

public extension Worksheet {

  // MARK: Convenience Properties

  /// number of rows in worksheet
  var numberOfRows: Int { self.data?.rows.count ?? 0 }

  /// the first row in the spread sheet
  var firstRow: Row? { self.data?.rows.first }

  /// the last row in the spread sheet
  var lastRow: Row? { self.data?.rows.last }

  // MARK: - Editing Functions

  mutating func updateCell(
    at row: Int,
    column: ColumnReference,
    with value: String,
    sharedStrings: inout SharedStrings
  ) {

    //valid row index
    if row >= 0 && row < self.data?.rows.count ?? 0 {
      //update existing cell
      if let cellIndex: Int = self.data?.rows[row].cells.firstIndex(where: { $0.reference.column == column }) {
        self.data?.rows[row].cells[cellIndex].value = value
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
      print("No row found at index: \(row)")
    }

  } //end updateCell()

  mutating func updateRow(
    at index: Int,
    from column: Int = 1,
    with values: [String],
    sharedStrings: inout SharedStrings
  ) {

    //valid row index
    if index >= 0 && index < self.data?.rows.count ?? 0 {

      //update cells
      let firstColumn: Int = column < 1 ? 1 : column
      for (index, value) in values.enumerated() {
        if let columnReference = ColumnReference(index + firstColumn) {
          self.updateCell(at: index, column: columnReference, with: value, sharedStrings: &sharedStrings)
        }
      } //end for (values)

    }
    //invalid row index
    else {
      print("No row found at index: \(index)")
    }

  } //end updateRow()

  /**
    Add a row with the specified values
   */
  mutating func addRow(
    with values: [String] = [],
    sharedStrings: inout SharedStrings,
    height: Double? = nil,
    styles: [String?]? = nil
  ) {

    //get row reference
    let rowReference: UInt = UInt(self.data?.rows.count ?? 0) + 1

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
        let cellReference = CellReference(columnReference, rowReference)
        let cell = Cell(reference: cellReference, type: .sharedString, s: style, inlineString: nil, formula: nil, value: String(valueIndex))
        cells.append(cell)

      }
    } //end for (values)

    //create new row
    let row = Row(reference: rowReference, height: height, customHeight: nil, cells: cells)
    self.data?.rows.append(row)

  } //end addRow()

  /**
    Retrieve the styles for the specified row index
   */
  func stylesForRow(_ index: Int, range: Range<Int>? = nil) -> [String?] {
    var styles: [String?] = []

    //found rows
    if let rows = self.data?.rows {

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
    for cellIndex in range ?? 0..<self.cells.count {
      if cellIndex >= 0 && cellIndex < self.cells.count {
        styles.append(self.cells[cellIndex].s)
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
