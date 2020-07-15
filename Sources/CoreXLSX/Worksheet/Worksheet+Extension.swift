//
//  Worksheet+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

// MARK: Errors

public enum CoreXLSXDocumentError: Error {
  case rowsOutOfBounds
}

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

  /// create empty worksheet
  init() {
    properties = nil
    formatProperties = nil
    dimension = nil
    sheetViews = nil
    columns = nil
    data = nil
    mergeCells = nil
  } //end constructor()

  mutating func applyCells(_ cells: [Cell], sharedStrings: inout SharedStrings) {
    for cell in cells {
      self.applyCell(cell, sharedStrings: &sharedStrings)
    }
  } //end applyCells()

  mutating func applyCell(_ cell: Cell, sharedStrings: inout SharedStrings) {
    let rowIndex: UInt = cell.reference.row
    let columnIndex: ColumnReference = cell.reference.column

    //valid row index
    if self.data?.rowsByReference[rowIndex] != nil {
      self.data?.rowsByReference[rowIndex]?.cellsByReference[columnIndex] = cell
    }
    //invalid row index
    else {
      print("No row found at index: \(cell.reference)")
    }
  } //end applyCell()

  mutating func updateCellValue(
    at rowIndex: UInt,
    column: ColumnReference,
    with value: String,
    sharedStrings: inout SharedStrings,
    newCellStyle: String? = nil
  ) {

    //add shared string
    let valueIndex: Int = sharedStrings.addString(value)

    //valid row index
    if self.data?.rowsByReference[rowIndex] != nil {
      //update existing cell
      if self.data?.rowsByReference[rowIndex]?.cellsByReference[column] != nil {
        self.data?.rowsByReference[rowIndex]?.cellsByReference[column]?.value = String(valueIndex)
        self.data?.rowsByReference[rowIndex]?.cellsByReference[column]?.type = .sharedString
      }
      //create new cell
      else {
        let cellReference = CellReference(column, rowIndex)
        let cell = Cell(reference: cellReference, type: .sharedString, s: newCellStyle, inlineString: nil, formula: nil, value: String(valueIndex))
        self.data?.rowsByReference[rowIndex]?.cellsByReference[column] = cell
      }
    }
    //invalid row index
    else {
      print("No row found at index: \(rowIndex)")
    }

  } //end updateCellValue()

  mutating func updateRowValues(
    at rowIndex: UInt,
    column: UInt = 1,
    with values: [String],
    sharedStrings: inout SharedStrings
  ) {

    //valid row index
    if self.data?.rowsByReference[rowIndex] != nil {

      //update cells
      let firstColumn: Int = Int(column < 1 ? 1 : column)
      for (index, value) in values.enumerated() {
        if let columnReference = ColumnReference(index + firstColumn) {
          self.updateCellValue(at: rowIndex, column: columnReference, with: value, sharedStrings: &sharedStrings)
        }
      } //end for (values)

    }
    //invalid row index
    else {
      print("No row found at index: \(rowIndex)")
    }

  } //end updateRowValues()


  mutating func updateColumnValues(
    at columnIndex: Int,
    row: UInt = 1,
    with values: [String],
    sharedStrings: inout SharedStrings
  ) {
    //valid index
    if let columnReference = ColumnReference(columnIndex) {

      //update values
      let firstRow: UInt = row < 1 ? 1 : row
      for (index, value) in values.enumerated() {

        //create row (if required)
        let rowIndex: UInt = firstRow + UInt(index)
        if self.data?.rowsByReference[rowIndex] == nil {
          let row = Row(reference: rowIndex, height: nil, customHeight: nil, cells: [])
          self.data?.rowsByReference[rowIndex] = row
        }

        //set value
        self.updateCellValue(at: rowIndex, column: columnReference, with: value, sharedStrings: &sharedStrings)

      } //end for (values)

    }
    //invalid row index
    else {
      print("Invalid column index: \(columnIndex)")
    }

  } //end updateColumnValues()

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
    with values: [String],
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

    //move subsequent rows (as rows are moved down an error will never be thrown)
    try? self.shiftRows(in: Int(rowIndex)..<numberOfRows, by: 1)

    //create new row
    let row = Row(reference: rowIndex, height: height, customHeight: nil, cells: cells)
    self.data?.rowsByReference[rowIndex] = row

  } //end insertRow()

  mutating func deleteRows(in range: Range<Int>) {
    //bounds check range
    var boundedRange: Range<Int> = range
    if boundedRange.lowerBound <= 0 {
      boundedRange = 1..<range.count
    }

    //found rows
    if let rowsInRange: [Row] = self.data?.rows.filter({ boundedRange.contains(Int($0.reference)) }) {
      //delete rows
      for row in rowsInRange {
        self.data?.rowsByReference.removeValue(forKey: row.reference)
      }

      //shift remaining rows up
      try? self.shiftRows(in: boundedRange.upperBound..<self.numberOfRows, by: -boundedRange.count)

    } //end if (found rows)
  } //end deleteRows()

  /**
    Move rows to a new row reference position

    This is for processing of the internal data structure. Public users should use insertRow / deleteRow functionality to achieve the same result
   */
  mutating internal func shiftRows(in range: Range<Int>, by rowsShift: Int) throws {
    //found rows
    if var rowsInRange: [Row] = self.data?.rows.filter({ range.contains(Int($0.reference)) }) {

      //shifting down (start with last row)
      if rowsShift > 0 {
        rowsInRange.reverse()
      }

      //move rows
      for var row in rowsInRange {

        //determine index
        let shiftedIndex: Int = Int(row.reference) + rowsShift
        //valid shift position
        if shiftedIndex >= 1 {
          let newRowIndex = UInt(shiftedIndex)
          for (index, _) in row.cellsByReference {
            row.cellsByReference[index]?.reference.row = newRowIndex
          }
          self.data?.rowsByReference.removeValue(forKey: row.reference)
          row.reference = newRowIndex
          self.data?.rowsByReference[newRowIndex] = row
        }
        //moving beyond zero row
        else {
          throw CoreXLSXDocumentError.rowsOutOfBounds
        }
      }
    } //end if (found rows)

  } //end shiftRows()

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
    let allCells: [Cell] = self.cells
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
