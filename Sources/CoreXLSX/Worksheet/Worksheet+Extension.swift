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
      return Int(last.reference)
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
    data = Worksheet.Data(rows: []) //Data is required for a valid XLSX
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
    let rowIndex: UInt = UInt(self.numberOfRows + 1) //XLSX rows start at 1 - using a cell reference with row 0 will crash Excel
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

    //update subsequent rows
    let startIndex = Int(rowIndex)
    if numberOfRows > startIndex {

      //move subsequent rows (as rows are moved down an error will never be thrown)
      try? self.shiftRows(in: startIndex..<numberOfRows, by: 1)
    }

    //create new row
    let row = Row(reference: rowIndex, height: height, customHeight: nil, cells: cells)
    self.data?.rowsByReference[rowIndex] = row

  } //end insertRow()

  /**
   Delete rows from the worksheet
   
    - parameters:
      - range: The Int range of rows to delete from the spread sheet
   */
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
   Delete columns from the worksheet
   
    - parameters:
      - range: The Int range of columns to delete from the spread sheet
   */
  mutating func deleteColumns(in range: Range<Int>) {

    //bounds check range
    var boundedRange: Range<Int> = range
    if boundedRange.lowerBound <= 0 {
      boundedRange = 1..<range.count
    }

    //found columns data
    if let columns: [Column] = self.columns?.items { //}.filter({ boundedRange.contains(Int($0.min)) || boundedRange.contains(Int($0.max)) }) {

      //process columns
      for (index, column) in columns.enumerated().reversed() {

        //split column (column covers delete range)
        if column.min < range.lowerBound && column.max >= range.upperBound {
          //create new column
          let newColumn = Column(min: UInt32(range.upperBound), max: column.max, width: column.width, style: column.style, customWidth: column.customWidth)
          self.columns?.items.insert(newColumn, at: index+1)

          //update existing column bounds
          self.columns?.items[index].max = UInt32(range.lowerBound)
        }

        //delete column (column exists entirely inside delete range)
        else if column.min >= range.lowerBound && column.max < range.upperBound {
          self.columns?.items.remove(at: index)
        }

        //update min range
        else if column.min < range.lowerBound {
          if boundedRange.contains(Int(column.max)) {
            self.columns?.items[index].max = UInt32(range.lowerBound - 1)
          }
        }

        //update max range
        else if column.max >= range.lowerBound {
          if boundedRange.contains(Int(column.min)) {
            self.columns?.items[index].min = UInt32(range.upperBound)
          }
        }
      } //end for (columns)

    } //end if (found columns)

    //found rows
    if let rows: [UInt: Row] = self.data?.rowsByReference {

      //delete rows
      for (key, row) in rows {
        var offset: Int = 0

        //remove cells from row (assumes 'cells' is an oredered list)
        var filteredCells: [ColumnReference: Cell] = [:]
        for var cell in row.cells {

          //get column index for cell
          let columnIndex: ColumnReference = cell.reference.column

          //update cell column index
          if let newIndex = ColumnReference(cell.reference.column.intValue + offset) {
            cell.reference.column = newIndex
          }

          //found cell to delete
          if boundedRange.contains(columnIndex.intValue) {
            offset += 1
          }
          //keep cell
          else {
            filteredCells[columnIndex] = cell
          }

        } //end for (cells in row)

        //update cell references
        self.data?.rowsByReference[key]?.cellsByReference = filteredCells

      } //end for (rows)

      //shift remaining rows up
      try? self.shiftRows(in: boundedRange.upperBound..<self.numberOfRows, by: -boundedRange.count)

    } //end if (found rows)

  } //end deleteColumns()

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
