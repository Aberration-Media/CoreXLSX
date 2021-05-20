//
//  Workbook+Extensions.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

public extension Workbook {

  /// create empty workbook
  init() {
    self.properties = nil
    self.views = nil
    self.sheets = Sheets()
  }

  mutating func addSheet(named: String, relationshipId: String) {
    let sheetID: String = "\(self.sheets.items.count + 1)" //id's start at 1
    let sheet = Workbook.Sheet(name: named, id: sheetID, relationship: relationshipId)
    self.sheets.items.append(sheet)

  } //end addSheet

} // end extension Workbook

public extension Workbook.Sheets {

  /// create empty workbook sheets
  init() {
    items = []
  }

} // end extension Workbook.Sheets
