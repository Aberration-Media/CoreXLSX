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
} // end extension Workbook

public extension Workbook.Sheets {
  /// create empty workbook sheets
  init() {
    items = []
  }
} // end extension Workbook.Sheets
