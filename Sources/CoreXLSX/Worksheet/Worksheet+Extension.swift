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
