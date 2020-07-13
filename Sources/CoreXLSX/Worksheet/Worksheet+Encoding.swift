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
