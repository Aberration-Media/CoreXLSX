//
//  SharedStrings+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension SharedStrings: ExcludeAttributeType {
  public static var nonAttributeKeys: [String]? {
    return ["si"]
  }
}

public extension SharedStrings {
  /// create empty strings object
  init() {
    uniqueCount = 0
    items = []
  }
} // end extension SharedStrings
