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

//public struct MutableItem: Item {
//  override var items: [Items] = []
//}

public extension SharedStrings {
  /// create empty strings object
  init() {
    uniqueCount = 0
    items = []
  }

  init(from decoder: Decoder) throws {
    let values = try decoder.container(keyedBy: CodingKeys.self)
    let decodedItems = try values.decode([Item].self, forKey: .items)
    self.items = decodedItems
    self.uniqueCount = try values.decodeIfPresent(UInt.self, forKey: .uniqueCount) ?? UInt(decodedItems.count)
  }

  mutating func addString(_ text: String) -> Int {

    //check if string exists
    if let index = self.items.firstIndex(where: { $0.text == text }) {
      return index
    }
    //add new string
    else {
      let item = Item(text: text, richText: [])
      self.items.append(item)
      self.uniqueCount += 1
      return self.items.count - 1
    }

  } //end addString()

} // end extension SharedStrings
