//
//  Dictionary+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 14/7/20.
//

import Foundation

extension Array {
  internal func dictionary<T: Hashable, U>(_ transform: (_ index: Int, _ element: Array.Element) -> (key: T, value: U)) -> [T: U] {
    var map: [T: U] = [:]
    for (index, element) in self.enumerated() {
      let data = transform(index, element)
      map[data.key] = data.value
    }
    return map
  } //end dictionary()
} //end extension Array
