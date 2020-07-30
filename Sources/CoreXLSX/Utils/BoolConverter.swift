//
//  BoolConverter.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 29/7/20.
//

struct BoolConverter {

  /// boolean value
  let value: Bool?

  /// boolean value represented as an integer
  let intValue: Int?

  public init(_ value: Bool?) {
    self.value = value
    switch value {
    case true: self.intValue = 1
    case false: self.intValue = 0
    default: self.intValue = nil
    }
  } //end constructor()

} //end struct BoolConverter
