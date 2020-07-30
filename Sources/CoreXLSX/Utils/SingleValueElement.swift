//
//  SingleValueElement.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 30/7/20.
//

import Foundation

// MARK: - Coding Keys

enum SingleValueCodingKeys: String, CodingKey {
  case value = "val"
}

// MARK: - SingleValueElement

protocol SingleValueElement: Codable, Equatable {
  associatedtype ValueType: Codable

  var value: ValueType { get }
  init(value: ValueType)
}

extension SingleValueElement {
  public init(from decoder: Decoder) throws {
    let values = try decoder.container(keyedBy: SingleValueCodingKeys.self)
    let value = try values.decode(ValueType.self, forKey: .value)
    self.init(value: value)
  }
  public func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: SingleValueCodingKeys.self)
      try container.encode(value, forKey: .value)
  } //end encode()
}

// MARK: - OptionalSingleValueElement

protocol OptionalSingleValueElement: Codable, Equatable {
  associatedtype ValueType: Codable

  var value: ValueType? { get }
  init(value: ValueType?)
}

extension OptionalSingleValueElement {
  public init(from decoder: Decoder) throws {
    let values = try decoder.container(keyedBy: SingleValueCodingKeys.self)
    let value = try values.decodeIfPresent(ValueType.self, forKey: .value)
    self.init(value: value)
  }
  public func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: SingleValueCodingKeys.self)
      try container.encodeIfPresent(value, forKey: .value)
  } //end encode()
}

// MARK: - DoubleValueElement

protocol DoubleValueElement: SingleValueElement where ValueType == Double {}

// MARK: - StringValueElement

protocol StringValueElement: SingleValueElement where ValueType == String {}

// MARK: - OptionalBoolValueElement

protocol OptionalBoolValueElement: OptionalSingleValueElement where ValueType == Bool {}

extension OptionalBoolValueElement {
  public func encode(to encoder: Encoder) throws {
      var container = encoder.container(keyedBy: SingleValueCodingKeys.self)
      try container.encodeIfPresent(BoolConverter(value).intValue, forKey: .value)
  } //end encode()
}
