//
//  AttributedType.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import XMLCoder

// MARK: - AttributeType

public protocol AttributeType: DynamicNodeEncoding, DynamicNodeDecoding {} // end protocol AttributeType

public extension AttributeType {
  static func nodeEncoding(for key: CodingKey) -> XMLEncoder.NodeEncoding {
    return .attribute
  }

  static func nodeDecoding(for key: CodingKey) -> XMLDecoder.NodeDecoding {
    return .attribute
  }
} // end extension AttributeType

// MARK: - ExcludeAttributeType

public protocol ExcludeAttributeType: DynamicNodeEncoding, DynamicNodeDecoding {
  /// properties to handle with default encoding
  static var nonAttributeKeys: [String]? { get }
}

public extension ExcludeAttributeType {
  static func nodeEncoding(for key: CodingKey) -> XMLEncoder.NodeEncoding {
    return nonAttributeKeys?.contains(key.stringValue) == true ? .default : .attribute
  }

  static func nodeDecoding(for key: CodingKey) -> XMLDecoder.NodeDecoding {
    return nonAttributeKeys?.contains(key.stringValue) == true ? .elementOrAttribute : .attribute
  }
} // end extension ExcludeAttributeType
