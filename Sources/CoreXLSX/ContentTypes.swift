//
//  ContentTypes.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

public struct ContentTypes: Codable {
  // MARK: - Default

  public struct Default: Codable, AttributeType {
    public let `extension`: String
    public let type: String

    enum CodingKeys: String, CodingKey {
      case type = "ContentType"
      case `extension` = "Extension"
    }
  }

  // MARK: - Override

  public struct Override: Codable, AttributeType {
    public let partName: String
    public let type: String

    enum CodingKeys: String, CodingKey {
      case type = "ContentType"
      case partName = "PartName"
    }
  }

  // MARK: Properties

  public var defaults: [Default]
  public var overrides: [Override]

  public enum CodingKeys: String, CodingKey {
    case defaults = "Default"
    case overrides = "Override"
  }
} // end struct ContentTypes

public extension ContentTypes {
  // MARK: Standard

  /// content types populated with default formats
  static let standard: ContentTypes = {
    let defaults: [Default] = [
      Default(extension: "xml", type: "application/xml"),
      Default(extension: "rels", type: "application/vnd.openxmlformats-package.relationships+xml"),
      Default(extension: "jpeg", type: "image/jpg"),
      Default(extension: "png", type: "image/png"),
      Default(extension: "bmp", type: "image/bmp"),
      Default(extension: "gif", type: "image/gif"),
      Default(extension: "tif", type: "image/tif"),
      Default(extension: "pdf", type: "application/pdf"),
      Default(extension: "mov", type: "application/movie"),
      Default(extension: "vml", type: "application/vnd.openxmlformats-officedocument.vmlDrawing"),
      Default(extension: "xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ]
    let overrides: [Override] = [
      Override(partName: "/docProps/core.xml", type: "application/vnd.openxmlformats-package.core-properties+xml"),
      Override(partName: "/docProps/app.xml", type: "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]

    var config = ContentTypes(defaults: defaults, overrides: overrides)
    return config
  }()

  // MARK: - Content Functions

  mutating func addOverride(with partName: String, type: String) {
    let override = Override(partName: partName, type: type)
    overrides.append(override)
  }
} // end extension ContentTypes
