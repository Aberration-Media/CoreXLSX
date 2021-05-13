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
  
  // MARK: Types

  enum ApplicationType: String {
    case core = "application/vnd.openxmlformats-package.core-properties+xml"
    case extended = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    case workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    case worksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    case sharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    case styles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    case theme = "application/vnd.openxmlformats-officedocument.theme+xml"

    public init?(from schemaType: Relationship.SchemaType) {

      //match SchemaType to ApplicationType
      switch schemaType {
        case .packageCoreProperties:
          self = .core
        case .extendedProperties:
          self = .extended
        case .worksheet:
          self = .worksheet
        case .sharedStrings:
          self = .sharedStrings
        case .styles:
          self = .styles
        case .theme:
          self = .theme
        default:
          return nil

      } //end switch (SchemaType)

    } //end constructor()

  } //end enum ApplicationType

  // MARK: Configurations

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
      Override(partName: "/docProps/core.xml", type: ApplicationType.core.rawValue),
      Override(partName: "/docProps/app.xml", type: ApplicationType.extended.rawValue),
    ]

    var config = ContentTypes(defaults: defaults, overrides: overrides)
    return config
  }()

  // MARK: - Content Functions

  /**
    Add an override part to the configuration
   
    - parameters:
      - partName: The localalised file path to the associatied override
      - type: The raw string type
   */
  mutating func addOverride(with partName: String, type: ApplicationType) {
    let override = Override(partName: partName, type: type.rawValue)
    overrides.append(override)
  }

  func containsOverride(for partName: String) -> Bool {
    return overrides.contains(where: { $0.partName == partName })
  }

} // end extension ContentTypes
