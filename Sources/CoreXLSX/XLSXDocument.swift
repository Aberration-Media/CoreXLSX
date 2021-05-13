//
//  XLSXDocument.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation
import XMLCoder
import ZIPFoundation

// MARK: Errors

public enum CoreXLSXWriteError: Error, CustomStringConvertible {
  case fileAlreadyExists
  case couldNotCreateArchive
  case archiveEntryAlreadyExists
  case unrecognizedContentType(path: String, type: String)

  public var description: String {
    var message: String = "CoreXLSXWriteError: "
    switch self {
      case .fileAlreadyExists:
        message += "File already exists"
      case .couldNotCreateArchive:
        message += "Could not create archive"
      case .archiveEntryAlreadyExists:
        message += "Entry already exists"
      case let .unrecognizedContentType(path, type):
        message += "Unrecognized content type(\(type)) at path: \(path)"
      default:
        message += "Unknown"
    }
    return message
  }
} //end enum CoreXLSXWriteError

public class XLSXDocument {

  // MARK: Configuration XML

  // disabled lint rule to allow hard coded XML entries
  // swiftlint:disable line_length
  /// hardcoded app.xml document data
  private static let appXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"/>"

  /// hardcoded core.xml document data
  private static let coreXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>"
  // swiftlint:enable line_length

  /// workbook properties default theme version
  private static let defaultThemeVersion: Int = 164011

  // MARK: Document Paths

  /// '_rels/.rels' document path
  internal static let rootRelationshipsPath = Path("/_rels/.rels")

  /// 'docProps' document path
  internal static let propertiesPath = Path("/docProps")

  /// 'xl' document path
  internal static let excelPath = Path("/xl")

  /// 'xl/worksheets' document path
  internal static let worksheetsPath = excelPath.path(byAppending: "worksheets")

  /// 'xl/_rels' document path
  internal static let documentRelationshipsPath = excelPath.path(byAppending: "_rels")

  /// 'xl/styles.xml' document path
  internal static let stylesPath = excelPath.path(byAppending: "styles.xml")

  /// 'xl/sharedStrings.xml' document path
  internal static let sharedStringsPath = excelPath.path(byAppending: "sharedStrings.xml")

  /// 'docProps/app.xml' document path
  internal static let appPath = propertiesPath.path(byAppending: "app.xml")

  /// 'docProps/core.xml' document path
  internal static let corePath = propertiesPath.path(byAppending: "core.xml")

  // MARK: Attributes

  /// XML header
  private static let header: XMLHeader = XMLHeader(version: 1.0, encoding: "UTF-8")

  /// XML base element attributes
  private static let baseAttributes: [String: String] = [
    "xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  ]

  /// XML document element attributes
  private static let documentAttributes: [String: String] = [
    "xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  ]

  /// XML content types attributes
  private static let contentTypesAttributes: [String: String] = [
    "xmlns": "http://schemas.openxmlformats.org/package/2006/content-types",
  ]

  /// XML relationships element attributes
  private static let relationshipsAttributes: [String: String] = [
    "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships",
  ]

  /// XML styles element attributes
  private static let stylesAttributes: [String: String] = [
    "xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mc:Ignorable": "x14ac x16r2",
    "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xmlns:x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
  ]

  // MARK: Document Properties

  /// delegate for document events
  public weak var documentDelegate: XLSXDocumentDelegate?

  /// file associated with document
  public let file: XLSXFile?

  /// list of work books in document
  public var workbooks: [Workbook] {
    return workbooksMap.map { $0.book }
  }

  /// list of work books in document mapped to their corresponding path
  internal lazy var workbooksMap: [(path: Path, book: Workbook)] = {
    var books: [(Path, Workbook)] = []

    // file exists
    if let safeFile = file {
      do {
        // get document paths
        let paths = try safeFile.parseDocumentPaths()
        // retrieve workbooks
        for path in paths {
          do {
            let workbook: Workbook = try safeFile.parseWorkbook(path: path)
            books.append((path: Path(rootPath: path), book: workbook))
          } catch {
            print("Error loading workbook(\(path)): \(error)")
            documentDelegate?.didReceiveError(for: self, error: error)
          }
        } // end for()

      } catch {
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    return books

  }()

  /// list of work sheets in document
  public var worksheets: [Worksheet] {
    return worksheetsMap.map { $0.sheet }
  }

  /// list of work sheets in document mapped to their corresponding path
  public lazy var worksheetsMap: [(path: Path, sheet: Worksheet)] = {
    var sheets: [(path: Path, sheet: Worksheet)] = []

    // file exists
    if let safeFile = file {
      // get worksheet paths
      do {
        let paths = try safeFile.parseWorksheetPaths()
        // retrieve work sheets
        for path in paths {
          do {
            let worksheet: Worksheet = try safeFile.parseWorksheet(at: path)
            sheets.append((path: Path(rootPath: path), sheet: worksheet))
          } catch {
            print("Error loading worksheet(\(path)): \(error)")
            documentDelegate?.didReceiveError(for: self, error: error)
          }
        } // end for()
      } catch {
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    return sheets

  }()

  /// root relationships for document
  public lazy var relationships: Relationships = {
    var relations = Relationships()

    // file exists
    if let safeFile = file {
      // get relationships
      do {
        relations = try safeFile.parseRelationships()
        print("GOT DOCUMENT RELATIONS: \(relations)")
      } catch {
        print("Error loading root relationships: \(error)")
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    return relations

  }()

  /// list of document relationships
  public lazy var documentRelationships: [(path: Path, relations: Relationships)] = {
    var relations: [(path: Path, relations: Relationships)] = []

    // file exists
    if let safeFile = file {
      // get document paths
      do {
        let paths = try safeFile.parseDocumentPaths()

        // retrieve relationship
        for path in paths {
          do {
            let relationships: (Path, Relationships) = try safeFile.parseDocumentRelationships(path: path)
            relations.append(relationships)
          } catch {
            print("Error loading document relationships(\(path)): \(error)")
          }
        } // end for()
      } catch {
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    return relations

  }()

  /// list of styles in document
  public lazy var styles: Styles? = {
    var result: Styles?

    // file exists
    if let safeFile = file {
      // get styles
      do {
        result = try safeFile.parseStyles()
      } catch {
        print("Error loading styles: \(error)")
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    return result

  }()

  /// shared strings in document
  public lazy var sharedStrings: SharedStrings = {

    // file exists
    if let safeFile = file {
      // get shared strings
      do {
        return try safeFile.parseSharedStrings()
      } catch {
        print("Error loading shared strings: \(error)")
        documentDelegate?.didReceiveError(for: self, error: error)
      }
    } // end if (found associated file)

    //create empty shared strings
    let result = SharedStrings()
    return result

  }()

  // MARK: Encoding Properties

  /// XML encoder
  private lazy var encoder: XMLEncoder = {
    let result = XMLEncoder()
    return result
  }()

  // MARK: - Configuration Functions

  public init() {
    // initialise properties
    file = nil
    let workbook: Workbook = createWorkbook(named: "workbook")
    createWorksheet(named: "sheet1", for: workbook)
  } // end constructor()

  public init(with file: XLSXFile) {
    // store file
    self.file = file
  } // end constructor()

  public init(filepath: String, bufferSize: UInt32 = 10 * 1024 * 1024, errorContextLength: UInt = 0) {
    // create file
    self.file = XLSXFile(filepath: filepath, bufferSize: bufferSize, errorContextLength: errorContextLength)
  } // end constructor()

  #if swift(>=5.0)
  public init(data: Data, bufferSize: UInt32 = 10 * 1024 * 1024, errorContextLength: UInt = 0) throws {
    // create file
    try file = XLSXFile(data: data, bufferSize: bufferSize, errorContextLength: errorContextLength)
  } // end constructor()
  #endif

  // MARK: - Document Functions

  public func save(to filePath: String, overwrite: Bool = false) throws {
    // create archive URL
    let archiveURL = URL(fileURLWithPath: filePath)

    // check for existing file
    let fileManger = FileManager()
    if fileManger.fileExists(atPath: archiveURL.path) {
      // overwrite
      if overwrite {
        // delete previous file
        try fileManger.removeItem(at: archiveURL)
      }
      // throw error
      else {
        throw CoreXLSXWriteError.fileAlreadyExists
      }
    } // end if (file exists)

    // TODO: There seems to be a Excel compatibility issue with ZipFoundation
    // Possible causes related to the default system byte or version information:
    // https://github.com/mvdnes/zip-rs/issues/23
    // https://github.com/mvdnes/zip-rs/issues/72
    // create archive
    if let archive = Archive(url: archiveURL, accessMode: .create) {
      // create content type file
      var contentConfiguration: ContentTypes = .standard

      // write root relationships
      try writeEntry(self.relationships, Self.rootRelationshipsPath.value, in: archive, rootAttributes: Self.relationshipsAttributes)

      // TODO: docProps are not required for a valid XLSX document - these should be optionally added to the document
      // write support files
      try writeSupportFile(Self.appXML, path: Self.appPath.value, in: archive)
      try writeSupportFile(Self.coreXML, path: Self.corePath.value, in: archive)

      // write shared strings
      try writeEntry(self.sharedStrings, Self.sharedStringsPath.value, in: archive, withRootKey: "sst", rootAttributes: Self.baseAttributes)
      // append styles to content types
      contentConfiguration.addOverride(with: Self.sharedStringsPath.value, type: .sharedStrings)

      // save work books
      for var data in workbooksMap {

        //ensure workbook properties exist with defaultThemeVersion (required for Excel compatibility)
        if data.book.properties == nil {
          data.book.properties = Workbook.Properties(defaultThemeVersion: Self.defaultThemeVersion, dateCompatibility: nil)
        } else if data.book.properties?.defaultThemeVersion == nil {
          data.book.properties?.defaultThemeVersion = Self.defaultThemeVersion
        }

        // write workbook
        try writeEntry(data.book, data.path.value, in: archive, withRootKey: "workbook")

        // append workbook path to content type configuration
        contentConfiguration.addOverride(with: data.path.value, type: .workbook)

      } // end for (workbooks)

      // write styles
      if let safeStyles = styles {
        try self.writeEntry(safeStyles, Self.stylesPath.value, in: archive, withRootKey: "styleSheet", rootAttributes: Self.stylesAttributes)
        // append styles to content types
        contentConfiguration.addOverride(with: Self.stylesPath.value, type: .styles)
      }

      // save work sheets
      for data in worksheetsMap {

        // write worksheet
        try self.writeEntry(data.sheet, data.path.value, in: archive, withRootKey: "worksheet")
        // append worksheet path to content type configuration
        contentConfiguration.addOverride(with: data.path.value, type: .worksheet)

      } // end for (worksheets)

      // save document relationships
      for data in documentRelationships {

        // get data components
        let documentPath: Path = data.path
        let relationships: Relationships = data.relations

        // valid document path
        if let documentFileName: String = documentPath.lastPathComponent, !documentFileName.isEmpty {

          // create relationship path
          let relationshipsPath: Path = Self.documentRelationshipsPath.path(byAppending: "\(documentFileName).rels")

          // write relationships
          try self.writeEntry(relationships, relationshipsPath.value, in: archive, rootAttributes: Self.relationshipsAttributes)

        } // end if (valid document path)

        //save any unsupported files (relationships that do not correspond to a Swift data object)
        try self.saveUnsupportedFiles(from: data.relations, to: archive, configuration: &contentConfiguration, rootPath: Self.excelPath)

      } // end for (relationships)

      // write content types xml file
      try self.writeEntry(contentConfiguration, "[Content_Types].xml", in: archive, withRootKey: "Types", rootAttributes: Self.contentTypesAttributes)

      //save any unsupported/missing files (relationships that do not correspond to a Swift data object)
      try self.saveUnsupportedFiles(from: self.relationships, to: archive, configuration: &contentConfiguration)

    } // end if (created archive)

    // could not create archive
    else {
      throw CoreXLSXWriteError.couldNotCreateArchive
    }
  } // end saveWorkbooks()

  private func saveUnsupportedFiles(from relationships: Relationships, to targetArchive: Archive, configuration: inout ContentTypes, rootPath: Path? = nil) throws {

    //TODO: All relationships should have an associated XML coding object (this copy phase could then be skipped)
    //find additional relationships
    for relationship in relationships.items {

      //check if item exists in new archive
      do {

        //create root based path
        let fullPath: String
        if let path = rootPath {
          fullPath = path.path(byAppending: relationship.target).value
        } else {
          fullPath = relationship.target
        }

        //copy file
        try self.file?.copyEntry(at: fullPath, to: targetArchive)

        //add file to content types configuration
        if !configuration.containsOverride(for: fullPath) {

          //known application type
          if let type = ContentTypes.ApplicationType(from: relationship.type) {
            configuration.addOverride(with: fullPath, type: type)
          }
          //unknown type
          else {
            throw CoreXLSXWriteError.unrecognizedContentType(path: relationship.target, type: relationship.type.rawValue)
          }
        }
      }
      catch CoreXLSXWriteError.archiveEntryAlreadyExists {
        //do nothing - entity already exists
      }
      catch {
        throw error
      }

    } //end for (relationships)

  } //end saveUnsupportedFiles()

  // MARK: - File Functions

  private func writeSupportFile(_ contents: String, path: String, in archive: Archive) throws {
    if let data: Data = contents.data(using: .utf8) {
      // add file to archive
      try archive.addEntry(with: path, type: .file, uncompressedSize: UInt32(data.count), compressionMethod: .deflate, provider: { (_, _) -> Data in
        data
      })
    } // end if (got data)
  } // end writeSupportFile()

  private func writeEntry<T: Encodable>(
    _ entry: T,
    _ pathString: String,
    in archive: Archive,
    withRootKey: String? = nil,
    rootAttributes: [String: String]? = XLSXDocument.documentAttributes,
    header: XMLHeader? = XLSXDocument.header
  ) throws {
    // get entry path
    let path = Path(pathString)
    let entryPath = path.isRoot ? path.components.joined(separator: "/") : pathString

    // check if entry already exists
    if archive[entryPath] != nil {
      throw CoreXLSXWriteError.archiveEntryAlreadyExists
    }

    // encode entry
    let entryData: Data = try encoder.encode(entry, withRootKey: withRootKey, rootAttributes: rootAttributes, header: header)
    let size = entryData.count
    try archive.addEntry(with: entryPath, type: .file, uncompressedSize: UInt32(size), compressionMethod: .deflate, provider: {
      (position, bufferSize) -> Data in

      //copy data chunk
      let upperBound = Swift.min(size, position + bufferSize)
      let range = Range(uncheckedBounds: (lower: position, upper: upperBound))
      return entryData.subdata(in: range)
    })

  } // end writeEntry()

} // end struct XLSXDocument
