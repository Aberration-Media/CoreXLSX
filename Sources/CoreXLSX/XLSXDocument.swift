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

public enum CoreXLSXWriteError: Error {
  case fileAlreadyExists
  case couldNotCreateArchive
  case archiveEntryAlreadyExists
}

public class XLSXDocument {
  // MARK: Configuration XML

  // disabled lint rule to allow hard coded XML entries
  // swiftlint:disable line_length
  /// hardcoded app.xml document data
  private static let appXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"/>"

  /// hardcoded core.xml document data
  private static let coreXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>"
  // swiftlint:enable line_length

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

  // MARK: Document Properties

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
      // get document paths
      if let paths = try? safeFile.parseDocumentPaths() {
        // retrieve workbooks
        for path in paths {
          do {
            let workbook: Workbook = try safeFile.parseWorkbook(path: path)
            books.append((path: Path(rootPath: path), book: workbook))
          } catch {
            print("Error loading workbook(\(path)): \(error)")
          }
        } // end for()
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
      if let paths = try? safeFile.parseWorksheetPaths() {

        // retrieve work sheets
        for path in paths {
          do {
            let worksheet: Worksheet = try safeFile.parseWorksheet(at: path)
            sheets.append((path: Path(rootPath: path), sheet: worksheet))
          } catch {
            print("Error loading worksheet(\(path)): \(error)")
          }
        } // end for()
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
      } catch {
        print("Error loading root relationships: \(error)")
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
      if let paths = try? safeFile.parseDocumentPaths() {

        // retrieve relationship
        for path in paths {
          do {
            let relationships: (Path, Relationships) = try safeFile.parseDocumentRelationships(path: path)
            relations.append(relationships)
          } catch {
            print("Error loading document relationships(\(path)): \(error)")
          }
        } // end for()
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
    file = XLSXFile(filepath: filepath, bufferSize: bufferSize, errorContextLength: errorContextLength)
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
    print("save to archive: \(archiveURL)")
    // create archive
    if let archive = Archive(url: archiveURL, accessMode: .create) {
      // create content type file
      var contentConfiguration: ContentTypes = .standard

      // write root relationships
      try writeEntry(relationships, Self.rootRelationshipsPath.value, in: archive, rootAttributes: Self.relationshipsAttributes)

      // TODO: docProps are not required for a valid XLSX document - these should be optionally added to the document
      // write support files
      try writeSupportFile(Self.appXML, path: Self.appPath.value, in: archive)
      try writeSupportFile(Self.coreXML, path: Self.corePath.value, in: archive)

      // write styles
      if let safeStyles = styles {
        try writeEntry(safeStyles, Self.stylesPath.value, in: archive, withRootKey: "styleSheet", rootAttributes: Self.baseAttributes)
        // append styles to content types
        contentConfiguration.addOverride(with: Self.stylesPath.value, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
      }

      // write shared strings
      try writeEntry(self.sharedStrings, Self.sharedStringsPath.value, in: archive, withRootKey: "sst", rootAttributes: Self.baseAttributes)
      // append styles to content types
      contentConfiguration.addOverride(with: Self.sharedStringsPath.value, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")

      // save work books
      for data in workbooksMap {
        // write workbook
        try writeEntry(data.book, data.path.value, in: archive, withRootKey: "workbook")
        print("wrote workbook: \(data.path)")
        // append workbook path to content type configuration
        contentConfiguration.addOverride(with: data.path.value, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
      } // end for (workbooks)

      // save work sheets
      for data in worksheetsMap {
        // write worksheet
        try writeEntry(data.sheet, data.path.value, in: archive, withRootKey: "worksheet")
        // append worksheet path to content type configuration
        contentConfiguration.addOverride(with: data.path.value, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
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
          try writeEntry(relationships, relationshipsPath.value, in: archive, rootAttributes: Self.relationshipsAttributes)
        } // end if (valid document path)
      } // end for (relationships)

      // write content types xml file
      try writeEntry(contentConfiguration, "[Content_Types].xml", in: archive, withRootKey: "Types", rootAttributes: Self.contentTypesAttributes)
    } // end if (created archive)

    // could not create archive
    else {
      throw CoreXLSXWriteError.couldNotCreateArchive
    }
  } // end saveWorkbooks()

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
    let data: Data = try encoder.encode(entry, withRootKey: withRootKey, rootAttributes: rootAttributes, header: header)
    try archive.addEntry(with: entryPath, type: .file, uncompressedSize: UInt32(data.count), compressionMethod: .none, provider: { (_, _) -> Data in
      data
    })
  } // end writeEntry()
} // end struct XLSXDocument
