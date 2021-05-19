// Copyright 2020 CoreOffice contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
//  Created by Max Desiatov on 27/10/2018.
//

import Foundation
import XMLCoder
import ZIPFoundation


// MARK: -
// MARK: - CoreXLSXError -

@available(*, deprecated, renamed: "CoreXLSXError")
public typealias XLSXReaderError = CoreXLSXError

public enum CoreXLSXError: Error {

  case dataIsNotAnArchive
  case archiveEntryNotFound
  case invalidCellReference
  case unsupportedWorksheetPath
  case invalidDocumentPath

} //end enum CoreXLSXError




// MARK: -
// MARK: - XLSXFile -

/** The entry point class that represents an open file handle to an existing `.xlsx` file on the
 user's filesystem.
 */
public class XLSXFile {

  // MARK:
  // MARK: Path Constants

  /// directory name for relationship files
  internal static let relationshipsFolderName: Substring = "_rels"

  /// file name for relationship files
  internal static let relationshipsFileExtension: Substring = ".rels"


  // MARK:
  // MARK: File Properties

  /// URL of XLSX file
  public var fileURL: URL { return archive.url }

  /// archive containing XLSX data
  private let archive: Archive


  // MARK:
  // MARK: Decoding Properties

  /// XML decoder for parsing XLSX documents
  private let decoder: XMLDecoder = {
    let result = XMLDecoder()
    result.trimValueWhitespaces = false
    result.shouldProcessNamespaces = true
    return result
  }()

  /// Buffer size passed to `archive.extract` call
  private let bufferSize: UInt32


  // MARK:
  // MARK: - Configuration Functions

  /// - Parameters:
  ///   - filepath: path to the `.xlsx` file to be processed.
  ///   - bufferSize: ZIP archive buffer size in bytes. The default is 10KB.
  /// You may need to set a bigger buffer size for bigger files.
  ///   - errorContextLength: The error context length. The default is `0`.
  /// Non-zero length makes an error thrown from
  /// the XML parser with line/column location repackaged with a context
  /// around that location of specified length. For example, if an error was
  /// thrown indicating that there's an unexpected character at line 3, column
  /// 15 with `errorContextLength` set to 10, a new error type is rethrown
  /// containing 5 characters before column 15 and 5 characters after, all on
  /// line 3. Line wrapping should be handled correctly too as the context can
  /// span more than a few lines.
  public init?(
    filepath: String,
    bufferSize: UInt32 = 10 * 1024 * 1024,
    errorContextLength: UInt = 0
  ) {
    let archiveURL = URL(fileURLWithPath: filepath)

    guard let archive = Archive(url: archiveURL, accessMode: .read) else {
      return nil
    }

    self.archive = archive
    self.bufferSize = bufferSize
    decoder.errorContextLength = errorContextLength

  } //end constructor()


#if swift(>=5.0)
  /// - Parameters:
  ///   - data: content of the `.xlsx` file to be processed.
  ///   - bufferSize: ZIP archive buffer size in bytes. The default is 10KB.
  /// You may need to set a bigger buffer size for bigger files.
  ///   - errorContextLength: The error context length. The default is `0`.
  /// Non-zero length makes an error thrown from
  /// the XML parser with line/column location repackaged with a context
  /// around that location of specified length.
  public init(
    data: Data,
    bufferSize: UInt32 = 10 * 1024 * 1024,
    errorContextLength: UInt = 0
  ) throws {

    guard let archive = Archive(data: data, accessMode: .read)
    else { throw CoreXLSXError.dataIsNotAnArchive }

    self.archive = archive
    self.bufferSize = bufferSize
    decoder.errorContextLength = errorContextLength

  } //end constructor()

#endif


  // MARK: - Decoding Functions

  /// Parse a file within `archive` at `path`. Parsing result is
  /// an instance of `type`.
  func parseEntry<T: Decodable>(
    _ pathString: String,
    _ type: T.Type
  ) throws -> T {
    let path = Path(pathString)
    let entryPath = path.isRoot ?
      path.components.joined(separator: "/") :
      pathString

    guard let entry = archive[entryPath] else {
      throw CoreXLSXError.archiveEntryNotFound
    }

    var data = Data()
    _ = try archive.extract(entry, bufferSize: bufferSize) {
      data += $0
    }

    return try decoder.decode(type, from: data)

  } //end parseEntry()


  // MARK: - File Functions

  /**
    Copy file entry to another archive
   
    - parameters:
      - pathString: Entry String path in file archive
      - destination: Target archive which will add a copy of the orginal file
      - compressionMethod: Optional compression method setting, the default is 'deflate'
   */
  public func copyEntry(at path: Path, to destination: Archive, compressionMethod: CompressionMethod = .deflate) throws {
    let entryPath = path.relativePath

    //check if entry exists in file archive
    guard let entry: Entry = archive[entryPath] else {
      print("failed to find entry: \(entryPath)")
      throw CoreXLSXError.archiveEntryNotFound
    }

    // check if entry already exists in destination archive
    if destination[entryPath] != nil {
      throw CoreXLSXWriteError.archiveEntryAlreadyExists
    }

    //copy data from existing archive
    let entryData = NSMutableData()
    _ = try archive.extract(entry, bufferSize: UInt32(entry.uncompressedSize), skipCRC32: false, progress: nil, consumer: { data in
      entryData.append(data)
    })

    //copy data into destination archive
    let size = entryData.count
    try destination.addEntry(with: entryPath, type: .file, uncompressedSize: UInt32(size), compressionMethod: .deflate, provider: {
      (position, bufferSize) -> Data in

      //copy data chunk
      let upperBound = Swift.min(size, position + bufferSize)
      let range = NSRange(location: position, length: upperBound - position)
      return entryData.subdata(with: range)
    })

  } //end copyEntry()



  // MARK: - Relationship Functions

  /// returns the relationship path associated with the specified document path
  internal static func relationshipPath(for documentPath: Path) -> Path? {

    var components = documentPath.evaluatedPathComponents

    //found file components
    if !components.isEmpty, let fileName: Substring = components.last {

      //insert relationships folder path
      components.insert(Self.relationshipsFolderName, at: components.count - 1)

      //replace path file name with relationships filename
      components[components.count - 1] = Substring(fileName.appending(Self.relationshipsFileExtension))

      return Path(components.joined(separator: "/"))

    } //end if (found file components)

    return nil

  } //end relationshipPath()


  public func parseRelationships() throws -> Relationships {
    decoder.keyDecodingStrategy = .useDefaultKeys

    //parse root relationships
    return try parseEntry("\(Self.relationshipsFolderName)/\(Self.relationshipsFileExtension)", Relationships.self)

  } //end parseRelationships()


  /// Return an array of paths to relationships of type `officeDocument`
  public func parseDocumentPaths() throws -> [String] {

    //get relationships
    let relationships: Relationships = try parseRelationships()

    //get document paths
    return try self.parsePaths(of: .officeDocument, in: relationships)

  } //end parseDocumentPaths()


  /// Return an array of paths to relationships of the specified type
  public func parsePaths(of type: Relationship.SchemaType, in relationships: Relationships) throws -> [String] {

    //find paths
    return relationships.items
      .filter { $0.type == type }
      .map { $0.target }

  } //end parsePaths()


  public func parseStyles() throws -> Styles {
    decoder.keyDecodingStrategy = .useDefaultKeys

    return try parseEntry("xl/styles.xml", Styles.self)

  } //end parseStyles()


  public func parseSharedStrings() throws -> SharedStrings {
    decoder.keyDecodingStrategy = .useDefaultKeys

    return try parseEntry("xl/sharedStrings.xml", SharedStrings.self)

  } //end parseSharedStrings()


  // MARK: - Comment Processing

  private func buildCommentsPath(forWorksheet path: String) throws -> String {
    let pattern = "xl\\/worksheets\\/sheet(\\d+)[.]xml"
    let regex = try NSRegularExpression(pattern: pattern, options: [])
    let range = NSRange(location: 0, length: path.utf16.count)

    if let match = regex.firstMatch(in: path, options: [], range: range),
      let worksheetIdRange = Range(match.range(at: 1), in: path) {
      let worksheetId = path[worksheetIdRange]
      return "xl/comments\(worksheetId).xml"
    }

    throw CoreXLSXError.unsupportedWorksheetPath
  }

  public func parseComments(forWorksheet path: String) throws -> Comments {
    let commentsPath = try buildCommentsPath(forWorksheet: path)

    decoder.keyDecodingStrategy = .useDefaultKeys

    return try parseEntry(commentsPath, Comments.self)
  }


  // MARK: - Document Processing

  /// Parse and return an array of workbooks in this file.
  /// Worksheet names can be read as properties on the `Workbook` model type.
  public func parseWorkbooks() throws -> [Workbook] {
    let paths = try parseDocumentPaths()

    decoder.keyDecodingStrategy = .useDefaultKeys

    return try paths.map {
      try parseEntry($0, Workbook.self)
    }
  }

  /// Parse and return a workbook in this file.
  public func parseWorkbook(path: String) throws -> Workbook {
    try parseEntry(path, Workbook.self)
  }

//  public func parseTheme(path: String) throws -> Theme {
//    //decoder.keyDecodingStrategy = .useDefaultKeys
//    return try parseEntry(path, Theme.self)
//  }

  /** Return pairs of parsed document paths with corresponding relationships.

   **Deprecation warning**: this function doesn't handle root paths correctly,
   even though some XLSX files do contain root paths instead of relative
   paths. Use `parseDocumentRelationships(path:)` instead.
   */
  @available(*, deprecated, renamed: "parseDocumentRelationships(path:)")
  public func parseDocumentRelationships() throws -> [([Substring], Relationships)] {
    decoder.keyDecodingStrategy = .useDefaultKeys

    return try parseDocumentPaths()
      .compactMap { path -> ([Substring], Relationships)? in
        let originalComponents = path.split(separator: "/")
        var components = originalComponents

        components.insert(Self.relationshipsFolderName, at: 1)
        guard let filename = components.last else { return nil }
        components[components.count - 1] =
          Substring(filename.appending(Self.relationshipsFileExtension))

        let relationships = try parseEntry(
          components.joined(separator: "/"),
          Relationships.self
        )
        return (originalComponents, relationships)
      }

  } //end parseDocumentRelationships()


  /// Return parsed path with a parsed relationships model for a document at
  /// given path. Use `parseDocumentPaths` first to get a string path to pass
  /// as an argument to this function.
  public func parseDocumentRelationships(path: String) throws -> (Path, Relationships) {
    decoder.keyDecodingStrategy = .useDefaultKeys

    let originalPath = Path(path)
    guard let relationshipPath: Path = Self.relationshipPath(for: originalPath) else {
      throw CoreXLSXError.invalidDocumentPath
    }

    //parse entry
    let relationships = try parseEntry(relationshipPath.relativePath, Relationships.self )

    return (originalPath, relationships)

  } //end parseDocumentRelationships()


  // MARK: - Worksheet Processing

  /// Parse and return an array of worksheets in this XLSX file with their corresponding names.
  public func parseWorksheetPathsAndNames(workbook: Workbook) throws -> [(name: String?, path: String)] {

    return try parseDocumentPaths().map {
      try parseDocumentRelationships(path: $0)
    }.flatMap { (path, relationships) -> [(name: String?, path: String)] in
      let worksheets = relationships.items.filter { $0.type == .worksheet }

      guard !path.isRoot else { return worksheets.map { (name: nil, path: $0.target) } }

      // .rels file has paths relative to its directory,
      // storing that path in `pathPrefix`
      let pathPrefix = path.components.dropLast().joined(separator: "/")

      let sheetIDs = Dictionary(uniqueKeysWithValues: workbook.sheets.items.compactMap { sheet in
        sheet.name.flatMap { (sheet.relationship, $0) }
      })

      return worksheets.map { (name: sheetIDs[$0.id], path: "\(pathPrefix)/\($0.target)") }
    }

  } //end parseWorksheetPathsAndNames()


  /// Parse and return an array of worksheets in this XLSX file.
  public func parseWorksheetPaths() throws -> [String] {
    return try parseDocumentPaths().map {
      try parseDocumentRelationships(path: $0)
    }.flatMap { (path, relationships) -> [String] in
      let worksheets = relationships.items.filter { $0.type == .worksheet }

      guard !path.isRoot else { return worksheets.map { $0.target } }

      // .rels file has paths relative to its directory,
      // storing that path in `pathPrefix`
      let pathPrefix = path.components.dropLast().joined(separator: "/")

      return worksheets.map { "\(pathPrefix)/\($0.target)" }
    }

  } //end parseWorksheetPaths()


  /// Parse a worksheet at a given path contained in this XLSX file.
  public func parseWorksheet(at path: String) throws -> Worksheet {
    decoder.keyDecodingStrategy = .useDefaultKeys

    return try parseEntry(path, Worksheet.self)

  } //end parseWorksheet()


  /// Return all cells that are contained in a given worksheet and set of rows.
  @available(*, deprecated, renamed: "Worksheet.cells(atRows:)")
  public func cellsInWorksheet(at path: String, rows: [Int]) throws
    -> [Cell] {
    let ws = try parseWorksheet(at: path)

    return ws.data?.rows.filter { rows.contains(Int($0.reference)) }
      .reduce([]) { $0 + $1.cells } ?? []

  } //end cellsInWorksheet()


  /// Return all cells that are contained in a given worksheet and set of
  /// columns. This overloaded version is deprecated, you should pass
  /// an array of `ColumnReference` values as `columns` instead of an array
  /// of `String`s.
  @available(*, deprecated, renamed: "Worksheet.cells(atColumns:)")
  public func cellsInWorksheet(at path: String, columns: [String]) throws -> [Cell] {
    let ws = try parseWorksheet(at: path)

    return ws.data?.rows.map {
      let rowReference = $0.reference
      let targetReferences = columns
        .compactMap { (c: String) -> CellReference? in
          guard let columnReference = ColumnReference(c) else { return nil }
          return CellReference(columnReference, rowReference)
        }
      return $0.cells.filter { targetReferences.contains($0.reference) }
    }
    .reduce([]) { $0 + $1 } ?? []

  } //end cellsInWorksheet()

} //end class XLSXFile
