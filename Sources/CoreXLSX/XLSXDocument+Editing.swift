//
//  XLSXDocument+Editing.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

extension XLSXDocument {
  // MARK: - Editing Functions

  @discardableResult public func createWorkbook(named: String) -> Workbook {
    // create workbook
    let path: Path = Self.excelPath.path(byAppending: "\(named).xml")
    let workbook = Workbook()
    workbooksMap.append((path: path, book: workbook))
    relationships.addRelationship(with: .officeDocument, target: path.value)
    return workbook
  } // end createWorkbook()

  @discardableResult public func createWorksheet(named: String, for workbook: Workbook) -> Worksheet? {
    // find workbook path
    guard let workbookPath: Path = workbooksMap.first(where: { $0.book == workbook })?.path else {
      print("Workbook does not belong to this document")
      return nil
    }

    // create new worksheet
    let workSheetPath: Path = Self.worksheetsPath.path(byAppending: "\(named).xml")
    let worksheet = Worksheet()
    worksheets.append((path: workSheetPath, sheet: worksheet))

    // TODO: add worksheet to workbook

    // add worksheet relationship
    addDocumentRelationship(for: workbookPath, with: .worksheet, target: workSheetPath.value)

    return worksheet
  } // end createWorksheet()

  internal func addDocumentRelationship(for relationshipsPath: Path, with type: Relationship.SchemaType, target: String) {
    // find existing relationships
    if let index: Int = documentRelationships.firstIndex(where: { $0.path == relationshipsPath }) {
      documentRelationships[index].relations.addRelationship(with: type, target: target)
    }
    // create new relationships
    else {
      var relationships = Relationships()
      relationships.addRelationship(with: type, target: target)
      documentRelationships.append((path: relationshipsPath, relations: relationships))
    }
  } // end addDocumentRelationship()
} // end extension XLSXDocument
