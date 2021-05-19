//
//  XLSXDocument+Editing.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

extension XLSXDocument {

  // MARK: - Document Functions

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
    worksheetsMap.append((path: workSheetPath, sheet: worksheet))

    // TODO: add worksheet to workbook

    // add worksheet relationships
    addDocumentRelationship(for: workbookPath, with: .worksheet, target: workSheetPath.value)
    addDocumentRelationship(for: workbookPath, with: .sharedStrings, target: Self.sharedStringsPath.value)

    return worksheet
  } // end createWorksheet()


  @discardableResult internal func addDocumentRelationship(for relationshipsPath: Path, with type: Relationship.SchemaType, target: String) -> Relationship {
    let relationship: Relationship

    // find existing relationships
    if let index: Int = documentRelationships.firstIndex(where: { $0.path == relationshipsPath }) {
      relationship = documentRelationships[index].relations.addRelationship(with: type, target: target)
    }
    // create new relationships
    else {
      var relationships = Relationships()
      relationship = relationships.addRelationship(with: type, target: target)
      documentRelationships.append((path: relationshipsPath, relations: relationships))
    }

    return relationship

  } // end addDocumentRelationship()



  // MARK: - Editing Functions

  public func modifyWorksheet(at index: Int, modifications: (inout Worksheet, inout SharedStrings) -> Void) {

    //valid index
    if index >= 0 && index < self.worksheetsMap.count {
      modifications(&self.worksheetsMap[index].sheet, &self.sharedStrings)
    }

  } //end modifyWorksheet()

} // end extension XLSXDocument
