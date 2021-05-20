//
//  XLSXDocument+Editing.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

extension XLSXDocument {

  // MARK: - Document Creation Functions

  @discardableResult public func createWorkbook(named: String) -> Path {

    // create workbook
    let path: Path = Self.excelPath.path(byAppending: "\(named).xml")
    let workbook = Workbook()
    workbooksMap.append((path: path, book: workbook))
    relationships.addRelationship(with: .officeDocument, target: path.value)

    return path

  } // end createWorkbook()


  @discardableResult public func createWorksheet(named: String, for workbook: inout Workbook) -> Worksheet? {

    // find workbook path
    guard let workbookPath: Path = workbooksMap.first(where: { $0.book == workbook })?.path else {
      print("Workbook does not belong to this document")
      return nil
    }

    // create new worksheet
    let absoluteWorkSheetPath: Path = Self.absoluteWorksheetsPath.path(byAppending: "\(named).xml")
    let relativeWorkSheetPath: Path = Self.relativeWorksheetsPath.path(byAppending: "\(named).xml")
    let worksheet = Worksheet()
    worksheetsMap.append((path: absoluteWorkSheetPath, sheet: worksheet))

    // add worksheet relationships
    let sheetRelationship = addDocumentRelationship(for: workbookPath, with: .worksheet, target: relativeWorkSheetPath.value)
    addDocumentRelationship(for: workbookPath, with: .sharedStrings, target: Self.sharedStringsFileName)

    // add worksheet to workbook
    workbook.addSheet(named: named, relationshipId: sheetRelationship.id)
    
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



  // MARK: - Document Retrieval Functions

  internal func indexForWorkbook(with path: Path) -> Int? {

    //find index of first workbook with a matching path
    return self.workbooksMap.firstIndex {
      (bookPath: Path, _: Workbook) in
      
      return bookPath == path
    }

  } //end indexForWorkbook()
  

  // MARK: - Editing Functions

  public func modifyWorksheet(at index: Int, modifications: (inout Worksheet, inout SharedStrings) -> Void) {

    //valid index
    if index >= 0 && index < self.worksheetsMap.count {
      modifications(&self.worksheetsMap[index].sheet, &self.sharedStrings)
    }

  } //end modifyWorksheet()

} // end extension XLSXDocument
