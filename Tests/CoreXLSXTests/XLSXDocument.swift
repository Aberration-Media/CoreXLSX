//
//  XLSXDocument.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

@testable import CoreXLSX
import Foundation
import XCTest

public class XLSXDocumentTest: XCTestCase, XLSXDocumentDelegate {


  // MARK: Test Properties

  /// ouput folder URL for saved documents
  private let outputFolderURL: URL = {
    URL(fileURLWithPath: currentWorkingPath).deletingLastPathComponent().appendingPathComponent("XLSXTestsOutput")
  }()



  // MARK: - Configuration Functions

  override public func setUp() {
    // ensure output folder exists
    let manager = FileManager.default
    var isDirectory: ObjCBool = false
    if !(manager.fileExists(atPath: outputFolderURL.path, isDirectory: &isDirectory) && isDirectory.boolValue) {
      do {
        try manager.createDirectory(at: outputFolderURL, withIntermediateDirectories: false, attributes: nil)
      } catch {
        fatalError("Could not create output directory(\(outputFolderURL)): \(error)")
      }
    } // end if (folder does not exist)
  } // end setUp()



  // MARK: - XLSX Write Test Functions

  func testSaveEmptyDocument() {

    // create blank document
    let document = XLSXDocument()
    document.documentDelegate = self

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent("Empty.xlsx")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }

  } // end testSaveEmptyDocument()


  func testSaveNewDocument() {

    // create document
    let document = XLSXDocument()
    document.documentDelegate = self

    //edit document
    document.modifyWorksheet(at: 0) {
      (worksheet: inout Worksheet, sharedStrings: inout SharedStrings) in

      //add new row
      let styles: [String] = []
      worksheet.addRow(with: ["abc", "def", "hij", "klm"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.addRow(with: ["1", "2", "3", "4"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.insertRow(at: 6, with: ["in1", "in2"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.updateRowValues(at: 4, with: ["replace RICH"], sharedStrings: &sharedStrings)
      worksheet.updateColumnValues(at: 5, row: 2, with: ["test1", "test2", "test3", "test4", "test5", "test6"], sharedStrings: &sharedStrings)
      worksheet.addRow(with: ["A", "B", "C", "D"], sharedStrings: &sharedStrings, styles: styles)


    } //end modify closure

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent("NewDocument.xlsx")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }

  } // end testSaveNewDocument()



  func testSaveExistingDocument() {

    // open test document
    let fileName: String = "Dates.xlsx"
    //let fileName: String = "jewelershealthcare.com-census.1.xlsx"
    //let fileName: String = "categories.xlsx"
    guard let file =
      XLSXFile(filepath: "\(currentWorkingPath)/\(fileName)") else {
      XCTAssert(false, "failed to open the file")
      return
    }

    // create document
    let document = XLSXDocument(with: file)
    document.documentDelegate = self

    do {

      let filePath: URL = outputFolderURL.appendingPathComponent("ExistingSaved.xlsx")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }
  } // end testSaveExistingDocument()



  func testSaveEditedDocument() {

    let fileName: String = "Dates.xlsx"
    guard let file =
      XLSXFile(filepath: "\(currentWorkingPath)/\(fileName)") else {
      XCTAssert(false, "failed to open the file")
      return
    }

    // create document
    let document = XLSXDocument(with: file)
    document.documentDelegate = self
    document.modifyWorksheet(at: 0) { ( worksheet: inout Worksheet, sharedStrings: inout SharedStrings) in

      //add new row
      let styles = worksheet.lastRow?.styles()
      worksheet.addRow(with: ["abc", "def", "hij", "klm"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.insertRow(at: 3, with: ["in1", "in2"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.deleteRows(in: 0..<3)
      worksheet.updateRowValues(at: 4, with: ["replace RICH"], sharedStrings: &sharedStrings)
      worksheet.updateColumnValues(at: 5, row: 2, with: ["test1", "test2", "test3", "test4", "test5", "test6"], sharedStrings: &sharedStrings)
      worksheet.deleteColumns(in: 0..<1)

    } //end modify closure

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent("EditedExisting.xlsx")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }
  } // end testSaveEmptyDocument()



  // MARK: - XLSXDocumentDelegate Functions

  public func didReceiveError(for document: XLSXDocument, error: Error) {
    XCTAssert(false, "Error retrieving file data: \(error)")
  }

} // end class XLSXDocumentTest
