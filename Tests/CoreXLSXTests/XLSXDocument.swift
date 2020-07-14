//
//  XLSXDocument.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

@testable import CoreXLSX
import Foundation
import XCTest

public class XLSXDocumentTest: XCTestCase {
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

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent("Empty.xlsx")
      print("output path: \(filePath)")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }
  } // end testSaveEmptyDocument()

  func testSaveExistingDocument() {
    // open test document
    let fileName: String = "jewelershealthcare.com-census.1.xlsx"
    guard let file =
      XLSXFile(filepath: "\(currentWorkingPath)/\(fileName)") else {
      XCTAssert(false, "failed to open the file")
      return
    }

    // create document
    let document = XLSXDocument(with: file)

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent(fileName)
      print("output path: \(filePath)")
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
    document.modifyWorksheet(at: 0) { ( worksheet: inout Worksheet, sharedStrings: inout SharedStrings) in

      //add new row
      let styles = worksheet.lastRow?.styles()
      worksheet.addRow(with: ["abc", "def", "hij", "klm"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.insertRow(at: 3, with: ["in1", "in2"], sharedStrings: &sharedStrings, styles: styles)
      worksheet.deleteRows(in: 0..<3)
      worksheet.updateRowValues(at: 4, with: ["replace RICH"], sharedStrings: &sharedStrings)
    }

//    //debug cells
//    for worksheet in document.worksheets {
//      for row in worksheet.data?.rows ?? [] {
//        for c in row.cells {
//          print("CELL: \(c)")
//        }
//      }
//    } //end for (worksheets)

    do {
      let filePath: URL = outputFolderURL.appendingPathComponent("Test.xlsx")
      print("output path: \(filePath)")
      try document.save(to: filePath.path, overwrite: true)

    } catch {
      XCTAssert(false, "Error saving file: \(error)")
    }
  } // end testSaveEmptyDocument()
} // end class XLSXDocumentTest
