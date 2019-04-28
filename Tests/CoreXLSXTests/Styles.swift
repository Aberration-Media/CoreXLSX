//
//  Styles.swift
//  CoreXLSX
//
//  Created by Max Desiatov on 16/03/2019.
//

@testable import CoreXLSX
import XCTest
import XMLCoder

private let numberFormats = NumberFormats(
  items: [
    NumberFormat(
      id: 0,
      formatCode: "General"
    ),
  ], count: 1
)

private let fonts = Fonts(
  items: [
    Font(
      size: Font.Size(value: 10),
      color: Color(indexed: 8, auto: nil, rgb: nil),
      name: Font.Name(value: "Helvetica Neue"),
      bold: nil,
      italic: nil,
      strike: nil
    ),
    Font(
      size: Font.Size(value: 12),
      color: Color(indexed: 8, auto: nil, rgb: nil),
      name: Font.Name(value: "Helvetica Neue"),
      bold: nil,
      italic: nil,
      strike: nil
    ),
    Font(
      size: Font.Size(value: 10),
      color: Color(indexed: 8, auto: nil, rgb: nil),
      name: Font.Name(value: "Helvetica Neue"),
      bold: Font.Bold(value: true),
      italic: nil,
      strike: nil
    ),
  ], count: 3
)

private let fills = Fills(items: [
  Fill(patternFill: PatternFill(
    patternType: "none",
    foregroundColor: nil,
    backgroundColor: nil
  )),
  Fill(patternFill: PatternFill(
    patternType: "gray125",
    foregroundColor: nil,
    backgroundColor: nil
  )),
  Fill(patternFill: PatternFill(
    patternType: "solid",
    foregroundColor: Color(indexed: 9, auto: nil, rgb: nil),
    backgroundColor: Color(indexed: nil, auto: 1, rgb: nil)
  )),
  Fill(patternFill: PatternFill(
    patternType: "solid",
    foregroundColor: Color(indexed: 11, auto: nil, rgb: nil),
    backgroundColor: Color(indexed: nil, auto: 1, rgb: nil)
  )),
  Fill(patternFill: PatternFill(
    patternType: "solid",
    foregroundColor: Color(indexed: 13, auto: nil, rgb: nil),
    backgroundColor: Color(indexed: nil, auto: 1, rgb: nil)
  )),
], count: 5)

final class StylesTests: XCTestCase {
  func testStyles() throws {
    guard let file =
      XLSXFile(filepath: "\(currentWorkingPath)/categories.xlsx") else {
      XCTAssert(false, "failed to open the file")
      return
    }

    let styles = try file.parseStyles()

    XCTAssertEqual(styles.numberFormats!, numberFormats)
    XCTAssertEqual(styles.fills!, fills)
    XCTAssertEqual(styles.fonts!, fonts)
    XCTAssertEqual(styles.borders!.count, 13)
    XCTAssertEqual(styles.cellStyleFormats!.count, 1)
    XCTAssertEqual(styles.cellFormats!.count, 28)
    XCTAssertEqual(styles.cellStyles!.count, 1)
    XCTAssertEqual(styles.differentialFormats!.count, 0)
    XCTAssertEqual(styles.tableStyles!.count, 0)
    XCTAssertEqual(styles.colors!.indexed.rgbColors.count, 14)
  }

  static let allTests = [
    ("testStyles", testStyles),
  ]
}