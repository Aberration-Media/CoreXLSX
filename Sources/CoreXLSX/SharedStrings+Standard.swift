//
//  SharedStrings+Standard.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation
import XMLCoder

public extension SharedStrings {
  //
  static var standard: SharedStrings = {
    let stringsXML: String = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst uniqueCount=\"0\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>"

    let decoder = XMLDecoder()
    decoder.trimValueWhitespaces = false
    decoder.shouldProcessNamespaces = true
    do {
      return try decoder.decode(SharedStrings.self, from: stringsXML.data(using: .utf8) ?? Data())

    } catch {
      fatalError("Could not create standard Shared Strings: \(error)")
    }

  }()
} // end extension SharedStrings
