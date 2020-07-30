//
//  XLSXDocumentDelegate.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 30/7/20.
//

import Foundation

public protocol XLSXDocumentDelegate: NSObjectProtocol {
  func didReceiveError(for document: XLSXDocument, error: Error)
}
