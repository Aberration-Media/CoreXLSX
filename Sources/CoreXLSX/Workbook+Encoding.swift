//
//  Workbook+Encoding.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension Workbook.Properties: AttributeType {}
extension Workbook.Sheet: AttributeType {}
extension Workbook.View: AttributeType {}
