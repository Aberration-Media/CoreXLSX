//
//  SharedStrings+Encoding.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 14/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension RichText.Size: AttributeType {}
extension RichText.Color: AttributeType {}
extension RichText.Font: AttributeType {}
extension RichText.Bold: AttributeType {}
extension RichText.Italic: AttributeType {}
