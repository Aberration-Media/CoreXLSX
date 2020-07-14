//
//  Relationships+Extensions.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 10/7/20.
//

import Foundation

// enable model attributes to be encoded correctly
extension Relationship: AttributeType {}

public extension Relationships {
  // MARK: - Initialization Functions

  /// create empty relationships object
  init() {
    items = []
  }

  // MARK: - Editing Functions

  /**
   Add a relationship

    - parameters:
       - type: type of relationship
       - target: target document path for the relationship
      */
  @discardableResult mutating func addRelationship(with type: Relationship.SchemaType, target: String) -> Relationship {
    // create new relationship
    let relationshipId = "rId\(items.count + 1)"
    let relationship = Relationship(id: relationshipId, type: type, target: target)
    items.append(relationship)

    return relationship

  } // end addRelationship()

} // end extension Relationship
