//
//  Path+Extension.swift
//  CoreXLSXmacOS
//
//  Created by Daniel Welsh on 9/7/20.
//

import Foundation

extension Path: Equatable {

  // MARK: Convenience Properties

  /// evalutated path components (processes` up folder` `"../"` components)
  public var evaluatedPathComponents: [Substring] {

    var adjustedComponents: [Substring] = []
    for component in self.components {
      //found 'level up' operator
      if component == ".." {
        if !adjustedComponents.isEmpty {
          adjustedComponents.removeLast(1)
        }
      }
      //standard path component
      else {
        adjustedComponents.append(component)
      }
    }

    return adjustedComponents
    
  }
  

  /// retrieve component after the final path separator('/') if one exists
  public var lastPathComponent: String? {
    if let last: Substring = components.last {
      return String(last)
    }
    return nil
  }

  /// the directory path for the current path item
  public var directoryPath: String {

    //default to root path
    var path = "/"

    //find directory path
    if self.components.count > 1 {

      //get directory path
      path = self.components[0..<(self.components.count - 1)].joined(separator: "/")

      //add directory separator
      if path.last != "/" {
        path += "/"
      }
    }

    return path
  }

  ///path as a relative locator (exlcudes opening '/' character)
  public var relativePath: String {
    return self.evaluatedPathComponents.joined(separator: "/")
  }
  
  ///path as a absolute locator (enforces opening '/' character)
  public var absolutePath: String {
    return "/" + self.evaluatedPathComponents.joined(separator: "/")
  }

  // MARK: - Configuration Functions

  public init(rootPath: String) {
    self.init(rootPath.first != "/" ? "/\(rootPath)" : rootPath)
  }

  // MARK: - Util Functions

  /**
   Create new Path with the specified string as the last path component

   - parameters:
      - pathComponent: the component to append to the end of the current path
   */
  public func path(byAppending pathComponent: String) -> Path {
    // compile new path string
    var path: String = value

    // strip white space
    let component: String = pathComponent.trimmingCharacters(in: .whitespacesAndNewlines)

    // valid component
    if !component.isEmpty {
      // check for separator
      if component.first == "/" {
        // duplicate separator
        if path.last == "/" {
          // strip separator
          path += component.dropFirst()
        }
        // separator already exists
        else if path.last == "/" {
          path += component
        }
      }
      // separator already exists
      else if path.last == "/" {
        path += component
      }
      // append separator
      else {
        path += "/\(component)"
      }
    } // end if (valid component)

    return Path(path)
  } // end path(byAppending:)

  // MARK: - Equatable Functions

  public static func ==(lhs: Self, rhs: Self) -> Bool {
    return lhs.value == rhs.value
  }
} // end extension Path
