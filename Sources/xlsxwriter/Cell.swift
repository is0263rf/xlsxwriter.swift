//
//  Cell.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public struct Cell: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
  let row: UInt32, col: UInt16
  public init(stringLiteral value: String) {
    (self.row, self.col) = value.withCString { (lxw_name_to_row($0), lxw_name_to_col($0)) }
  }
  public init(arrayLiteral elements: Int...) {
    precondition(elements.count == 2, "[row, col]")
    self.row = UInt32(elements[0])
    self.col = UInt16(elements[1])
  }

  public init(_ row: UInt32, _ col: UInt16) {
    self.row = row
    self.col = col
  }
}
