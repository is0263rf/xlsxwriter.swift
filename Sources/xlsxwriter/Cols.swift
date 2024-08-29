//
//  Cols.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public struct Cols: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
  let col: UInt16, col2: UInt16
  public init(stringLiteral value: String) {
    (self.col, self.col2) = value.withCString { (lxw_name_to_col($0), lxw_name_to_col_2($0)) }
  }
  public init(arrayLiteral elements: Int...) {
    precondition(elements.count == 2, "[col, col2]")
    self.col = UInt16(elements[0])
    self.col2 = UInt16(elements[1])
  }
  init(_ col: UInt16, _ col2: UInt16) {
    self.col = col
    self.col2 = col2
  }
}
