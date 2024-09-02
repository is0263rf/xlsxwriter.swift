//
//  Range.swift
//
//
//  Created by Daniel Müllenborn on 02.01.21.
//

import libxlsxwriter

public struct Range: ExpressibleByStringLiteral, ExpressibleByArrayLiteral {
  let row: UInt32
  let col: UInt16
  let row2: UInt32
  let col2: UInt16

  /// Convert an Excel A1:B2 range into a (first_row, first_col, last_row, last_col) sequence.
  public init(stringLiteral value: String) {
    (self.row, self.col, self.row2, self.col2) = value.withCString {
      (lxw_name_to_row($0), lxw_name_to_col($0), lxw_name_to_row_2($0), lxw_name_to_col_2($0))
    }
  }

  public init(arrayLiteral elements: Int...) {
    precondition(elements.count == 4, "[row, col, row2, col2]")
    self.row = UInt32(elements[0])
    self.col = UInt16(elements[1])
    self.row2 = UInt32(elements[2])
    self.col2 = UInt16(elements[3])
  }
}
