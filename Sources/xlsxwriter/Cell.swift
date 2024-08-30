//
//  Cell.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public struct Cell {
  let row: UInt32
  let col: UInt16

  public init(row: UInt32, col: UInt16) {
    self.row = row
    self.col = col
  }

  public init(_ value: String) {
    (self.row, self.col) = value.withCString { (lxw_name_to_row($0), lxw_name_to_col($0)) }
  }
}
