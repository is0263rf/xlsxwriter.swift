//
//  Cols.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public struct Cols {
  let col: UInt16
  let col2: UInt16

  public init(_ cols: String) {
    (self.col, self.col2) = cols.withCString { (lxw_name_to_col($0), lxw_name_to_col_2($0)) }
  }

  public init(_ col: UInt16, _ col2: UInt16) {
    self.col = col
    self.col2 = col2
  }
}
