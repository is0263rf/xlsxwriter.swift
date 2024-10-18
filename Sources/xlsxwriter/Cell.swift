//
//  Cell.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent an Excel cell.
public struct Cell {
    /// Integer data type to represent a row value.
    /// The maximum row in Excel is 1,048,576.
    let row: UInt32

    /// Integer data type to represent a column value.
    /// The maximum column in Excel is 16,384.
    let col: UInt16

    public init(row: UInt32, col: UInt16) {
        self.row = row
        self.col = col
    }

    /// Convert an Excel A1 cell string into a (row, col) pair.
    public init(_ value: String) {
        (self.row, self.col) = value.withCString { (lxw_name_to_row($0), lxw_name_to_col($0)) }
    }
}
