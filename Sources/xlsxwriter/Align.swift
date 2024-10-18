//
//  Align.swift
//
//
//  Created by Yoshinori Takada on 2024/08/30.
//

import Foundation

/// Alignment values
public enum Align: UInt8 {
    /// Left horizontal alignment
    case left

    /// Center horizontal alignment
    case center

    /// Right horizontal alignment
    case right

    /// Center fill horizontal alignment
    case fill

    /// Justify horizontal alignment
    case justify

    /// Center Across horizontal alignment
    case centerAcross

    /// Left horizontal alignment
    case distributed

    /// Top vertical alignment
    case verticalTop

    /// Bottom vertical alignment
    case verticalBottom

    /// Center vertical alignment
    case verticalCenter

    /// Justify vertical alignment
    case verticalJustify

    /// Distributed vertical alignment
    case verticalDistributed
}
