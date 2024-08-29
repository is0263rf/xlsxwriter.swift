//
//  HorizontalAlignment.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Alignment values for format.set(alignment:)
public enum HorizontalAlignment: UInt8 {
  case none = 0
  case left
  case center
  case right
  case fill
  case justify
  case center_across
  case distributed
}
