//
//  LegendPosition.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Chart legend positions.
public enum LegendPosition: UInt8 {
  /// No chart legend.
  case none = 0

  /// Chart legend positioned at right side.
  case right

  /// Chart legend positioned at left side.
  case left

  /// Chart legend positioned at top.
  case top

  /// Chart legend positioned at bottom.
  case bottom

  /// Chart legend positioned at top right.
  case topRight

  /// Chart legend positioned at right side.
  case overlayRight

  /// Chart legend positioned at left side.
  case overlayLeft

  /// Chart legend positioned at top right.
  case overlayTopRight
}
