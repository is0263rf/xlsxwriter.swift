//
//  Border.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Cell border styles for use with format.set_border()
public enum Border: UInt8 {
  /// No border
  case none

  /// Thin border style
  case thin

  /// Medium border style
  case medium

  /// Dashed border style
  case dashed

  /// Dotted border style
  case dotted

  /// Thick border style
  case thick

  /// Double border style
  case double

  /// Hair border style
  case hair

  /// Medium dashed border style
  case mediumDashed

  /// Dash-dot border style
  case dashDot

  /// Medium dash-dot border style
  case mediumDashDot

  /// Dash-dot-dot border style
  case dashDotDot

  /// Medium dash-dot-dot border style
  case mediumDashDotDot

  /// Slant dash-dot border style
  case slantDashDot
}
