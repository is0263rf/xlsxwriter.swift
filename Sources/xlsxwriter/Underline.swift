//
//  Underline.swift
//
//
//  Created by Yoshinori Takada on 2024/08/30.
//

import libxlsxwriter

/// Format underline values
public enum Underline: UInt8 {
  /// Single underline
  case single

  /// Double underline
  case double

  /// Single accounting underline
  case singleAccounting

  /// Double accounting underline
  case doubleAccounting
}
