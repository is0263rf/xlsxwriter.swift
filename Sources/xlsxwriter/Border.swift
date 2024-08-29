//
//  Border.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Cell border styles for use with format.set_border()
public enum Border: UInt8 {
  case noBorder
  case thin
  case medium
  case dashed
  case dotted
  case thick
  case double
  case hair
  case medium_dashed
  case dash_dot
  case medium_dash_dot
  case dash_dot_dot
  case medium_dash_dot_dot
  case slant_dash_dot
}
