//
//  TotalFunction.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

public enum TotalFunction: UInt8, ExpressibleByIntegerLiteral {
  case none = 0
  /** Use the average function as the table total. */
  case average = 101
  /** Use the count numbers function as the table total. */
  case nums = 102
  /** Use the count function as the table total. */
  case count = 103
  /** Use the max function as the table total. */
  case max = 104
  /** Use the min function as the table total. */
  case min = 105
  /** Use the standard deviation function as the table total. */
  case std_dev = 107
  /** Use the sum function as the table total. */
  case sum = 109

  public init(integerLiteral value: Int) {
    if let function = TotalFunction(rawValue: UInt8(value)) {
      self = function
    } else {
      self = .none
    }
  }
}
