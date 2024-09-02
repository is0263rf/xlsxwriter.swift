//
//  TrendlineType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Series trendline/regression types.
public enum TrendlineType: UInt8 {
  /// Trendline type: Linear.
  case linear

  /// Trendline type: Logarithm.
  case log

  /// Trendline type: Polynomial.
  case poly

  /// Trendline type: Power.
  case power

  /// Trendline type: Exponential.
  case exp

  /// Trendline type: Moving Average.
  case average
}
