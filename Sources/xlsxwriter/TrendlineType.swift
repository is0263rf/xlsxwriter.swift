//
//  TrendlineType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Series trendline/regression types.
public enum TrendlineType {
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

    public var rawValue: lxw_chart_trendline_type {
        switch self {
        case .linear:
            return LXW_CHART_TRENDLINE_TYPE_LINEAR
        case .log:
            return LXW_CHART_TRENDLINE_TYPE_LOG
        case .poly:
            return LXW_CHART_TRENDLINE_TYPE_POLY
        case .power:
            return LXW_CHART_TRENDLINE_TYPE_POWER
        case .exp:
            return LXW_CHART_TRENDLINE_TYPE_EXP
        case .average:
            return LXW_CHART_TRENDLINE_TYPE_AVERAGE
        }
    }
}
