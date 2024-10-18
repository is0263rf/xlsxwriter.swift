//
//  Underline.swift
//
//
//  Created by Yoshinori Takada on 2024/08/30.
//

import libxlsxwriter

/// Format underline values
public enum Underline {
    /// Single underline
    case single

    /// Double underline
    case double

    /// Single accounting underline
    case singleAccounting

    /// Double accounting underline
    case doubleAccounting

    public var rawValue: lxw_format_underlines {
        switch self {
        case .single:
            return LXW_UNDERLINE_SINGLE
        case .double:
            return LXW_UNDERLINE_DOUBLE
        case .singleAccounting:
            return LXW_UNDERLINE_SINGLE_ACCOUNTING
        case .doubleAccounting:
            return LXW_UNDERLINE_DOUBLE_ACCOUNTING
        }
    }
}
