//
//  Border.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Cell border styles for use with format.set_border()
public enum Border {
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

    public var rawValue: lxw_format_borders {
        switch self {
        case .none:
            return LXW_BORDER_NONE
        case .thin:
            return LXW_BORDER_THIN
        case .medium:
            return LXW_BORDER_MEDIUM
        case .dashed:
            return LXW_BORDER_DASHED
        case .dotted:
            return LXW_BORDER_DOTTED
        case .thick:
            return LXW_BORDER_THICK
        case .double:
            return LXW_BORDER_DOUBLE
        case .hair:
            return LXW_BORDER_HAIR
        case .mediumDashed:
            return LXW_BORDER_MEDIUM_DASHED
        case .dashDot:
            return LXW_BORDER_DASH_DOT
        case .mediumDashDot:
            return LXW_BORDER_MEDIUM_DASH_DOT
        case .dashDotDot:
            return LXW_BORDER_DASH_DOT_DOT
        case .mediumDashDotDot:
            return LXW_BORDER_MEDIUM_DASH_DOT_DOT
        case .slantDashDot:
            return LXW_BORDER_SLANT_DASH_DOT
        }
    }
}
