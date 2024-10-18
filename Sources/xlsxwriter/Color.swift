//
//  Color.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Structure for color which contains common colors.
public struct Color {
    public var hex: UInt32
    public init(hex: UInt32) {
        self.hex = hex
    }

    public static var black: Self = Self(hex: 0x1000000)
    public static var blue: Self = Self(hex: 0x0000FF)
    public static var brown: Self = Self(hex: 0x800000)
    public static var cyan: Self = Self(hex: 0x00FFFF)
    public static var gray: Self = Self(hex: 0x808080)
    public static var green: Self = Self(hex: 0x008000)
    public static var lime: Self = Self(hex: 0x00FF00)
    public static var magenta: Self = Self(hex: 0xFF00FF)
    public static var navy: Self = Self(hex: 0x000080)
    public static var orange: Self = Self(hex: 0xFF6600)
    public static var purple: Self = Self(hex: 0x800080)
    public static var red: Self = Self(hex: 0xFF0000)
    public static var silver: Self = Self(hex: 0xC0C0C0)
    public static var white: Self = Self(hex: 0xFFFFFF)
    public static var yellow: Self = Self(hex: 0xFFFF00)
}
