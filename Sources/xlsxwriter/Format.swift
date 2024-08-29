//
//  Format.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent the formatting properties of an Excel format.
public struct Format {
  let lxwFormat: UnsafeMutablePointer<lxw_format>
  init(_ lxw_format: UnsafeMutablePointer<lxw_format>) { self.lxwFormat = lxw_format }
  /// Turn on bold for the format font.
  @discardableResult public func bold() -> Format {
    format_set_bold(lxwFormat)
    return self
  }
  /// Turn on italic for the format font.
  @discardableResult public func italic() -> Format {
    format_set_italic(lxwFormat)
    return self
  }
  /// Set the cell border style.
  @discardableResult public func border(style: Border) -> Format {
    format_set_border(lxwFormat, style.rawValue)
    return self
  }
  /// Set the cell top border style.
  @discardableResult public func top(style: Border) -> Format {
    format_set_top(lxwFormat, style.rawValue)
    return self
  }
  /// Set the cell bottom border style.
  @discardableResult public func bottom(style: Border) -> Format {
    format_set_bottom(lxwFormat, style.rawValue)
    return self
  }
  /// Set the cell left border style.
  @discardableResult public func left(style: Border) -> Format {
    format_set_left(lxwFormat, style.rawValue)
    return self
  }
  /// Set the cell right border style.
  @discardableResult public func right(style: Border) -> Format {
    format_set_right(lxwFormat, style.rawValue)
    return self
  }
  /// Set the horizontal alignment for data in the cell.
  @discardableResult public func align(horizontal: HorizontalAlignment) -> Format {
    format_set_align(lxwFormat, horizontal.rawValue)
    return self
  }
  /// Set the vertical alignment for data in the cell.
  @discardableResult public func align(vertical: VerticalAlignment) -> Format {
    format_set_align(lxwFormat, vertical.rawValue)
    return self
  }
  /// Set the vertical alignment and horizontal alignment to center.
  @discardableResult public func center() -> Format {
    format_set_align(lxwFormat, HorizontalAlignment.center.rawValue)
    format_set_align(lxwFormat, VerticalAlignment.center.rawValue)
    return self
  }
  /// Set the number format for a cell.
  @discardableResult public func set(num_format: String) -> Format {
    num_format.withCString { format_set_num_format(lxwFormat, $0) }
    return self
  }
  /// Set the Excel built-in number format for a cell.
  @discardableResult public func set(num_format index: Int) -> Format {
    format_set_num_format_index(lxwFormat, UInt8(index))
    return self
  }
  /// Turn on the text "shrink to fit" for a cell.
  @discardableResult public func shrink() -> Format {
    format_set_shrink(lxwFormat)
    return self
  }
  /// Set the font used in the cell.
  @discardableResult public func font(name: String) -> Format {
    name.withCString { format_set_font_name(lxwFormat, $0) }
    return self
  }
  /// Set the color of the cell border.
  @discardableResult public func border(color: Color) -> Format {
    format_set_border_color(lxwFormat, color.hex)
    return self
  }
  /// Set the color of the font used in the cell.
  @discardableResult public func font(color: Color) -> Format {
    format_set_font_color(lxwFormat, color.hex)
    return self
  }
  /// Set the pattern background color for a cell.
  @discardableResult public func background(color: Color) -> Format {
    format_set_pattern(lxwFormat, 1)
    format_set_bg_color(lxwFormat, color.hex)
    return self
  }
  /// Set the rotation of the text in a cell.
  @discardableResult public func rotation(angle: Int) -> Format {
    format_set_rotation(lxwFormat, Int16(angle))
    return self
  }
  /// Set the size of the font used in the cell.
  @discardableResult public func font(size: Double) -> Format {
    format_set_font_size(lxwFormat, size)
    return self
  }
  /// Set the text wrap to cell. This is required which cell's text contains line break to show correctly.
  @discardableResult public func setTextWrap() -> Format {
    format_set_text_wrap(lxwFormat)
    return self
  }
}
