//
//  Series.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent an Excel chart data series.
public struct Series {

  let lxw_chart_series: UnsafeMutablePointer<lxw_chart_series>
  init(_ lxw_chart_series: UnsafeMutablePointer<lxw_chart_series>) {
    self.lxw_chart_series = lxw_chart_series
  }
  /// Set the name of a chart series range.
  @discardableResult public func set(name: String) -> Series {
    let _ = name.withCString { chart_series_set_name(lxw_chart_series, $0) }
    return self
  }

  /// The function is used to specify the the series marker:
  @discardableResult public func set(marker: Int, size: Int) -> Series {
    chart_series_set_marker_type(lxw_chart_series, UInt8(marker))
    chart_series_set_marker_size(lxw_chart_series, UInt8(size))
    return self
  }

  /// Set a series "values" range using row and column values.
  @discardableResult public func values(sheet: Worksheet, range: Range) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_values(lxw_chart_series, $0, range.row, range.col, range.row2, range.col2)
    }
    return self
  }
  /// Set a series "categories" range using row and column values.
  @discardableResult public func categories(sheet: Worksheet, range: Range) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_categories(
        lxw_chart_series, $0, range.row, range.col, range.row2, range.col2)
    }
    return self
  }
  /// Smooth a line or scatter chart series.
  @discardableResult public func set(smooth: Bool) -> Series {
    let smooth: UInt8 = smooth ? 1 : 0
    chart_series_set_smooth(lxw_chart_series, smooth)
    return self
  }
  /// Set a series name formula using row and column values.
  @discardableResult public func name(sheet: Worksheet, cell: Cell) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_name_range(lxw_chart_series, $0, cell.row, cell.col)
    }
    return self
  }
  /// Turn on a trendline for a chart data series.
  @discardableResult public func trendline(type: TrendlineType, value: Int = 2) -> Series {
    chart_series_set_trendline(lxw_chart_series, type.rawValue, UInt8(value))
    return self
  }
  /// Set the trendline name for a chart data series.
  @discardableResult public func trendline(name: String) -> Series {
    let _ = name.withCString { chart_series_set_trendline_name(lxw_chart_series, $0) }
    return self
  }
  /// Set the trendline line properties for a chart data series.
  @discardableResult public func trendline(
    color: Color = .black, width: Float = 2.25, dash_type: Int, transparency: Int = 0,
    hide: Bool = false
  )
    -> Series
  {
    var line = lxw_chart_line(
      color: Color.black.hex, none: (hide ? 1 : 0), width: width, dash_type: UInt8(dash_type),
      transparency: UInt8(transparency))
    chart_series_set_trendline_line(lxw_chart_series, &line)
    return self
  }
  /// Display the equation of a trendline for a chart data series.
  @discardableResult public func trendline_equation() -> Series {
    chart_series_set_trendline_equation(lxw_chart_series)
    return self
  }
  /// Display the R squared value of a trendline for a chart data series.
  @discardableResult public func trendline_r_squared() -> Series {
    chart_series_set_trendline_r_squared(lxw_chart_series)
    return self
  }
}
