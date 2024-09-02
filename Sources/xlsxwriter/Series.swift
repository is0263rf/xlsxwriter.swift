//
//  Series.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent an Excel chart data series.
public struct Series {

  let lxwChartSeries: UnsafeMutablePointer<lxw_chart_series>
  init(_ lxw_chart_series: UnsafeMutablePointer<lxw_chart_series>) {
    self.lxwChartSeries = lxw_chart_series
  }
  /// Set the name of a chart series range.
  @discardableResult public func set(name: String) -> Series {
    let _ = name.withCString { chart_series_set_name(lxwChartSeries, $0) }
    return self
  }

  /// The function is used to specify the the series marker:
  @discardableResult public func set(marker: Int, size: Int) -> Series {
    chart_series_set_marker_type(lxwChartSeries, UInt8(marker))
    chart_series_set_marker_size(lxwChartSeries, UInt8(size))
    return self
  }

  /// Set a series "values" range using row and column values.
  @discardableResult public func values(sheet: Worksheet, range: Range) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_values(lxwChartSeries, $0, range.row, range.col, range.row2, range.col2)
    }
    return self
  }
  /// Set a series "categories" range using row and column values.
  @discardableResult public func categories(sheet: Worksheet, range: Range) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_categories(
        lxwChartSeries, $0, range.row, range.col, range.row2, range.col2)
    }
    return self
  }
  /// Smooth a line or scatter chart series.
  @discardableResult public func set(smooth: Bool) -> Series {
    let smooth: UInt8 = smooth ? 1 : 0
    chart_series_set_smooth(lxwChartSeries, smooth)
    return self
  }
  /// Set a series name formula using row and column values.
  @discardableResult public func name(sheet: Worksheet, cell: Cell) -> Series {
    let _ = sheet.name.withCString {
      chart_series_set_name_range(lxwChartSeries, $0, cell.row, cell.col)
    }
    return self
  }
  /// Turn on a trendline for a chart data series.
  @discardableResult public func trendline(type: TrendlineType, value: Int = 2) -> Series {
    chart_series_set_trendline(lxwChartSeries, UInt8(type.rawValue.rawValue), UInt8(value))
    return self
  }
  /// Set the trendline name for a chart data series.
  @discardableResult public func trendline(name: String) -> Series {
    let _ = name.withCString { chart_series_set_trendline_name(lxwChartSeries, $0) }
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
    chart_series_set_trendline_line(lxwChartSeries, &line)
    return self
  }
  /// Display the equation of a trendline for a chart data series.
  @discardableResult public func trendline_equation() -> Series {
    chart_series_set_trendline_equation(lxwChartSeries)
    return self
  }
  /// Display the R squared value of a trendline for a chart data series.
  @discardableResult public func trendline_r_squared() -> Series {
    chart_series_set_trendline_r_squared(lxwChartSeries)
    return self
  }
}
