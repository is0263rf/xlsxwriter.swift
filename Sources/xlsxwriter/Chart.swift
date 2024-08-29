//
//  Chart.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent an Excel chart.
public struct Chart {
  let lxw_chart: UnsafeMutablePointer<lxw_chart>
  init(_ lxw_chart: UnsafeMutablePointer<lxw_chart>) { self.lxw_chart = lxw_chart }
  /// Add a data series to a chart.
  @discardableResult public func addSeries(values: String? = nil, name: String? = nil) -> Series {
    let series: Series
    if let values = values {
      series = values.withCString { Series(chart_add_series(lxw_chart, nil, $0)) }
    } else {
      series = Series(chart_add_series(lxw_chart, nil, nil))
    }
    if let name = name { series.set(name: name) }
    return series
  }

  /// Remove/hide one or more series in a chart legend (the series will still display on the chart).
  @discardableResult public func remove(legends: Int...) -> Chart {
    var array = legends.map { Int16($0) }
    array.append(-1)
    chart_legend_delete_series(lxw_chart, &array)
    return self
  }

  /// Set the minimum and maximum value for the axis range.
  @discardableResult public func set(x_axis: ClosedRange<Double>) -> Chart {
    chart_axis_set_min(lxw_chart.pointee.x_axis, x_axis.lowerBound)
    chart_axis_set_max(lxw_chart.pointee.x_axis, x_axis.upperBound)
    return self
  }

  /// Set the minimum and maximum value for the axis range.
  @discardableResult public func set(y_axis: ClosedRange<Double>) -> Chart {
    chart_axis_set_min(lxw_chart.pointee.y_axis, y_axis.lowerBound)
    chart_axis_set_max(lxw_chart.pointee.y_axis, y_axis.upperBound)
    return self
  }

  /// Set the name caption of the x axis.
  @discardableResult public func set(x_axis name: String) -> Chart {
    name.withCString { chart_axis_set_name(lxw_chart.pointee.x_axis, $0) }
    return self
  }
  /// Set the name caption of the y axis.
  @discardableResult public func set(y_axis name: String) -> Chart {
    name.withCString { chart_axis_set_name(lxw_chart.pointee.y_axis, $0) }
    return self
  }
  /// Set the title of the chart.
  @discardableResult public func title(name: String) -> Chart {
    name.withCString { chart_title_set_name(lxw_chart, $0) }

    return self
  }
  /// Turn off/hide axis.
  @discardableResult public func axis_off(_ axis: Axes) -> Chart {
    switch axis {
    case .x: chart_axis_off(lxw_chart.pointee.x_axis)
    case .y: chart_axis_off(lxw_chart.pointee.y_axis)
    }
    return self
  }
  /// Turn on/off the major gridlines for an axis.
  @discardableResult public func major_gridlines(_ axis: Axes, visible: Bool = true) -> Chart {
    let visible = UInt8(visible ? 1 : 0)
    switch axis {
    case .x: chart_axis_major_gridlines_set_visible(lxw_chart.pointee.x_axis, visible)
    case .y: chart_axis_major_gridlines_set_visible(lxw_chart.pointee.y_axis, visible)
    }
    return self
  }
  /// Turn on/off the minor gridlines for an axis.
  @discardableResult public func minor_gridlines(_ axis: Axes, visible: Bool = true) -> Chart {
    let visible = UInt8(visible ? 1 : 0)
    switch axis {
    case .x: chart_axis_minor_gridlines_set_visible(lxw_chart.pointee.x_axis, visible)
    case .y: chart_axis_minor_gridlines_set_visible(lxw_chart.pointee.y_axis, visible)
    }
    return self
  }
  /// Set the increment of the major units in the axis
  @discardableResult public func major_unit(_ axis: Axes, _ unit: Double) -> Chart {
    switch axis {
    case .x: chart_axis_set_major_unit(lxw_chart.pointee.x_axis, unit)
    case .y: chart_axis_set_major_unit(lxw_chart.pointee.y_axis, unit)
    }
    return self
  }
  /// Set the increment of the minor units in the axis.
  @discardableResult public func minor_unit(_ axis: Axes, _ unit: Double) -> Chart {
    switch axis {
    case .x: chart_axis_set_minor_unit(lxw_chart.pointee.x_axis, unit)
    case .y: chart_axis_set_minor_unit(lxw_chart.pointee.y_axis, unit)
    }
    return self
  }
  /// Turn off an automatic chart title.
  @discardableResult public func title_off() -> Chart {
    chart_title_off(lxw_chart)
    return self
  }
  /// Set the position of the chart legend.
  @discardableResult public func legend(position: LegendPosition) -> Chart {
    chart_legend_set_position(lxw_chart, position.rawValue)
    chart_axis_off(lxw_chart.pointee.y_axis)
    return self
  }
  /// Set the Pie/Doughnut chart rotation.
  @discardableResult public func set(rotation: Int) -> Chart {
    chart_set_rotation(lxw_chart, UInt16(rotation))
    return self
  }
  /// Turn on a data table below the horizontal axis.
  @discardableResult public func table(style: Int) -> Chart {
    chart_set_table(lxw_chart)
    return self
  }
  /// Turn on a data table below the horizontal axis.
  @discardableResult public func set(style: Int) -> Chart {
    chart_set_style(lxw_chart, UInt8(style))
    return self
  }
}
