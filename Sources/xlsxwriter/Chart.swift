//
//  Chart.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent an Excel chart.
public struct Chart {
  let lxwChart: UnsafeMutablePointer<lxw_chart>
  init(_ lxw_chart: UnsafeMutablePointer<lxw_chart>) { self.lxwChart = lxw_chart }
  /// Add a data series to a chart.
  @discardableResult public func addSeries(values: String? = nil, name: String? = nil) -> Series {
    let series: Series
    if let values = values {
      series = values.withCString { Series(chart_add_series(lxwChart, nil, $0)) }
    } else {
      series = Series(chart_add_series(lxwChart, nil, nil))
    }
    if let name = name { series.set(name: name) }
    return series
  }

  /// Remove/hide one or more series in a chart legend (the series will still display on the chart).
  @discardableResult public func remove(legends: Int...) -> Chart {
    var array = legends.map { Int16($0) }
    array.append(-1)
    chart_legend_delete_series(lxwChart, &array)
    return self
  }

  /// Set the minimum and maximum value for the axis range.
  @discardableResult public func set(xAxis: ClosedRange<Double>) -> Chart {
    chart_axis_set_min(lxwChart.pointee.x_axis, xAxis.lowerBound)
    chart_axis_set_max(lxwChart.pointee.x_axis, xAxis.upperBound)
    return self
  }

  /// Set the minimum and maximum value for the axis range.
  @discardableResult public func set(yAxis: ClosedRange<Double>) -> Chart {
    chart_axis_set_min(lxwChart.pointee.y_axis, yAxis.lowerBound)
    chart_axis_set_max(lxwChart.pointee.y_axis, yAxis.upperBound)
    return self
  }

  /// Set the name caption of the x axis.
  @discardableResult public func set(xAxis name: String) -> Chart {
    name.withCString { chart_axis_set_name(lxwChart.pointee.x_axis, $0) }
    return self
  }
  /// Set the name caption of the y axis.
  @discardableResult public func set(yAxis name: String) -> Chart {
    name.withCString { chart_axis_set_name(lxwChart.pointee.y_axis, $0) }
    return self
  }
  /// Set the title of the chart.
  @discardableResult public func title(name: String) -> Chart {
    name.withCString { chart_title_set_name(lxwChart, $0) }

    return self
  }
  /// Turn off/hide axis.
  @discardableResult public func axisOff(_ axis: Axis) -> Chart {
    switch axis {
    case .x: chart_axis_off(lxwChart.pointee.x_axis)
    case .y: chart_axis_off(lxwChart.pointee.y_axis)
    }
    return self
  }
  /// Turn on/off the major gridlines for an axis.
  @discardableResult public func majorGridlines(_ axis: Axis, visible: Bool = true) -> Chart {
    let visible = UInt8(visible ? 1 : 0)
    switch axis {
    case .x: chart_axis_major_gridlines_set_visible(lxwChart.pointee.x_axis, visible)
    case .y: chart_axis_major_gridlines_set_visible(lxwChart.pointee.y_axis, visible)
    }
    return self
  }
  /// Turn on/off the minor gridlines for an axis.
  @discardableResult public func minorGridlines(_ axis: Axis, visible: Bool = true) -> Chart {
    let visible = UInt8(visible ? 1 : 0)
    switch axis {
    case .x: chart_axis_minor_gridlines_set_visible(lxwChart.pointee.x_axis, visible)
    case .y: chart_axis_minor_gridlines_set_visible(lxwChart.pointee.y_axis, visible)
    }
    return self
  }
  /// Set the increment of the major units in the axis
  @discardableResult public func majorUnit(_ axis: Axis, _ unit: Double) -> Chart {
    switch axis {
    case .x: chart_axis_set_major_unit(lxwChart.pointee.x_axis, unit)
    case .y: chart_axis_set_major_unit(lxwChart.pointee.y_axis, unit)
    }
    return self
  }
  /// Set the increment of the minor units in the axis.
  @discardableResult public func minorUnit(_ axis: Axis, _ unit: Double) -> Chart {
    switch axis {
    case .x: chart_axis_set_minor_unit(lxwChart.pointee.x_axis, unit)
    case .y: chart_axis_set_minor_unit(lxwChart.pointee.y_axis, unit)
    }
    return self
  }
  /// Turn off an automatic chart title.
  @discardableResult public func titleOff() -> Chart {
    chart_title_off(lxwChart)
    return self
  }
  /// Set the position of the chart legend.
  @discardableResult public func legend(position: LegendPosition) -> Chart {
    chart_legend_set_position(lxwChart, position.rawValue)
    chart_axis_off(lxwChart.pointee.y_axis)
    return self
  }
  /// Set the Pie/Doughnut chart rotation.
  @discardableResult public func set(rotation: Int) -> Chart {
    chart_set_rotation(lxwChart, UInt16(rotation))
    return self
  }
  /// Turn on a data table below the horizontal axis.
  @discardableResult public func table(style: Int) -> Chart {
    chart_set_table(lxwChart)
    return self
  }
  /// Turn on a data table below the horizontal axis.
  @discardableResult public func set(style: Int) -> Chart {
    chart_set_style(lxwChart, UInt8(style))
    return self
  }
}
