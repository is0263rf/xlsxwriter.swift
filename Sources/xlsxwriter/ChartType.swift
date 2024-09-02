//
//  ChartType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Available chart types.
public enum ChartType {
  /// None.
  case none

  /// Area chart.
  case area

  /// Area chart - stacked.
  case areaStacked

  /// Area chart - percentage stacked.
  case areaStackedPercent

  /// Bar chart.
  case bar

  /// Bar chart - stacked
  case barStacked

  /// Bar chart - percentage stacked.
  case barStackedPercent

  /// Column chart.
  case column

  /// Column chart - percentage stacked.
  case columnStacked

  /// Column chart - percentage stacked.
  case columnStackedPercent

  /// Doughnut chart.
  case doughnut

  /// Line chart.
  case line

  /// Line chart - stacked.
  case lineStacked

  /// Line chart - percentage stacked.
  case lineStackedPercent

  /// Pie chart.
  case pie

  /// Scatter chart.
  case scatter

  /// Scatter chart - straight.
  case scatterStraight

  /// Scatter chart - straight with markers.
  case scatterStraightWithMarkers

  /// Scatter chart - smooth.
  case scatterSmooth

  /// Scatter chart - smooth with markers.
  case scatterSmoothWithMarkers

  /// Radar chart.
  case radar

  /// Radar chart - with markers.
  case radarWithMarkers

  /// Radar chart - filled.
  case radarFilled

  public var rawValue: lxw_chart_type {
    switch self {
    case .none:
      return LXW_CHART_NONE
    case .area:
      return LXW_CHART_AREA
    case .areaStacked:
      return LXW_CHART_AREA_STACKED
    case .areaStackedPercent:
      return LXW_CHART_AREA_STACKED_PERCENT
    case .bar:
      return LXW_CHART_BAR
    case .barStacked:
      return LXW_CHART_BAR_STACKED
    case .barStackedPercent:
      return LXW_CHART_BAR_STACKED_PERCENT
    case .column:
      return LXW_CHART_COLUMN
    case .columnStacked:
      return LXW_CHART_COLUMN_STACKED
    case .columnStackedPercent:
      return LXW_CHART_COLUMN_STACKED_PERCENT
    case .doughnut:
      return LXW_CHART_DOUGHNUT
    case .line:
      return LXW_CHART_LINE
    case .lineStacked:
      return LXW_CHART_LINE_STACKED
    case .lineStackedPercent:
      return LXW_CHART_LINE_STACKED_PERCENT
    case .pie:
      return LXW_CHART_PIE
    case .scatter:
      return LXW_CHART_SCATTER
    case .scatterStraight:
      return LXW_CHART_SCATTER_STRAIGHT
    case .scatterStraightWithMarkers:
      return LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
    case .scatterSmooth:
      return LXW_CHART_SCATTER_SMOOTH
    case .scatterSmoothWithMarkers:
      return LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
    case .radar:
      return LXW_CHART_RADAR
    case .radarWithMarkers:
      return LXW_CHART_RADAR_WITH_MARKERS
    case .radarFilled:
      return LXW_CHART_RADAR_FILLED
    }
  }
}
