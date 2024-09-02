//
//  ChartType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Available chart types.
public enum ChartType: UInt8 {
  /// None.
  case none

  /// Area chart.
  case area

  /// Area chart - stacked.
  case areaStacked

  /// Area chart - percentage stacked.
  case areaPercentageStacked

  /// Bar chart.
  case bar

  /// Bar chart - stacked
  case barStacked

  /// Bar chart - percentage stacked.
  case barPercentageStacked

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
  case linePercentageStacked

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
}
