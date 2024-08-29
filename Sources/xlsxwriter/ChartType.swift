//
//  ChartType.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

/// Available chart types.
public enum ChartType: UInt8 {
  case none
  case area
  case area_stacked
  case area_percentage_stacked
  case bar
  case bar_stacked
  case bar_percentage_stacked
  case column
  case column_stacked
  case column_percentage_stacked
  case doughnut
  case line
  case line_stacked
  case line_percentage_stacked
  case pie
  case scatter
  case scatter_straight
  case scatter_straight_with_markers
  case scatter_smooth
  case scatter_smooth_with_markers
  case radar
  case radar_with_markers
  case radar_filled
}
