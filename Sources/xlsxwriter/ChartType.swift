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
  case areaStacked
  case areaPercentageStacked
  case bar
  case barStacked
  case barPercentageStacked
  case column
  case columnStacked
  case columnPercentageStacked
  case doughnut
  case line
  case lineStacked
  case linePercentageStacked
  case pie
  case scatter
  case scatterStraight
  case scatterStraightWithMarkers
  case scatterSmooth
  case scatterSmoothWithMarkers
  case radar
  case radarWithMarkers
  case radarFilled
}
