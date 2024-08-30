//
//  Datetime.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public struct Datetime {
  private var lxwDatetime: lxw_datetime

  public var year: Int32 {
    return lxwDatetime.year
  }

  public var month: Int32 {
    return lxwDatetime.month
  }

  public var day: Int32 {
    return lxwDatetime.day
  }

  public var hour: Int32 {
    return lxwDatetime.hour
  }

  public var min: Int32 {
    return lxwDatetime.min
  }

  public var sec: Double {
    return lxwDatetime.sec
  }

  public init(
    year: Int32 = 0, month: Int32 = 0, day: Int32 = 0, hour: Int32 = 0, min: Int32 = 0,
    sec: Double = 0
  ) {
    lxwDatetime = .init(year: year, month: month, day: day, hour: hour, min: min, sec: sec)
  }
}
