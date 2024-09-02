//
//  Datetime.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

/// Struct to represent a date and time in Excel.
public struct Datetime {
  var lxwDatetime: lxw_datetime

  /// Year : 1900 - 9999
  public var year: Int32 {
    get {
      return lxwDatetime.year
    }
    set(y) {
      lxwDatetime.year = y
    }
  }

  /// Month : 1 - 12
  public var month: Int32 {
    get {
      return lxwDatetime.month
    }
    set(m) {
      lxwDatetime.month = m
    }
  }

  /// Day : 1 - 31
  public var day: Int32 {
    get {
      return lxwDatetime.day
    }
    set(d) {
      lxwDatetime.day = d
    }
  }

  /// Hour : 0 - 23
  public var hour: Int32 {
    get {
      return lxwDatetime.hour
    }
    set(h) {
      lxwDatetime.hour = h
    }
  }

  /// Minute : 0 - 59
  public var min: Int32 {
    get {
      return lxwDatetime.min
    }
    set(m) {
      lxwDatetime.min = m
    }
  }

  /// Seconds : 0 - 59.999
  public var sec: Double {
    get {
      return lxwDatetime.sec
    }
    set(s) {
      lxwDatetime.sec = s
    }
  }

  public init(
    year: Int32 = 0, month: Int32 = 0, day: Int32 = 0, hour: Int32 = 0, min: Int32 = 0,
    sec: Double = 0
  ) {
    lxwDatetime = .init(year: year, month: month, day: day, hour: hour, min: min, sec: sec)
  }
}
