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
    get {
      return lxwDatetime.year
    }
    set(y) {
      lxwDatetime.year = y
    }
  }

  public var month: Int32 {
    get {
      return lxwDatetime.month
    }
    set(m) {
      lxwDatetime.month = m
    }
  }

  public var day: Int32 {
    get {
      return lxwDatetime.day
    }
    set(d) {
      lxwDatetime.day = d
    }
  }

  public var hour: Int32 {
    get {
      return lxwDatetime.hour
    }
    set(h) {
      lxwDatetime.hour = h
    }
  }

  public var min: Int32 {
    get {
      return lxwDatetime.min
    }
    set(m) {
      lxwDatetime.min = m
    }
  }

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
