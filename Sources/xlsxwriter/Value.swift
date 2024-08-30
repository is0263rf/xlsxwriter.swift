//
//  Value.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import Foundation

public enum Value: ExpressibleByFloatLiteral, ExpressibleByStringLiteral {
  case url(URL)
  case blank
  case comment(String)
  case number(Double)
  case string(String)
  case boolean(Bool)
  case formula(String)
  case datetime(Datetime)
  case unixtime(Int64)
  public init(floatLiteral value: Double) { self = .number(value) }
  public init(stringLiteral value: String) { self = .string(value) }
}
