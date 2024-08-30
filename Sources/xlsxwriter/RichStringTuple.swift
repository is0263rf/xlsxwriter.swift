//
//  RichStringTuple.swift
//
//
//  Created by Yoshinori Takada on 2024/08/30.
//

public struct RichStringTuple {
  let format: Format
  let string: String

  init(format: Format, string: String) {
    self.format = format
    self.string = string
  }
}
