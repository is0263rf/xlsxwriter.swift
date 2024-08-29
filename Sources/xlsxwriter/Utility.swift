//
//  Utility.swift
//
//
//  Created by Yoshinori Takada on 2024/08/29.
//

import libxlsxwriter

public class Utility {
  /// Retrieve the library version.
  public static func version() -> String {
    return String(cString: lxw_version())
  }
}
