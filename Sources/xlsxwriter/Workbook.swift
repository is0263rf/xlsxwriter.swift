//
//  Workbook.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel workbook.
public struct Workbook {
  var lxwWorkbook: UnsafeMutablePointer<lxw_workbook>

  /// Create a new workbook object.
  public init(name: String) {
    self.lxwWorkbook = name.withCString { workbook_new($0) }
  }

  /// Close the Workbook object and write the XLSX file.
  public func close() {
    let error = workbook_close(lxwWorkbook)
    if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }
  }

  /// Add a new worksheet to the Excel workbook.
  public func addWorksheet(name: String? = nil) -> Worksheet {
    let worksheet: UnsafeMutablePointer<lxw_worksheet>
    if let name = name {
      worksheet = name.withCString { workbook_add_worksheet(lxwWorkbook, $0) }
    } else {
      worksheet = workbook_add_worksheet(lxwWorkbook, nil)
    }
    return Worksheet(worksheet)
  }

  /// Add a new chartsheet to a workbook.
  public func addChartsheet(name: String? = nil) -> Chartsheet {
    let chartsheet: UnsafeMutablePointer<lxw_chartsheet>
    if let name = name {
      chartsheet = name.withCString { workbook_add_chartsheet(lxwWorkbook, $0) }
    } else {
      chartsheet = workbook_add_chartsheet(lxwWorkbook, nil)
    }
    return Chartsheet(chartsheet)
  }

  /// Add a new format to the Excel workbook.
  public func addFormat() -> Format {
    Format(workbook_add_format(lxwWorkbook))
  }

  /// Create a new chart to be added to a worksheet
  public func addChart(type: ChartType) -> Chart {
    Chart(workbook_add_chart(lxwWorkbook, type.rawValue))
  }

  /// Get a worksheet object from its name.
  public subscript(worksheet name: String) -> Worksheet? {
    guard let ws = name.withCString({ s in workbook_get_worksheet_by_name(lxwWorkbook, s) }) else {
      return nil
    }
    return Worksheet(ws)
  }

  /// Get a chartsheet object from its name.
  public subscript(chartsheet name: String) -> Chartsheet? {
    guard let cs = name.withCString({ s in workbook_get_chartsheet_by_name(lxwWorkbook, s) })
    else { return nil }
    return Chartsheet(cs)
  }

  /// Validate a worksheet or chartsheet name.
  public func validate(sheetName: String) {
    let _ = sheetName.withCString { workbook_validate_sheet_name(lxwWorkbook, $0) }
  }

  /// Add a recommendation to open the file in "read-only" mode.
  public func readOnlyRecommended() {
    let _ = workbook_read_only_recommended(lxwWorkbook)
  }

  /// Get the default URL format
  public func defaultUrlFormat() -> Format {
    let f = workbook_get_default_url_format(lxwWorkbook)!
    return .init(f)
  }
}
