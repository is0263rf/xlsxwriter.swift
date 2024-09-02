//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import Foundation
import libxlsxwriter

/// Struct to represent an Excel worksheet.
public struct Worksheet {
  private var lxwWorksheet: UnsafeMutablePointer<lxw_worksheet>

  var name: String { String(cString: lxwWorksheet.pointee.name) }

  init(_ lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) {
    self.lxwWorksheet = lxw_worksheet
  }

  /// Insert a chart object into a worksheet.
  public func insert(chart: Chart, _ cell: Cell) {
    let _ = worksheet_insert_chart(lxwWorksheet, cell.row, cell.col, chart.lxwChart)
  }

  /// Insert a chart object into a worksheet, with options.
  public func insert(chart: Chart, _ cell: Cell, scale: (x: Double, y: Double)) {
    var o = lxw_chart_options(
      x_offset: 0, y_offset: 0, x_scale: scale.x, y_scale: scale.y, object_position: 2,
      description: nil, decorative: 0)
    worksheet_insert_chart_opt(lxwWorksheet, cell.row, cell.col, chart.lxwChart, &o)
  }

  public func insert(image: String, _ cell: Cell) {
    let fileName = image.cString(using: .utf8)
    let _ = worksheet_insert_image(lxwWorksheet, cell.row, cell.col, fileName)
  }

  public func write(number: Double, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_number(lxwWorksheet, cell.row, cell.col, number, format?.lxwFormat)
  }

  public func write(string: String, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_string(lxwWorksheet, cell.row, cell.col, string, format?.lxwFormat)
  }

  public func write(url: URL, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_url(
      lxwWorksheet, cell.row, cell.col, url.absoluteString, format?.lxwFormat)
  }

  public func write(comment: String, _ cell: Cell) {
    let _ = worksheet_write_comment(lxwWorksheet, cell.row, cell.col, comment)
  }

  public func write(formula: String, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_formula(lxwWorksheet, cell.row, cell.col, formula, format?.lxwFormat)
  }

  public func write(arrayFormula: String, first: Cell, last: Cell, format: Format? = nil) {
    let _ = worksheet_write_array_formula(
      lxwWorksheet, first.row, first.col, last.row, last.col, arrayFormula, format?.lxwFormat)
  }

  public func write(arrayFormula: String, range: Range, format: Format? = nil) {
    let _ = worksheet_write_array_formula(
      lxwWorksheet, range.row, range.col, range.row2, range.col2, arrayFormula, format?.lxwFormat)
  }

  public func write(datetime: Datetime, _ cell: Cell, format: Format? = nil) {
    var d = lxw_datetime(
      year: datetime.year, month: datetime.month, day: datetime.day, hour: datetime.hour,
      min: datetime.min, sec: datetime.sec)
    let _ = worksheet_write_datetime(lxwWorksheet, cell.row, cell.col, &d, format?.lxwFormat)
  }

  public func write(unixtime: Int64, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_unixtime(lxwWorksheet, cell.row, cell.col, unixtime, format?.lxwFormat)
  }

  public func write(dynamicFormula: String, _ cell: Cell, format: Format? = nil) {
    let _ = worksheet_write_dynamic_formula(
      lxwWorksheet, cell.row, cell.col, dynamicFormula, format?.lxwFormat)
  }

  public func write(dynamicArrayFormula: String, _ range: Range, format: Format? = nil) {
    let _ = worksheet_write_dynamic_array_formula(
      lxwWorksheet, range.row, range.col, range.row2, range.col2, dynamicArrayFormula,
      format?.lxwFormat)
  }

  /// Set a worksheet tab as selected.
  public func select() {
    worksheet_select(lxwWorksheet)
  }

  /// Hide the current worksheet.
  public func hide() {
    worksheet_hide(lxwWorksheet)
  }

  /// Make a worksheet the active, i.e., visible worksheet.
  public func activate() {
    worksheet_activate(lxwWorksheet)
  }

  /// Hide zero values in worksheet cells.
  public func hideZero() {
    worksheet_hide_zero(lxwWorksheet)
  }

  /// Set the paper type for printing.
  public func paper(type: PaperType) {
    worksheet_set_paper(lxwWorksheet, type.rawValue)
  }

  /// Set the properties for one or more columns of cells.
  public func column(_ cols: Cols, width: Double, format: Format? = nil) {
    let first = cols.col
    let last = cols.col2
    let f = format?.lxwFormat
    let _ = worksheet_set_column(lxwWorksheet, first, last, width, f)
  }

  public func column(_ cols: Cols, pixels: UInt32, format: Format? = nil) {
    let _ = worksheet_set_column_pixels(lxwWorksheet, cols.col, cols.col, pixels, format?.lxwFormat)
  }

  /// Set the properties for a row of cells
  public func row(_ row: UInt32, height: Double, format: Format? = nil) {
    let f = format?.lxwFormat
    let _ = worksheet_set_row(lxwWorksheet, row, height, f)
  }

  /// Set the properties for one or more columns of cells.
  public func hideColumns(_ col: Int, width: Double = 8.43) {
    let first = UInt16(col)
    let cols: Cols = .init("A:XFD")
    let last = cols.col2
    var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
    _ = worksheet_set_column_opt(lxwWorksheet, first, last, width, nil, &o)
  }

  /// Set the color of the worksheet tab.
  public func tab(color: Color) {
    worksheet_set_tab_color(lxwWorksheet, color.hex)
  }

  /// Set the default row properties.
  public func defaultRow(height: Double, hideUnusedRows: Bool = true) {
    let hide: UInt8 = hideUnusedRows ? 1 : 0
    worksheet_set_default_row(lxwWorksheet, height, hide)
  }

  /// Set the print area for a worksheet.
  public func printArea(range: Range) {
    let _ = worksheet_print_area(lxwWorksheet, range.row, range.col, range.row2, range.col2)
  }

  /// Set the autofilter area in the worksheet.
  public func autofilter(range: Range) {
    let _ = worksheet_autofilter(lxwWorksheet, range.row, range.col, range.row2, range.col2)
  }

  /// Set the option to display or hide gridlines on the screen and the printed page.
  public func gridline(screen: Bool, print: Bool = false) {
    worksheet_gridlines(lxwWorksheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0)))
  }

  /// Set a table in the worksheet.
  public func table(
    range: Range, name: String? = nil, header: [(String, Format?)] = []
  ) {
    table(
      range: range, name: name, header: header.map { $0.0 }, format: header.map { $0.1 },
      totalRow: [])
  }

  /// Merge a range of cells in the worksheet.
  public func merge(range: Range, string: String, format: Format? = nil) {
    worksheet_merge_range(
      lxwWorksheet, range.row, range.col, range.row2, range.col2, string, format?.lxwFormat)
  }

  /// Set a table in the worksheet.
  public func table(
    range: Range, name: String? = nil, header: [String] = [], format: [Format?] = [],
    totalRow: [TotalFunction] = []
  ) {
    var options = lxw_table_options()
    if let name = name { options.name = makeCString(from: name) }
    options.style_type = UInt8(LXW_TABLE_STYLE_TYPE_MEDIUM.rawValue)
    options.style_type_number = 7
    options.total_row = totalRow.isEmpty ? UInt8(LXW_FALSE.rawValue) : UInt8(LXW_TRUE.rawValue)
    var table_columns = [lxw_table_column]()
    let buffer = UnsafeMutableBufferPointer<UnsafeMutablePointer<lxw_table_column>?>.allocate(
      capacity: header.count + 1)
    defer { buffer.deallocate() }
    if !header.isEmpty {
      table_columns = Array(repeating: lxw_table_column(), count: header.count)
      for i in header.indices {
        table_columns[i].header = makeCString(from: header[i])
        if format.endIndex > i {
          table_columns[i].header_format = format[i]?.lxwFormat
        }
        if totalRow.endIndex > i {
          table_columns[i].total_function = totalRow[i].rawValue
        }
        withUnsafeMutablePointer(to: &table_columns[i]) {
          buffer.baseAddress?.advanced(by: i).pointee = $0
        }
      }
      options.columns = buffer.baseAddress
    }
    _ = worksheet_add_table(
      lxwWorksheet, range.row, range.col, range.row2 + (totalRow.isEmpty ? 0 : 1), range.col2,
      &options)
    if let _ = name { options.name.deallocate() }
    table_columns.forEach { $0.header.deallocate() }
  }
}

private func makeCString(from str: String) -> UnsafePointer<CChar> {
  let count = str.utf8.count + 1
  let result = UnsafeMutablePointer<CChar>.allocate(capacity: count)
  str.withCString { result.initialize(from: $0, count: count) }
  return UnsafePointer(result)
}
