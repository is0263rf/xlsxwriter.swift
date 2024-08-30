//
//  Worksheet.swift
//  Created by Daniel Müllenborn on 31.12.20.
//

import libxlsxwriter

/// Struct to represent an Excel worksheet.
public struct Worksheet {
  private var lxwWorksheet: UnsafeMutablePointer<lxw_worksheet>

  var name: String { String(cString: lxwWorksheet.pointee.name) }

  init(_ lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) {
    self.lxwWorksheet = lxw_worksheet
  }

  /// Insert a chart object into a worksheet.
  public func insert(chart: Chart, _ pos: (row: UInt32, col: UInt16)) -> Worksheet {
    _ = worksheet_insert_chart(lxwWorksheet, pos.row, pos.col, chart.lxwChart)
    return self
  }

  /// Insert a chart object into a worksheet, with options.
  public func insert(chart: Chart, _ pos: (row: UInt32, col: UInt16), scale: (x: Double, y: Double))
    -> Worksheet
  {
    var o = lxw_chart_options(
      x_offset: 0, y_offset: 0, x_scale: scale.x, y_scale: scale.y, object_position: 2,
      description: nil, decorative: 0)
    worksheet_insert_chart_opt(lxwWorksheet, pos.row, pos.col, chart.lxwChart, &o)
    return self
  }

  public func insert(image: String, _ cell: Cell) {
    let fileName = image.cString(using: .utf8)
    worksheet_insert_image(lxwWorksheet, cell.row, cell.col, fileName)
  }

  /// Write a column of data starting from (row, col).
  @discardableResult public func write(column values: [Value], _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    var r = cell.row
    let c = cell.col
    for value in values {
      write(value, .init(row: r, col: c), format: format)
      r += 1
    }
    return self
  }

  /// Write a row of data starting from (row, col).
  @discardableResult public func write(row values: [Value], _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    let r = cell.row
    var c = cell.col
    for value in values {
      write(value, .init(row: r, col: c), format: format)
      c += 1
    }
    return self
  }

  /// Write a row of Double values starting from (row, col).
  @discardableResult public func write(
    _ numbers: [Double], row: Int, col: Int = 0, format: Format? = nil
  ) -> Worksheet {
    let f = format?.lxwFormat
    let r = UInt32(row)
    var c = UInt16(col)
    for number in numbers {
      worksheet_write_number(lxwWorksheet, r, c, number, f)
      c += 1
    }

    return self
  }

  /// Write a row of String values starting from (row, col).
  @discardableResult public func write(
    _ strings: [String], row: Int, col: Int = 0, format: Format? = nil
  ) -> Worksheet {
    let f = format?.lxwFormat
    let r = UInt32(row)
    var c = UInt16(col)
    for string in strings {
      let _ = string.withCString { s in worksheet_write_string(lxwWorksheet, r, c, s, f) }
      c += 1
    }

    return self
  }

  /// Write data to a worksheet cell by calling the appropriate
  /// worksheet_write_*() method based on the type of data being passed.
  @discardableResult public func write(_ value: Value, _ cell: Cell, format: Format? = nil)
    -> Worksheet
  {
    let r = cell.row
    let c = cell.col
    let f = format?.lxwFormat
    let error: lxw_error
    switch value {
    case .number(let number): error = worksheet_write_number(lxwWorksheet, r, c, number, f)
    case .string(let string):
      error = string.withCString { s in worksheet_write_string(lxwWorksheet, r, c, s, f) }
    case .url(let url):
      error = url.absoluteString.withCString { s in worksheet_write_url(lxwWorksheet, r, c, s, f) }
    case .blank: error = worksheet_write_blank(lxwWorksheet, r, c, f)
    case .comment(let comment):
      error = comment.withCString { s in worksheet_write_comment(lxwWorksheet, r, c, s) }
    case .boolean(let boolean):
      error = worksheet_write_boolean(lxwWorksheet, r, c, Int32(boolean ? 1 : 0), f)
    case .formula(let formula):
      error = formula.withCString { s in worksheet_write_formula(lxwWorksheet, r, c, s, f) }
    case .datetime(let datetime):
      var d: lxw_datetime = .init(
        year: datetime.year, month: datetime.month, day: datetime.day, hour: datetime.hour,
        min: datetime.min, sec: datetime.sec)
      error = worksheet_write_datetime(lxwWorksheet, r, c, &d, f)
    case .unixtime(let unixtime):
      error = worksheet_write_unixtime(lxwWorksheet, r, c, unixtime, f)
    }
    if error.rawValue != 0 { fatalError(String(cString: lxw_strerror(error))) }

    return self
  }

  @discardableResult public func writeArrayFormula(
    _ formula: String, _ first: Cell, _ last: Cell, format: Format? = nil
  ) -> Worksheet {
    worksheet_write_array_formula(
      lxwWorksheet, first.row, first.col, last.row, last.col, formula, format?.lxwFormat)

    return self
  }

  @discardableResult public func writeArrayFormula(
    _ value: String, range: String, format: Format? = nil
  ) -> Worksheet {
    let firstRow = lxw_name_to_row(range)
    let firstCol = lxw_name_to_col(range)
    let lastRow = lxw_name_to_row_2(range)
    let lastCol = lxw_name_to_col_2(range)
    let formula = value.cString(using: .utf8)

    worksheet_write_array_formula(
      lxwWorksheet, firstRow, firstCol, lastRow, lastCol, range, format?.lxwFormat)

    return self
  }

  /// Set a worksheet tab as selected.
  @discardableResult public func select() -> Worksheet {
    worksheet_select(lxwWorksheet)
    return self
  }

  /// Hide the current worksheet.
  @discardableResult public func hide() -> Worksheet {
    worksheet_hide(lxwWorksheet)
    return self
  }

  /// Make a worksheet the active, i.e., visible worksheet.
  @discardableResult public func activate() -> Worksheet {
    worksheet_activate(lxwWorksheet)
    return self
  }

  /// Hide zero values in worksheet cells.
  @discardableResult public func hide_zero() -> Worksheet {
    worksheet_hide_zero(lxwWorksheet)
    return self
  }

  /// Set the paper type for printing.
  @discardableResult public func paper(type: PaperType) -> Worksheet {
    worksheet_set_paper(lxwWorksheet, type.rawValue)
    return self
  }

  /// Set the properties for one or more columns of cells.
  @discardableResult public func column(_ cols: Cols, width: Double, format: Format? = nil)
    -> Worksheet
  {
    let first = cols.col
    let last = cols.col2
    let f = format?.lxwFormat
    _ = worksheet_set_column(lxwWorksheet, first, last, width, f)
    return self
  }

  /// Set the properties for a row of cells
  @discardableResult public func row(_ row: UInt32, height: Double, format: Format? = nil)
    -> Worksheet
  {
    let f = format?.lxwFormat
    _ = worksheet_set_row(lxwWorksheet, row, height, f)
    return self
  }

  /// Set the properties for one or more columns of cells.
  @discardableResult public func hide_columns(_ col: Int, width: Double = 8.43) -> Worksheet {
    let first = UInt16(col)
    let cols: Cols = "A:XFD"
    let last = cols.col2
    var o = lxw_row_col_options(hidden: 1, level: 0, collapsed: 0)
    _ = worksheet_set_column_opt(lxwWorksheet, first, last, width, nil, &o)
    return self
  }

  /// Set the color of the worksheet tab.
  @discardableResult public func tab(color: Color) -> Worksheet {
    worksheet_set_tab_color(lxwWorksheet, color.hex)
    return self
  }

  /// Set the default row properties.
  @discardableResult public func set_default(row_height: Double, hide_unused_rows: Bool = true)
    -> Worksheet
  {
    let hide: UInt8 = hide_unused_rows ? 1 : 0
    worksheet_set_default_row(lxwWorksheet, row_height, hide)
    return self
  }

  /// Set the print area for a worksheet.
  @discardableResult public func print_area(range: Range) -> Worksheet {
    let _ = worksheet_print_area(lxwWorksheet, range.row, range.col, range.row2, range.col2)
    return self
  }

  /// Set the autofilter area in the worksheet.
  @discardableResult public func autofilter(range: Range) -> Worksheet {
    let _ = worksheet_autofilter(lxwWorksheet, range.row, range.col, range.row2, range.col2)
    return self
  }

  /// Set the option to display or hide gridlines on the screen and the printed page.
  @discardableResult public func gridline(screen: Bool, print: Bool = false) -> Worksheet {
    worksheet_gridlines(lxwWorksheet, UInt8((print ? 2 : 0) + (screen ? 1 : 0)))
    return self
  }

  /// Set a table in the worksheet.
  @discardableResult public func table(
    range: Range, name: String? = nil, header: [(String, Format?)] = []
  ) -> Worksheet {
    table(
      range: range, name: name, header: header.map { $0.0 }, format: header.map { $0.1 },
      totalRow: [])
  }

  /// Merge a range of cells in the worksheet.
  @discardableResult public func merge(range: Range, string: String, format: Format? = nil)
    -> Worksheet
  {
    worksheet_merge_range(
      lxwWorksheet, range.row, range.col, range.row2, range.col2, string, format?.lxwFormat)
    return self
  }

  /// Set a table in the worksheet.
  @discardableResult public func table(
    range: Range, name: String? = nil, header: [String] = [], format: [Format?] = [],
    totalRow: [TotalFunction] = []
  ) -> Worksheet {
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
    return self
  }
}

private func makeCString(from str: String) -> UnsafePointer<CChar> {
  let count = str.utf8.count + 1
  let result = UnsafeMutablePointer<CChar>.allocate(capacity: count)
  str.withCString { result.initialize(from: $0, count: count) }
  return UnsafePointer(result)
}
