import xlsxwriter

let datetime: Datetime = .init(year: 2013, month: 1, day: 23, hour: 12, min: 30, sec: 5.123)
var row: UInt32 = 0
var col: UInt16 = 0

let dateFormats: [String] = [
  "dd/mm/yy",
  "mm/dd/yy",
  "dd m yy",
  "d mm yy",
  "d mmm yy",
  "d mmmm yy",
  "d mmmm yyy",
  "d mmmm yyyy",
  "dd/mm/yy hh:mm",
  "dd/mm/yy hh:mm:ss",
  "dd/mm/yy hh:mm:ss.000",
  "hh:mm",
  "hh:mm:ss",
  "hh:mm:ss.000",
]

let workbook = Workbook(name: "date_and_times04.xlsx")
let worksheet = workbook.addWorksheet()

let bold = workbook.addFormat()
bold.bold()

worksheet.write(string: "Formatted date", .init(row: row, col: col), format: bold)
worksheet.write(string: "Format", .init(row: row, col: col + 1), format: bold)

worksheet.column(.init(0, 1), width: 20, format: nil)

for dateFormat in dateFormats {
  row += 1
  let format = workbook.addFormat()
  format.set(numFormat: dateFormat)
    format.align(.left)

    worksheet.write(datetime: datetime, .init(row: row, col: col), format: format)
    worksheet.write(string: dateFormat, .init(row: row, col: col + 1), format: nil)
}

workbook.close()
