import xlsxwriter

let workbook = Workbook(name: "format_num_format.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.column(.init(0, 0), width: 30, format: nil)

let format01 = workbook.addFormat()
let format02 = workbook.addFormat()
let format03 = workbook.addFormat()
let format04 = workbook.addFormat()
let format05 = workbook.addFormat()
let format06 = workbook.addFormat()
let format07 = workbook.addFormat()
let format08 = workbook.addFormat()
let format09 = workbook.addFormat()
let format10 = workbook.addFormat()
let format11 = workbook.addFormat()

format01.set(numFormat: "0.000")
format02.set(numFormat: "#,##0")
format03.set(numFormat: "#,##0.00")
format04.set(numFormat: "0.00")
format05.set(numFormat: "mm/dd/yy")
format06.set(numFormat: "mmm d yyyy")
format07.set(numFormat: "d mmmm yyyy")
format08.set(numFormat: "dd/mm/yyyy hh:mm AM/PM")
format09.set(numFormat: "0 \"dollar and\" .00 \"cents\"")

worksheet.write(number: 3.1415926, .init(row: 0, col: 0), format: nil)
worksheet.write(number: 3.1415926, .init(row: 1, col: 0), format: format01)
worksheet.write(number: 1234.56, .init(row: 2, col: 0), format: format02)
worksheet.write(number: 1234.56, .init(row: 3, col: 0), format: format03)
worksheet.write(number: 49.99, .init(row: 4, col: 0), format: format04)
worksheet.write(number: 36892.521, .init(row: 5, col: 0), format: format05)
worksheet.write(number: 36892.521, .init(row: 6, col: 0), format: format06)
worksheet.write(number: 36892.521, .init(row: 7, col: 0), format: format07)
worksheet.write(number: 36892.521, .init(row: 8, col: 0), format: format08)
worksheet.write(number: 1.87, .init(row: 9, col: 0), format: format09)

format10.set(numFormat: "[Green]General;[Red]-General;General")
worksheet.write(number: 123, .init(row: 10, col: 0), format: format10)
worksheet.write(number: -45, .init(row: 11, col: 0), format: format10)
worksheet.write(number: 0, .init(row: 12, col: 0), format: format10)

format11.set(numFormat: "00000")
worksheet.write(number: 1209, .init(row: 13, col: 0), format: format11)

workbook.close()
