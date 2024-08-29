import xlsxwriter

let workbook = Workbook(name: "format_num_format.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.column(.init(arrayLiteral: 0, 0), width: 30, format: nil)

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

format01.set(num_format: "0.000")
format02.set(num_format: "#,##0")
format03.set(num_format: "#,##0.00")
format04.set(num_format: "0.00")
format05.set(num_format: "mm/dd/yy")
format06.set(num_format: "mmm d yyyy")
format07.set(num_format: "d mmmm yyyy")
format08.set(num_format: "dd/mm/yyyy hh:mm AM/PM")
format09.set(num_format: "0 \"dollar and\" .00 \"cents\"")

worksheet.write(.number(3.1415926), .init(arrayLiteral: 0, 0), format: nil)
worksheet.write(.number(3.1415926), .init(arrayLiteral: 1, 0), format: format01)
worksheet.write(.number(1234.56), .init(arrayLiteral: 2, 0), format: format02)
worksheet.write(.number(1234.56), .init(arrayLiteral: 3, 0), format: format03)
worksheet.write(.number(49.99), .init(arrayLiteral: 4, 0), format: format04)
worksheet.write(.number(36892.521), .init(arrayLiteral: 5, 0), format: format05)
worksheet.write(.number(36892.521), .init(arrayLiteral: 6, 0), format: format06)
worksheet.write(.number(36892.521), .init(arrayLiteral: 7, 0), format: format07)
worksheet.write(.number(36892.521), .init(arrayLiteral: 8, 0), format: format08)
worksheet.write(.number(1.87), .init(arrayLiteral: 9, 0), format: format09)

format10.set(num_format: "[Green]General;[Red]-General;General")
worksheet.write(.number(123), .init(arrayLiteral: 10, 0), format: format10)
worksheet.write(.number(-45), .init(arrayLiteral: 11, 0), format: format10)
worksheet.write(.number(0), .init(arrayLiteral: 12, 0), format: format10)

format11.set(num_format: "00000")
worksheet.write(.number(1209), .init(arrayLiteral: 13, 0), format: format11)

workbook.close()
