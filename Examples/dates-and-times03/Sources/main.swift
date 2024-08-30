import xlsxwriter

let workbook = Workbook(name: "date_and_times03.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(numFormat: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(0, 0), width: 20, format: nil)
worksheet.write(unixtime: 0, .init(row: 0, col: 0), format: format)
worksheet.write(unixtime: 1577836_800, .init(row: 1, col: 0), format: format)
worksheet.write(unixtime: -2208988_800, .init(row: 2, col: 0), format: format)

workbook.close()
