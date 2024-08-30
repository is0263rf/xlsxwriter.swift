import xlsxwriter

let datetime: Datetime = .init(year: 2013, month: 2, day: 28, hour: 12)

let workbook = Workbook(name: "date_and_times02.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(numFormat: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(0, 0), width: 20, format: nil)
worksheet.write(datetime: datetime, .init(row: 0, col: 0), format: nil)
worksheet.write(datetime: datetime, .init(row: 1, col: 0), format: format)

workbook.close()
