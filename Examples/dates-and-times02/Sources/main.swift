import xlsxwriter

let datetime: Datetime = .init(year: 2013, month: 2, day: 28, hour: 12)

let workbook = Workbook(name: "date_and_times02.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(num_format: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(arrayLiteral: 0, 0), width: 20, format: nil)
worksheet.write(.datetime(datetime), .init(arrayLiteral: 0, 0), format: nil)
worksheet.write(.datetime(datetime), .init(arrayLiteral: 1, 0), format: format)

workbook.close()
