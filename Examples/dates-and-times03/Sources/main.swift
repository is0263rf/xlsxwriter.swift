import xlsxwriter

let workbook = Workbook(name: "date_and_times03.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(num_format: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(arrayLiteral: 0, 0), width: 20, format: nil)
worksheet.write(.unixtime(0), .init(0, 0), format: format)
worksheet.write(.unixtime(1577836800), .init(1, 0), format: format)
worksheet.write(.unixtime(-2208988800), .init(2, 0), format: format)

workbook.close()
