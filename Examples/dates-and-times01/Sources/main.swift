import xlsxwriter

let number = 41333.5

let workbook = Workbook(name: "date_and_times01.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(num_format: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(arrayLiteral: 0, 0), width: 20)
worksheet.write(.number(number), .init(arrayLiteral: 0, 0), format: nil)
worksheet.write(.number(number), .init(arrayLiteral: 1, 0), format: format)

workbook.close()
