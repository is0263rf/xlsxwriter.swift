import xlsxwriter

let workbook = Workbook(name: "format_font.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.column(.init(arrayLiteral: 0, 0), width: 20, format: nil)

let format1 = workbook.addFormat()
let format2 = workbook.addFormat()
let format3 = workbook.addFormat()

format1.bold()
format2.italic()
format3.bold()
format3.italic()

worksheet.write(.string("This is bold"), .init(arrayLiteral: 0, 0), format: format1)
worksheet.write(.string("This is italic"), .init(arrayLiteral: 1, 0), format: format2)
worksheet.write(.string("Bold and italic"), .init(arrayLiteral: 2, 0), format: format3)

workbook.close()
