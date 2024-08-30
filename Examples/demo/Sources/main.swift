import xlsxwriter

let workbook = Workbook(name: "demo.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.bold()

worksheet.column(.init(arrayLiteral: 0, 0), width: 20, format: nil)
worksheet.write(.string("Hello"), .init(arrayLiteral: 0, 0), format: nil)
worksheet.write(.string("World"), .init(arrayLiteral: 1, 0), format: format)
worksheet.write(.number(123), .init(arrayLiteral: 2, 0), format: nil)
worksheet.write(.number(123.456), .init(arrayLiteral: 3, 0), format: nil)
worksheet.insert(image: "logo.png", .init(1, 2))

workbook.close()
