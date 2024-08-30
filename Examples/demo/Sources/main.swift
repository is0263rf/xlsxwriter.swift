import xlsxwriter

let workbook = Workbook(name: "demo.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.bold()

worksheet.column(.init(0, 0), width: 20, format: nil)
worksheet.write(string: "Hello", .init(row: 0, col: 0), format: nil)
worksheet.write(string: "World", .init(row: 1, col: 0), format: format)
worksheet.write(number: 123, .init(row: 2, col: 0), format: nil)
worksheet.write(number: 123.456, .init(row: 3, col: 0), format: nil)
worksheet.insert(image: "logo.png", .init(row: 1, col: 2))

workbook.close()
