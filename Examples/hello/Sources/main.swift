import xlsxwriter

let workbook = Workbook(name: "hello_world.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.write(.string("Hello"), .init(arrayLiteral: 0, 0), format: nil)
worksheet.write(.number(123), .init(arrayLiteral: 1, 0), format: nil)

workbook.close()
