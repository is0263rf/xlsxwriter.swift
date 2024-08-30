import xlsxwriter

let workbook = Workbook(name: "hello_world.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.write(string: "Hello", .init(row: 0, col: 0), format: nil)
worksheet.write(number: 123, .init(row: 1, col: 0), format: nil)

workbook.close()
