import xlsxwriter

let workbook = Workbook(name: "utf8.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.write(.string("Это фраза на русском!"), .init(2, 1), format: nil)

workbook.close()
