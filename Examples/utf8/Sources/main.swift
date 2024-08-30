import xlsxwriter

let workbook = Workbook(name: "utf8.xlsx")
let worksheet = workbook.addWorksheet()

worksheet.write(string: "Это фраза на русском!", .init(row: 2, col: 1), format: nil)

workbook.close()
