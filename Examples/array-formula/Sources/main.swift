import xlsxwriter

let workbook = Workbook(name: "array_formula.xlxsx")
let worksheet = workbook.addWorksheet()

worksheet.write(number: 500, .init(row: 0, col: 1), format: nil)
worksheet.write(number: 10, .init(row: 1, col: 1), format: nil)
worksheet.write(number: 1, .init(row: 4, col: 1), format: nil)
worksheet.write(number: 2, .init(row: 5, col: 1), format: nil)
worksheet.write(number: 3, .init(row: 6, col: 1), format: nil)
worksheet.write(number: 300, .init(row: 0, col: 2), format: nil)
worksheet.write(number: 15, .init(row: 1, col: 2), format: nil)
worksheet.write(number: 20234, .init(row: 4, col: 2), format: nil)
worksheet.write(number: 21003, .init(row: 5, col: 2), format: nil)
worksheet.write(number: 10000, .init(row: 6, col: 2), format: nil)

worksheet.write(arrayFormula: "{=SUM(B1:C1*B2:C2)}", first: .init(row: 0, col: 0), last: .init(row: 0, col: 0), format: nil)
worksheet.write(arrayFormula: "{=SUM(B1:C1*B2:C2)}", range: "A2:A2", format: nil)
worksheet.write(arrayFormula: "{=TREND(C5:C7,B5:B7)}", first: .init(row: 4, col: 0), last: .init(row: 6, col: 0), format: nil)

workbook.close()
