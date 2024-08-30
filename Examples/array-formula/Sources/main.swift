import xlsxwriter

let workbook = Workbook(name: "array_formula.xlxsx")
let worksheet = workbook.addWorksheet()

worksheet.write(.number(500), .init(0, 1), format: nil)
worksheet.write(.number(10), .init(1, 1), format: nil)
worksheet.write(.number(1), .init(4, 1), format: nil)
worksheet.write(.number(2), .init(5, 1), format: nil)
worksheet.write(.number(3), .init(6, 1), format: nil)
worksheet.write(.number(300), .init(0, 2), format: nil)
worksheet.write(.number(15), .init(1, 2), format: nil)
worksheet.write(.number(20234), .init(4, 2), format: nil)
worksheet.write(.number(21003), .init(5, 2), format: nil)
worksheet.write(.number(10000), .init(6, 2), format: nil)

worksheet.writeArrayFormula("{=SUM(B1:C1*B2:C2)}", .init(0, 0), .init(0, 0), format: nil)
worksheet.writeArrayFormula("{=SUM(B1:C1*B2:C2)}",  range: "A2:A2", format: nil)
worksheet.writeArrayFormula("{=TREND(C5:C7,B5:B7)}", .init(4, 0), .init(6, 0), format: nil)

workbook.close()
