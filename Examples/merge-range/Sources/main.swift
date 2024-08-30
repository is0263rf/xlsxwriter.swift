import xlsxwriter

let workbook = Workbook(name: "merge_range.xlsx")
let worksheet = workbook.addWorksheet()
let mergeFormat = workbook.addFormat()

mergeFormat.align(.center)
mergeFormat.align(.verticalCenter)
mergeFormat.bold()
mergeFormat.background(color: .yellow)
mergeFormat.border(style: .thin)

worksheet.column(.init(1, 3), width: 12)
worksheet.row(3, height: 30)
worksheet.row(6, height: 30)
worksheet.row(7, height: 30)

worksheet.merge(range: .init(arrayLiteral: 3,1,3,3), string: "Merged Range", format: mergeFormat)
worksheet.merge(range: .init(arrayLiteral: 6,1,7,3), string: "Merged Range", format: mergeFormat)

workbook.close()
