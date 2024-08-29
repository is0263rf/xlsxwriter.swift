import xlsxwriter

let workbook = Workbook(name: "anatomy.xlsx")

let worksheet1 = workbook.addWorksheet(name: "Demo")
let worksheet2 = workbook.addWorksheet()

let myFormat1 = workbook.addFormat()
let myFormat2 = workbook.addFormat()

myFormat1.bold()
myFormat2.set(num_format: "$#,##0.00")

worksheet1.column(.init(arrayLiteral: 0, 0), width: 20)

worksheet1.write(.string("Peach"), .init(arrayLiteral: 0, 0), format: nil)
worksheet1.write(.string("Plum"), .init(arrayLiteral: 1, 0), format: nil)
worksheet1.write(.string("Pear"), .init(arrayLiteral: 2, 0), format: myFormat1)
worksheet1.write(.string("Persimmon"), .init(arrayLiteral: 3, 0), format: myFormat1)
worksheet1.write(.number(123), .init(arrayLiteral: 5, 0), format: nil)
worksheet1.write(.number(4567.555), .init(arrayLiteral: 6, 0), format: myFormat2)
worksheet2.write(.string("Some text"), .init(arrayLiteral: 0, 0), format: myFormat1)

workbook.close()
