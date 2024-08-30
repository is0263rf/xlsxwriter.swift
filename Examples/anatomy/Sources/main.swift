import xlsxwriter

let workbook = Workbook(name: "anatomy.xlsx")

let worksheet1 = workbook.addWorksheet(name: "Demo")
let worksheet2 = workbook.addWorksheet()

let myFormat1 = workbook.addFormat()
let myFormat2 = workbook.addFormat()

myFormat1.bold()
myFormat2.set(numFormat: "$#,##0.00")

worksheet1.column(.init(0, 0), width: 20)

worksheet1.write(string: "Peach", .init(row: 0, col: 0), format: nil)
worksheet1.write(string: "Plum", .init(row: 1, col: 0), format: nil)
worksheet1.write(string: "Pear", .init(row: 2, col: 0), format: myFormat1)
worksheet1.write(string: "Persimmon", .init(row: 3, col: 0), format: myFormat1)
worksheet1.write(number: 123, .init(row: 5, col: 0), format: nil)
worksheet1.write(number: 4567.555, .init(row: 6, col: 0), format: myFormat2)
worksheet2.write(string: "Some text", .init(row: 0, col: 0), format: myFormat1)

workbook.close()
