import xlsxwriter

let number = 41333.5

let workbook = Workbook(name: "date_and_times01.xlsx")
let worksheet = workbook.addWorksheet()

let format = workbook.addFormat()
format.set(numFormat: "mmm d yyyy hh:mm AM/PM")

worksheet.column(.init(0, 0), width: 20)
worksheet.write(number: number, .init(row: 0, col: 0), format: nil)
worksheet.write(number: number, .init(row: 1, col: 0), format: format)

workbook.close()
