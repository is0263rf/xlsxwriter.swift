import xlsxwriter

struct Expense {
    var item: String
    var cost: Int
    var datetime: Datetime
}

let expenses: [Expense] = [.init(item: "Rent", cost: 1000, datetime: .init(year: 2013, month: 1, day: 13)), .init(item: "Gas", cost: 100, datetime: .init(year: 2013, month: 1, day: 14)), .init(item: "Food", cost: 300, datetime: .init(year: 2013, month: 1, day: 16)), .init(item: "Gym", cost: 50, datetime: .init(year: 2013, month: 1, day: 20))]

let workbook = Workbook(name: "tutorial03.xlsx")
let worksheet = workbook.addWorksheet()

let bold = workbook.addFormat()
bold.bold()

let money = workbook.addFormat()
money.set(num_format: "$#,##0")

let dateFormat = workbook.addFormat()
dateFormat.set(num_format: "mmmm d yyyy")

var row = 0

worksheet.column(.init(arrayLiteral: row, 0), width: 15, format: nil)

worksheet.write(.string("Item"), .init(arrayLiteral: row, 0), format: bold)
worksheet.write(.string("Cost"), .init(arrayLiteral: row, 1), format: bold)

for (index, expense) in expenses.enumerated() {
    row = index + 1
    worksheet.write(.string(expense.item), .init(arrayLiteral: row, 0), format: nil)
    worksheet.write(.datetime(expense.datetime), .init(arrayLiteral: row, 1), format: dateFormat)
    worksheet.write(.number(Double(expense.cost)), .init(arrayLiteral: row, 2), format: money)
}

worksheet.write(.string("Total"), .init(arrayLiteral: row + 1, 0), format: bold)
worksheet.write(.formula("=SUM(C2:C5)"), .init(arrayLiteral: row + 1, 2), format: money)

workbook.close()
