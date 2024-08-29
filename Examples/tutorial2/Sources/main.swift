import xlsxwriter

struct Expense {
    var item: String
    var cost: Int
}

let expenses: [Expense] = [.init(item: "Rent", cost: 1000), .init(item: "Gas", cost: 100), .init(item: "Food", cost: 300), .init(item: "Gym", cost: 50)]

let workbook = Workbook(name: "tutorial02.xlsx")
let worksheet = workbook.addWorksheet()

let bold = workbook.addFormat()
bold.bold()

let money = workbook.addFormat()
money.set(num_format: "$#,##0")

var row = 0

worksheet.write(.string("Item"), .init(arrayLiteral: row, 0), format: bold)
worksheet.write(.string("Cost"), .init(arrayLiteral: row, 1), format: bold)

for (row, expense) in expenses.enumerated() {
    worksheet.write(.string(expense.item), .init(arrayLiteral: row + 1, 0), format: nil)
    worksheet.write(.number(Double(expense.cost)), .init(arrayLiteral: row + 1, 1), format: money)
}

worksheet.write(.string("Total"), .init(arrayLiteral: expenses.count + 1, 0), format: bold)
worksheet.write(.formula("=SUM(B2:B5)"), .init(arrayLiteral: expenses.count + 1, 1), format: money)

workbook.close()
