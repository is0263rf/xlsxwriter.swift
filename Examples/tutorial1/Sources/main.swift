import xlsxwriter

struct Expense {
    var item: String
    var cost: Int
}

let expenses: [Expense] = [.init(item: "Rent", cost: 1000), .init(item: "Gas", cost: 100), .init(item: "Food", cost: 300), .init(item: "Gym", cost: 50)]

let workbook = Workbook(name: "tutorial01.xlsx")
let worksheet = workbook.addWorksheet()

for (row, expense) in expenses.enumerated() {
    worksheet.write(.string(expense.item), .init(arrayLiteral: row, 0), format: nil)
    worksheet.write(.number(Double(expense.cost)), .init(arrayLiteral: row, 1), format: nil)
}

worksheet.write(.string("Total"), .init(arrayLiteral: expenses.count, 0), format: nil)
worksheet.write(.formula("=SUM(B1:B4)"), .init(arrayLiteral: expenses.count, 1), format: nil)

workbook.close()
