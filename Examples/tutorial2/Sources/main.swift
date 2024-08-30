import xlsxwriter

struct Expense {
  var item: String
  var cost: Int
}

let expenses: [Expense] = [
  .init(item: "Rent", cost: 1000), .init(item: "Gas", cost: 100), .init(item: "Food", cost: 300),
  .init(item: "Gym", cost: 50),
]

let workbook = Workbook(name: "tutorial02.xlsx")
let worksheet = workbook.addWorksheet()

let bold = workbook.addFormat()
bold.bold()

let money = workbook.addFormat()
money.set(numFormat: "$#,##0")

var row = 0

worksheet.write(string: "Item", .init(row: UInt32(row), col: 0), format: bold)
worksheet.write(string: "Cost", .init(row: UInt32(row), col: 1), format: bold)

for (index, expense) in expenses.enumerated() {
    worksheet.write(string: expense.item, .init(row: UInt32(index + 1), col: 0), format: nil)
    worksheet.write(number: Double(expense.cost), .init(row: UInt32(row + 1), col: 1), format: money)
}

worksheet.write(string: "Total", .init(row: UInt32(expenses.count + 1), col: 0), format: bold)
worksheet.write(formula: "=SUM(B2:B5)", .init(row: UInt32(expenses.count + 1), col: 1), format: money)

workbook.close()
