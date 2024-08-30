import xlsxwriter

struct WorksheetData {
    let col1: String
    let col2: String
    let col3: String
    let col4: Int
    
    init(col1: String, col2: String, col3: String, col4: Int) {
        self.col1 = col1
        self.col2 = col2
        self.col3 = col3
        self.col4 = col4
    }
}

func writeWorksheetData(worksheet: Worksheet, header: Format) {
    let worksheetData:[WorksheetData] = [
        .init(col1: "East", col2: "Tom", col3: "Apple", col4: 6380),
        .init(col1: "West", col2: "Fred", col3: "Grape", col4: 5619),
        .init(col1: "North", col2: "Amy", col3: "Pear", col4: 4565),
        .init(col1: "South", col2: "Sal", col3: "Banana", col4: 5323),
        .init(col1: "East", col2: "Fritz", col3: "Apple", col4: 4394),
        .init(col1: "West", col2: "Sravan", col3: "Grape", col4: 7195),
        .init(col1: "North", col2: "Xi", col3: "Pear", col4: 5231),
        .init(col1: "South", col2: "Hector", col3: "Banana", col4: 2427),
        .init(col1: "East", col2: "Tom", col3: "Banana", col4: 4213),
        .init(col1: "West", col2: "Fred", col3: "Pear", col4: 3239),
        .init(col1: "North", col2: "Amy", col3: "Grape", col4: 6520),
        .init(col1: "South", col2: "Sal", col3: "Apple", col4: 1310),
        .init(col1: "East", col2: "Fritz", col3: "Banana", col4: 6274),
        .init(col1: "West", col2: "Sravan", col3: "Pear", col4: 4894),
        .init(col1: "North", col2: "Xi", col3: "Grape", col4: 7580),
        .init(col1: "South", col2: "Hector", col3: "Apple", col4: 9814)
    ]
    
    worksheet.write(string: "Region", .init("A1"), format: header)
    worksheet.write(string: "Sales Rep", .init("B1"), format: header)
    worksheet.write(string: "Product", .init("C1"), format: header)
    worksheet.write(string: "Units", .init("D1"), format: header)
    
    for (index, data) in worksheetData.enumerated() {
        worksheet.write(string: data.col1, .init(row: UInt32(index + 1), col: 0), format: nil)
        worksheet.write(string: data.col2, .init(row: UInt32(index + 1), col: 1), format: nil)
        worksheet.write(string: data.col3, .init(row: UInt32(index + 1), col: 2), format: nil)
        worksheet.write(number: Double(data.col4), .init(row: UInt32(index + 1), col: 3), format: nil)
    }
}


let workbook = Workbook(name: "dynamic_arrays.xlsx")
let worksheet1 = workbook.addWorksheet(name: "Filter")
let worksheet2 = workbook.addWorksheet(name: "Unique")
let worksheet3 = workbook.addWorksheet(name: "Sort")
let worksheet4 = workbook.addWorksheet(name: "Sortby")
let worksheet5 = workbook.addWorksheet(name: "Xlookup")
let worksheet6 = workbook.addWorksheet(name: "Xmatch")
let worksheet7 = workbook.addWorksheet(name: "Randarray")
let worksheet8 = workbook.addWorksheet(name: "Sequence")
let worksheet9 = workbook.addWorksheet(name: "Spill ranges")
let worksheet10 = workbook.addWorksheet(name: "Older functions")

let header1 = workbook.addFormat()
header1.background(color: .init(hex: 0x74AC4C))
header1.font(color: .init(hex: 0xFFFFFF))

let header2 = workbook.addFormat()
header2.background(color: .init(hex: 0x528FD3))
header2.font(color: .init(hex: 0xFFFFFF))

worksheet1.write(dynamicFormula: "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)", .init("F2"), format: nil)
worksheet1.write(string: "Product", .init("K1"), format: nil)
worksheet1.write(string: "Apple", .init("K2"), format: nil)
worksheet1.write(string: "Region", .init("F1"), format: nil)
worksheet1.write(string: "Sales Rep", .init("G1"), format: nil)
worksheet1.write(string: "Product", .init("H1"), format: nil)
worksheet1.write(string: "Units", .init("I1"), format: nil)
worksheet1.column(.init("E:E"), pixels: 20, format: nil)
worksheet1.column(.init("J:J"), pixels: 20, format: nil)

worksheet2.write(dynamicFormula: "=_xlfn.UNIQUE(B2:B17)", .init("F2"), format: nil)
worksheet2.write(dynamicFormula: "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))", .init("H2"), format: nil)
writeWorksheetData(worksheet: worksheet2, header: header1)
worksheet2.column(.init("E:E"), pixels: 20, format: nil)
worksheet2.column(.init("G:G"), pixels: 20, format: nil)

worksheet3.write(dynamicFormula: "=_xlfn._xlws.SORT(B2:B17)", .init("F2"), format: nil)
worksheet3.write(dynamicFormula: "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)", .init("H2"), format: nil)
worksheet3.write(string: "Sales Rep", .init("F1"), format: header2)
worksheet3.write(string: "Product", .init("H1"), format: header2)
worksheet3.write(string: "Units", .init("I1"), format: header2)
writeWorksheetData(worksheet: worksheet3, header: header1)
worksheet3.column(.init("E:E"), pixels: 20, format: nil)
worksheet3.column(.init("G:G"), pixels: 20, format: nil)

worksheet4.write(dynamicFormula:"=_xlfn.SORTBY(A2:B9,B2:B9)", .init("D2"), format: nil)
worksheet4.write(string: "Name", .init("A1"), format: header1)
worksheet4.write(string: "Age", .init("B1"), format: header1)
worksheet4.write(string: "Tom", .init("A2"))
worksheet4.write(string: "Fred", .init("A3"))
worksheet4.write(string: "Amy", .init("A4"))
worksheet4.write(string: "Sal", .init("A5"))
worksheet4.write(string: "Fritz", .init("A6"))
worksheet4.write(string: "Srivan", .init("A7"))
worksheet4.write(string: "Xi", .init("A8"))
worksheet4.write(string: "Hector", .init("A9"))
worksheet4.write(number: 52, .init("B2"))
worksheet4.write(number: 65, .init("B3"))
worksheet4.write(number: 22, .init("B4"))
worksheet4.write(number: 73, .init("B5"))
worksheet4.write(number: 19, .init("B6"))
worksheet4.write(number: 39, .init("B7"))
worksheet4.write(number: 19, .init("B8"))
worksheet4.write(number: 66, .init("B9"))
worksheet4.write(string: "Name", .init("D1"), format: header2)
worksheet4.write(string: "Age", .init("E1"), format: header2)
worksheet4.column(.init("C:C"), pixels: 20)

worksheet5.write(dynamicFormula: "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)", .init("F1"))
worksheet5.write(string: "Country", .init("A1"),format: header1)
worksheet5.write(string: "Abr", .init("B1"), format: header1)
worksheet5.write(string: "Prefix", .init("C1"), format: header1)
worksheet5.write(string: "China", .init("A2"))
worksheet5.write(string: "India", .init("A3"))
worksheet5.write(string: "United States", .init("A4"))
worksheet5.write(string: "Indonesia", .init("A5"))
worksheet5.write(string: "Brazil", .init("A6"))
worksheet5.write(string: "Pakistan", .init("A7"))
worksheet5.write(string: "Nigeria", .init("A8"))
worksheet5.write(string: "Bangladesh", .init("A9"))
worksheet5.write(string: "CN", .init("B2"))
worksheet5.write(string: "IN", .init("B3"))
worksheet5.write(string: "US", .init("B4"))
worksheet5.write(string: "ID", .init("B5"))
worksheet5.write(string: "BR", .init("B6"))
worksheet5.write(string: "PK", .init("B7"))
worksheet5.write(string: "NG", .init("B8"))
worksheet5.write(string: "BD", .init("B9"))
worksheet5.write(number: 86, .init("C2"))
worksheet5.write(number: 91, .init("C3"))
worksheet5.write(number: 1, .init("C4"))
worksheet5.write(number: 62, .init("C5"))
worksheet5.write(number: 55, .init("C6"))
worksheet5.write(number: 92, .init("C7"))
worksheet5.write(number: 234, .init("C8"))
worksheet5.write(number: 880, .init("C9"))
worksheet5.write(string: "Brazil", .init("E1"), format: header2)
worksheet5.column(.init("A:A"), pixels: 100)
worksheet5.column(.init("D:D"), pixels: 20)

worksheet6.write(dynamicFormula:"=_xlfn.XMATCH(C2,A2:A6)", .init("D2"))
worksheet6.write(string: "Product", .init("A1"), format: header1)
worksheet6.write(string: "Apple", .init("A2"))
worksheet6.write(string: "Grape", .init("A3"))
worksheet6.write(string: "Pear", .init("A4"))
worksheet6.write(string: "Banana", .init("A5"))
worksheet6.write(string: "Cherry", .init("A6"))
worksheet6.write(string: "Product", .init("C1"))
worksheet6.write(string: "Position", .init("C2"))
worksheet6.write(string: "Grape", .init("C3"))
worksheet6.column(.init("B:B"), pixels: 20)

worksheet7.write(dynamicFormula: "=_xlfn.RANDARRAY(5,3,1,100, TRUE)", .init("A1"))

worksheet8.write(dynamicFormula: "=_xlfn.SEQUENCE(4,5)", .init("A1"))

worksheet9.write(dynamicFormula: "=_xlfn.ANCHORARRAY(F2)", .init("H2"))
worksheet9.write(dynamicFormula: "=COUNTA(_xlfn.ANCHORARRAY(F2))", .init("J2"))
worksheet9.write(dynamicFormula: "=_xlfn.UNIQUE(B2:B17)", .init("F2"))
worksheet9.write(string: "Unique", .init("F1"), format: header2)
worksheet9.write(string: "Spill", .init("H1"), format: header2)
worksheet9.write(string: "Spill", .init("J1"), format: header2)
worksheet9.column(.init("E:E"), pixels: 20)
worksheet9.column(.init("G:G"), pixels: 20)
worksheet9.column(.init("I:I"), pixels: 20)

worksheet10.write(dynamicArrayFormula: "=LEN(A1:A3)", "B1:B3")
worksheet10.write(string: "Foo", .init("A1"))
worksheet10.write(string: "Food", .init("A2"))
worksheet10.write(string: "Frood", .init("A3"))

workbook.close()
