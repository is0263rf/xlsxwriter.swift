//
//  WriteString.swift
//  xlsxwriter.swift
//
//  Created by Yoshinori Takada on 2024/10/18.
//

import CoreXLSX
import Testing

@testable import xlsxwriter

@Suite struct WriteString {
    let fileName = "write-string-test.xlsx"

    init() {
        let workbook = Workbook(name: fileName)
        let worksheet = workbook.addWorksheet()
        worksheet.write(string: "hello", .init("A1"))
        workbook.close()
    }

    @Test func writeString() throws {
        let file = try #require(XLSXFile(filepath: fileName))
        let worksheetPath = try #require(file.parseWorksheetPaths().first)
        let worksheet = try file.parseWorksheet(at: worksheetPath)
        let sharedStrings = try file.parseSharedStrings()
        let sharedStrings2 = try #require(sharedStrings)
        let cellValue = try #require(
            worksheet.cells(atColumns: [ColumnReference("A")!]).compactMap {
                $0.stringValue(sharedStrings2)
            }.first)

        #expect(cellValue == "hello")
    }
}
