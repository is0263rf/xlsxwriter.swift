//
//  WriteNumber.swift
//  xlsxwriter.swift
//
//  Created by Yoshinori Takada on 2024/10/18.
//

import CoreXLSX
import Testing

@testable import xlsxwriter

@Suite struct WriteNumber {
    let fileName = "write-number-test.xlsx"

    init() {
        let workbook = Workbook(name: fileName)
        let worksheet = workbook.addWorksheet()
        worksheet.write(number: 42, .init("A1"))
        workbook.close()
    }

    @Test func writeNumber() throws {
        let file = try #require(XLSXFile(filepath: fileName))
        let worksheetPath = try #require(file.parseWorksheetPaths().first)
        let worksheet = try file.parseWorksheet(at: worksheetPath)
        let cell = try #require(worksheet.data?.rows.first?.cells.first)

        #expect(cell.value == "42")
    }
}
