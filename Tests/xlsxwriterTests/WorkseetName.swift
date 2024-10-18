//
//  WorkseetName.swift
//  xlsxwriter.swift
//
//  Created by Yoshinori Takada on 2024/10/18.
//

import CoreXLSX
import Testing

@testable import xlsxwriter

@Suite struct WorkseetName {
    let fileName = "worksheet-name-test.xlsx"

    init() {
        let workbook = Workbook(name: fileName)
        let _ = workbook.addWorksheet(name: "worksheet1")
        let _ = workbook.addWorksheet(name: "worksheet2")

        workbook.close()
    }

    @Test func testWorkhseetName() throws {
        let file = try #require(XLSXFile(filepath: fileName))
        let workbook = try #require(file.parseWorkbooks().first)
        let worksheetNames = try file.parseWorksheetPathsAndNames(workbook: workbook).compactMap {
            $0.name
        }

        #expect(worksheetNames.contains("worksheet1"))
        #expect(worksheetNames.contains("worksheet2"))
    }
}
