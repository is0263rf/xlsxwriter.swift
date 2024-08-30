import xlsxwriter

let workbook = Workbook(name: "hyperlinks.xlsx")
let worksheet = workbook.addWorksheet()

let urlFormat = workbook.defaultUrlFormat()
let redFormat = workbook.addFormat()
redFormat.underline(.single)
redFormat.font(color: .red)

worksheet.column(.init(arrayLiteral: 0, 0), width: 30, format: nil)
worksheet.write(.url(.init(string: "http://libxlsxwriter.github.io")!), .init(0, 0), format: nil)
worksheet.write(.url(.init(string: "http://libxlsxwriter.github.io")!), .init(2, 0), format: nil)
worksheet.write(.string("Read the documentation."), .init(2, 0), format: urlFormat)
worksheet.write(.url(.init(string: "http://libxlsxwriter.github.io")!), .init(4, 0), format: redFormat)
worksheet.write(.url(.init(string: "mailto:jmcnamara@cpan.org")!), .init(6, 0), format: nil)
worksheet.write(.url(.init(string: "mailto:jmcnamara@cpan.org")!), .init(8, 0), format: nil)
worksheet.write(.string("Drop me a line."), .init(8, 0), format: urlFormat)

workbook.close()
