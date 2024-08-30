import xlsxwriter

let workbook = Workbook(name: "hyperlinks.xlsx")
let worksheet = workbook.addWorksheet()

let urlFormat = workbook.defaultUrlFormat()
let redFormat = workbook.addFormat()
redFormat.underline(.single)
redFormat.font(color: .red)

worksheet.column(.init(0, 0), width: 30, format: nil)
worksheet.write(url: .init(string: "http://libxlsxwriter.github.io")!, .init(row: 0, col: 0), format: nil)
worksheet.write(url: .init(string: "http://libxlsxwriter.github.io")!, .init(row: 2, col: 0), format: nil)
worksheet.write(string: "Read the documentation.", .init(row: 2, col: 0), format: urlFormat)
worksheet.write(url: .init(string: "http://libxlsxwriter.github.io")!, .init(row: 4, col: 0), format: redFormat)
worksheet.write(url: .init(string: "mailto:jmcnamara@cpan.org")!, .init(row: 6, col: 0), format: nil)
worksheet.write(url: .init(string: "mailto:jmcnamara@cpan.org")!, .init(row: 8, col: 0), format: nil)
worksheet.write(string: "Drop me a line.", .init(row: 8, col: 0), format: urlFormat)

workbook.close()
