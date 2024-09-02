import xlsxwriter

let workbook = Workbook(name: "lambda.xlsx")
defer { workbook.close() }

let worksheet = workbook.addWorksheet()

worksheet.write(
  dynamicFormula: "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)", .init("A1"))

workbook.define(name: "ToCelsius", "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))")

worksheet.write(dynamicFormula: "=ToCelsius(212)", .init("A2"))
