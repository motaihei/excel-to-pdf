Option Explicit

Dim Excel
Dim ExcelDoc

'Opens the Excel file'
Set Excel = CreateObject("Excel.Application")
Set ExcelDoc = Excel.Workbooks.Open(WScript.Arguments.Item(0))

'Creates the pdf file'
ExcelDoc.ExportAsFixedFormat 0, WScript.Arguments.Item(1), 0, 1, 0, , , 0

'Closes the Excel file'
Excel.ActiveWorkbook.Close
Excel.Application.Quit
