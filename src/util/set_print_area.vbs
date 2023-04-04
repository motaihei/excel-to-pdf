'Excelオブジェクトを作成する
Set objExcel = CreateObject("Excel.Application")

'Excelファイルを開く
Set objWorkbook = objExcel.Workbooks.Open("ファイル名.xlsx")

'全シートに対して処理を行う
For Each objWorksheet In objWorkbook.Worksheets

    '印刷範囲を設定する
    objWorksheet.PageSetup.PrintArea = "$A$1:$G$50"

    '印刷方向を横に設定する
    objWorksheet.PageSetup.Orientation = objWorksheet.PageSetup.xlLandscape

Next

'Excelファイルを保存して閉じる
objWorkbook.Save
objWorkbook.Close
