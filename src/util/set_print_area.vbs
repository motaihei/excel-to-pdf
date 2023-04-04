'Excelオブジェクトを作成する
Set objExcel = CreateObject("Excel.Application")

'Excelファイルを開く
Set objWorkbook = objExcel.Workbooks.Open("ファイル名.xlsx")

'全シートに対して処理を行う
For Each objWorksheet In objWorkbook.Worksheets

    '印刷範囲を設定する
    objWorksheet.PageSetup.PrintArea = "$A$1:$G$50"

    '印刷方向を横に設定する
    objWorksheet.PageSetup.Orientation = 2 ' xlLandscape の代わりに 2 を設定

    '印刷範囲を1ページに設定する
    objWorksheet.PageSetup.FitToPagesWide = 1
    objWorksheet.PageSetup.FitToPagesTall = 1

    'カーソルをA1に移動する
    objWorksheet.Range("A1").Select

Next

'拡大率を80%に設定する
objExcel.ActiveWindow.Zoom = 80

'改ページモードで表示する
objExcel.ActiveWindow.View = 3

'Excelファイルを保存して閉じる
objWorkbook.Save
objWorkbook.Close
