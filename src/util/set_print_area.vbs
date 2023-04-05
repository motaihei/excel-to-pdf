' @echo off
' cscript.exe //nologo "C:\test.vbs"

'Excelオブジェクトを作成する
Set objExcel = CreateObject("Excel.Application")

'Excelファイルを開く
Set objWorkbook = objExcel.Workbooks.Open("ファイル名.xlsx")

' Excelファイルを開く
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\test.xlsx")

' ' シートをアクティブにする
' Set objWorksheet = objWorkbook.Worksheets("シート名")
' objWorksheet.Activate

' 設定処理
Sub SetPrintArea()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        With ActiveSheet.PageSetup
            .PrintArea = "$A$1:$AZ$46" ' 印刷範囲を設定する
            .Zoom = False ' 縮尺を自動にしない
            .FitToPagesWide = 1 ' 1ページ幅に調整する
            .FitToPagesTall = False ' 高さは自動にする
            .Zoom = 80 ' 表示拡大率を80%にする
            .PrintTitleRows = "" ' 印刷タイトル行はなし
            .PrintTitleColumns = "" ' 印刷タイトル列はなし
            .PrintGridlines = True ' グリッド線を印刷する
            .CenterHorizontally = True ' 水平方向に中央揃えにする
            .CenterVertically = False ' 垂直方向は揃えない
            .Orientation = xlPortrait ' 用紙の向きを縦にする
            .PaperSize = xlPaperA4 ' 用紙サイズをA4にする
            .FirstPageNumber = xlAutomatic ' 自動ページ番号
            .Order = xlDownThenOver ' ページ印刷順序
            .BlackAndWhite = False ' 白黒印刷しない
            .Draft = False ' 下書き品質で印刷しない
            .PrintComments = xlPrintNoComments ' コメントは印刷しない
            .PrintErrors = xlPrintErrorsDisplayed ' エラーは印刷しない
            .PrintQuality = 600 ' 印刷品質を600にする
            .RightFooter = "Page &P of &N" ' 右下にページ番号を表示する
            .LeftHeader = "" ' 左上に何も表示しない
            .RightHeader = "" ' 右上に何も表示しない
            .CenterHeader = "" ' 上部中央に何も表示しない
            .CenterFooter = "" ' 下部中央に何も表示しない
        End With
        Range("A1").Select ' 処理シートのセルの選択を"A,1"にする
        ActiveWindow.Zoom = 80 ' 表示拡大率を80%にする
        ActiveWindow.View = xlPageBreakPreview ' 表示モードを「改ページモード」にする
    Next ws
End Sub

' Excelファイルを保存して閉じる
objWorkbook.Close True
objExcel.Quit
