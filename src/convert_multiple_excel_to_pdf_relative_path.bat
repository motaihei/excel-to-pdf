@echo off
REM 相対パスを使用して動作させる

setlocal EnableDelayedExpansion

REM 渡された引数を取得
set folder=%1

REM PDFフォルダを作成（既に存在している場合はスキップ）
set "pdf_folder=%~dp1pdf"
if not exist "%~dp1pdf" (
  mkdir "%~dp1pdf"
)

REM 渡されたフォルダの直下の.xlsxファイルをPDFに変換する
for %%i in ("%folder%\*.xlsx", "%folder%\*.xls") do (

  REM 拡張子が.xlsxまたは.xlsのファイルのみを処理する
  set "ext=%%~xi"
  if /i "!ext!" neq ".xlsx" if /i "!ext!" neq ".xls" (
    continue
  )

  REM ファイルパスを取得
  set "abs_path=%%~dpnxi"
  set "name=%%~ni"
  
  REM 出力先のパスを定義
  set "pdf_file=%pdf_folder%\!name!.pdf"
  
  REM ExcelファイルをPDFに変換
  echo !name!.xlsxをPDF変換します
  cscript //nologo "%~dp0vbs\excel_to_pdf.vbs" "!abs_path!" "!pdf_file!"
)

echo 処理終了
echo 出力先：%pdf_folder%

pause
