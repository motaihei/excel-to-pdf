@echo off
setlocal EnableDelayedExpansion

REM 渡された引数を取得
set folder=%1

REM PDFフォルダを作成（既に存在している場合はスキップ）
set "pdf_folder=%~dp1pdf"
if not exist "!pdf_folder!" (
  mkdir "!pdf_folder!"
)

REM 渡されたフォルダの直下の.xlsxファイルをPDFに変換する
for %%i in ("%folder%\*.xlsx", "%folder%\*.xls") do (

  REM 拡張子が.xlsxまたは.xlsのファイルのみを処理する
  set "ext=%%~xi"
  if /i "!ext!" neq ".xlsx" if /i "!ext!" neq ".xls" (
    continue
  )

  echo %%~nxi をPDF変換します
  set "abs_path=%%~fi"
  
  REM ファイル名と出力先のパスを定義
  for /f "delims=." %%a in ("%%~nxi") do (
    set "name=%%~na"
    set "ext=%%~xa"
  )
  set "pdf_file=!pdf_folder!\!name!.pdf"
  
  REM ExcelファイルをPDFに変換
  cscript //nologo "vbs\excel_to_pdf.vbs" "!abs_path!" "!pdf_file!"
)

echo 処理終了
echo 出力先：!pdf_folder! 

pause
