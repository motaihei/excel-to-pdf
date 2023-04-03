@echo off
setlocal EnableDelayedExpansion

REM 渡された引数を取得
set folder=%1

REM 渡されたフォルダが存在するかどうかを確認
if not exist "%folder%" (
  echo フォルダが存在しません。
  exit /b
)

REM 渡されたフォルダの直下のファイル一覧を表示
for %%i in ("%folder%\*.*") do (
  echo %%~nxi をPDF変換します
  set "abs_path=%%~fi"
  start "" /B call convert_excel_file_to_PDF.bat "!abs_path!"
)

pause
