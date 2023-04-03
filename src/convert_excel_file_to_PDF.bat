@echo off

setlocal enabledelayedexpansion
set "file=%~f1"
for /f "delims=." %%a in ("%file%") do (
  set "name=%%~na"
  set "ext=%%~xa"
)
set "pdf_file=%~dp1!name!.pdf"

cscript //nologo "vbs\excel_to_pdf.vbs" "%file%" "%pdf_file%"
