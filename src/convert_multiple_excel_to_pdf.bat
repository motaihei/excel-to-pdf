@echo off
setlocal EnableDelayedExpansion

REM �n���ꂽ�������擾
set folder=%1

REM PDF�t�H���_���쐬�i���ɑ��݂��Ă���ꍇ�̓X�L�b�v�j
set "pdf_folder=%~dp1pdf"
if not exist "!pdf_folder!" (
  mkdir "!pdf_folder!"
)

REM �n���ꂽ�t�H���_�̒�����.xlsx�t�@�C����PDF�ɕϊ�����
for %%i in ("%folder%\*.xlsx", "%folder%\*.xls") do (

  REM �g���q��.xlsx�܂���.xls�̃t�@�C���݂̂���������
  set "ext=%%~xi"
  if /i "!ext!" neq ".xlsx" if /i "!ext!" neq ".xls" (
    continue
  )

  echo %%~nxi ��PDF�ϊ����܂�
  set "abs_path=%%~fi"
  
  REM �t�@�C�����Əo�͐�̃p�X���`
  for /f "delims=." %%a in ("%%~nxi") do (
    set "name=%%~na"
    set "ext=%%~xa"
  )
  set "pdf_file=!pdf_folder!\!name!.pdf"
  
  REM Excel�t�@�C����PDF�ɕϊ�
  cscript //nologo "vbs\excel_to_pdf.vbs" "!abs_path!" "!pdf_file!"
)

echo �����I��
echo �o�͐�F!pdf_folder! 

pause
