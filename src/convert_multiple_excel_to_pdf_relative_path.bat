@echo off
REM ���΃p�X���g�p���ē��삳����

setlocal EnableDelayedExpansion

REM �n���ꂽ�������擾
set folder=%1

REM PDF�t�H���_���쐬�i���ɑ��݂��Ă���ꍇ�̓X�L�b�v�j
set "pdf_folder=%~dp1pdf"
if not exist "%~dp1pdf" (
  mkdir "%~dp1pdf"
)

REM �n���ꂽ�t�H���_�̒�����.xlsx�t�@�C����PDF�ɕϊ�����
for %%i in ("%folder%\*.xlsx", "%folder%\*.xls") do (

  REM �g���q��.xlsx�܂���.xls�̃t�@�C���݂̂���������
  set "ext=%%~xi"
  if /i "!ext!" neq ".xlsx" if /i "!ext!" neq ".xls" (
    continue
  )

  REM �t�@�C���p�X���擾
  set "abs_path=%%~dpnxi"
  set "name=%%~ni"
  
  REM �o�͐�̃p�X���`
  set "pdf_file=%pdf_folder%\!name!.pdf"
  
  REM Excel�t�@�C����PDF�ɕϊ�
  echo !name!.xlsx��PDF�ϊ����܂�
  cscript //nologo "%~dp0vbs\excel_to_pdf.vbs" "!abs_path!" "!pdf_file!"
)

echo �����I��
echo �o�͐�F%pdf_folder%

pause
