import os
import subprocess

folder = "Excelファイル格納フォルダのパス"
vbs_path = '"excel_to_pdf.vbs"を格納しているパス'
pdf_folder = os.path.join(folder, 'pdf')
if not os.path.exists(pdf_folder):
    os.mkdir(pdf_folder)

for file in os.listdir(folder):
    if file.endswith('.xlsx') or file.endswith('.xls'):
        file_path = os.path.join(folder, file)
        pdf_path = os.path.join(pdf_folder, os.path.splitext(file)[0] + '.pdf')
        subprocess.call(['cscript', '//nologo', vbs_path, file_path, pdf_path])
