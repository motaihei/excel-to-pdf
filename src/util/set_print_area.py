import openpyxl

# Excelファイルを開く
workbook = openpyxl.load_workbook('ファイル名.xlsx')

# 全シートに対して処理を行う
for sheetname in workbook.sheetnames:
    worksheet = workbook[sheetname]

    # 印刷範囲を設定する
    worksheet.print_area = 'A1:G50'

    # 印刷方向を横に設定する
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE

# Excelファイルを保存する
workbook.save('ファイル名.xlsx')
