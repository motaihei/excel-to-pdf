param(
  [string]$folder  # フォルダパスを表す引数
)

# 相対パスを使用して動作させる
Set-Location $PSScriptRoot

# 渡されたフォルダの直下にpdfフォルダを作成（既に存在している場合はスキップ）
$pdf_folder = Join-Path $folder "pdf"
if (!(Test-Path $pdf_folder)) {
  New-Item -ItemType Directory $pdf_folder
}

# 渡されたフォルダの直下にある.xlsxファイルをPDFに変換する
Get-ChildItem -Path $folder -Include *.xlsx,*.xls -File | ForEach-Object {

  # ファイルパスを取得
  $abs_path = $_.FullName
  $name = $_.BaseName

  # 出力先のパスを定義
  $pdf_file = Join-Path $pdf_folder "$name.pdf"

  # ExcelファイルをPDFに変換
  Write-Host "$name.xlsxをPDF変換します"
  cscript //nologo "$PSScriptRoot\vbs\excel_to_pdf.vbs" "$abs_path" "$pdf_file"
}

Write-Host "処理終了"
Write-Host "出力先：$pdf_folder"
