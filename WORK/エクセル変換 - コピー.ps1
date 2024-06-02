# 事前準備
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$fileType = "*xls"
$extension = ".xlsx"
$excel = New-Object -ComObject Excel.Application

# ファイルの一覧取得
$files = Get-ChildItem -Path "C:\Users\miyok\OneDrive\デスクトップ\テスト" -Include $fileType -Recurse 
 
# 1ファイルずつ処理
foreach($file in $files)
{
    # フォルダ取得
    $filePath = [System.IO.Path]::GetDirectoryName($file.FullName)
    # 拡張子を含まないファイル名取得
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
    # xlsxファイルを保存するパスを作成
    $excelPath  = (Join-Path $filePath $fileName) + $extension
    
    # Excelの処理開始
    $excel.Visible = $true
    $workbook = $excel.workbooks.open($file)
    $workbook.SaveAs($excelPath,51)
    $workbook.close($false)
 
    # Excelオブジェクト破棄
    $excel.Quit()
}
 
# 最終処理
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
 
Write-Host "処理完了"
pause