#Requires -Module ImportExcel

# This script generates an Excel file that mimics the vns-template.xlsx file.
# It manually builds the Excel package using the EPPlus API to ensure stability and control.

# --- 0. Preamble ---
try {
    # Manually load the EPPlus assembly to ensure types are recognized by the PowerShell session.
    # This is the key to solving the type not found errors.
    Add-Type -Path (Join-Path (Get-Module -ListAvailable ImportExcel | Where-Object { $_.Version.ToString() -eq '7.8.10' } | Select-Object -First 1).ModuleBase "EPPlus.dll")
}
catch {
    Write-Error "EPPlus.dll not found. Please ensure the ImportExcel module (version 7.8.10) is installed correctly."
    return
}

# --- 1. Data Preparation ---
Write-Host "Step 1: Preparing data..."
# Use a temporary CSV file to reliably extract all data from the template.
Import-Excel -Path .\vns-template.xlsx -WorksheetName '2023xxxxxxxx-vns_確認' | Export-Csv -Path ".\temp_data.csv" -NoTypeInformation -Encoding UTF8
# Import only the main 28 data rows, ignoring the footer notes in the template.
$fullData = Import-Csv -Path ".\temp_data.csv" | Select-Object -First 28

# Manually define the correct, multi-line headers as they appear in the template.
$correctHeaders = @(
    "[JVN]`n区分", "[JVN]`n番号", "[JVN]`nタイトル", "[JVN]`nCVSS v3`n深刻度", "[JVN]`nCVSS v2`n深刻度",
    "[JVN]`nCVSS v3 URL", "[JVN]`nCVSS v2 URL", "[JVN]`n更新日", "[JVN]`nCVE", "[Gauge]`nURL",
    "影響有無", "対象パッケージ／ソフトウェア", "現行バージョン／`n更新版バージョン", ""
)

# Define a unique output path to prevent file lock issues.
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$outputPath = ".\vns_report_$timestamp.xlsx"
$fileInfo = New-Object System.IO.FileInfo $outputPath
if ($fileInfo.Exists) { Remove-Item $outputPath }

# --- 2. Manual Excel Creation ---
Write-Host "Step 2: Creating a new Excel package..."
$pkg = New-Object OfficeOpenXml.ExcelPackage $fileInfo
$ws = $pkg.Workbook.Worksheets.Add('vns_確認')

# --- 3. Populate Data and Headers ---
Write-Host "Step 3: Populating headers and data..."
# Write headers cell by cell.
for ($i = 0; $i -lt $correctHeaders.Count; $i++) { $ws.Cells[1, ($i+1)].Value = $correctHeaders[$i] }

# Write data cell by cell for maximum stability.
$csvHeaders = $fullData[0].psobject.Properties.Name
for ($row = 0; $row -lt $fullData.Count; $row++) {
    for ($col = 0; $col -lt $csvHeaders.Count; $col++) {
        $ws.Cells[($row+2), ($col+1)].Value = $fullData[$row].($csvHeaders[$col])
    }
}

# --- 4. Apply All Formatting ---
Write-Host "Step 4: Applying formatting (columns, rows, fonts, borders)..."
$dataRowCount = $fullData.Count
$totalRange = "A1:N$($dataRowCount + 1)"
$ws.Cells["A1:N1"].AutoFilter = $true
$ws.Column(1).Width = 7.75; $ws.Column(2).Width = 19.33; $ws.Column(3).Width = 99.25; $ws.Column(4).Width = 11; $ws.Column(5).Width = 10.08; $ws.Column(6).Width = 33.16; $ws.Column(7).Width = 22.83; $ws.Column(12).Width = 30; $ws.Column(13).Width = 30;
$ws.Row(1).Height = 39; $ws.Row(1).CustomHeight = $true
for ($i = 2; $i -le ($dataRowCount + 1); $i++) { $ws.Row($i).CustomHeight = $true; $ws.Row($i).Height = 18 }
$headerRange = $ws.Cells["A1:N1"]; $headerRange.Style.Font.Name = 'ＭＳ Ｐゴシック'; $headerRange.Style.Font.Bold = $true; $headerRange.Style.WrapText = $true;
$dataRange = $ws.Cells["A2:N$($dataRowCount + 1)"]; $dataRange.Style.Font.Name = '游ゴシック'; $dataRange.Style.Font.Size = 11;
$ws.Cells[$totalRange].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells["A1:N1"].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium

# --- 5. Conditional Formatting ---
Write-Host "Step 5: Applying conditional formatting..."
$cfRange = $ws.Cells["D2:D$($dataRowCount + 1)"]
# Rule 1: >= 9.0 (Red)
$cf_red = $ws.ConditionalFormatting.AddGreaterThanOrEqual($cfRange); $cf_red.Formula = "9.0"; $cf_red.Style.Fill.PatternType = 'Solid'; $cf_red.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#ffc7ce'))
# Rule 2: Between 7.0 and 8.9 (Yellow)
$cf_yellow = $ws.ConditionalFormatting.AddBetween($cfRange); $cf_yellow.Formula = "7.0"; $cf_yellow.Formula2 = "8.9"; $cf_yellow.Style.Fill.PatternType = 'Solid'; $cf_yellow.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#ffeb9c'))
# Rule 3: Between 4.0 and 6.9 (Green)
$cf_green = $ws.ConditionalFormatting.AddBetween($cfRange); $cf_green.Formula = "4.0"; $cf_green.Formula2 = "6.9"; $cf_green.Style.Fill.PatternType = 'Solid'; $cf_green.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#c6efce'))

# --- 6. Finalization ---
Write-Host "Step 6: Saving and closing the file..."
$pkg.Save()
$pkg.Dispose()
Remove-Item .\temp_data.csv

Write-Host "--------------------------------------------------"
Write-Host "Success! The definitive script has been created as 'generate_excel.ps1'."
Write-Host "The generated Excel file is: $outputPath"
Write-Host "--------------------------------------------------"
