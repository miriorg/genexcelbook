#Requires -Module ImportExcel
#Requires -Version 7.0

<#
.SYNOPSIS
    CSVファイルを読み込み、Excel帳票生成用のデータ形式に変換します。
    (Step 2: Excelブック作成とデータ出力)

.DESCRIPTION
    指定されたCSVファイルを読み込み、以下のルールに基づいてバージョン情報の結合処理を行います。
    - [CurrentVersion] / [UpdateVersion]
    - [CurrentVersion] /
    - (空文字) ※Updateのみ、または両方なしの場合

    変換後のデータは PSCustomObject の配列として保持されます。
    変換後、指定された不要な列を削除し、Excelファイルを出力します。

.PARAMETER CsvPath
    読み込むCSVファイルのパス。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, ValueFromPipeline=$true, Position=0)]
    [string]$CsvPath
)

begin {
    Set-StrictMode -Version Latest

    # 処理結果を保持する配列
    $script:processedData = @()

    # 削除対象の列名（正規表現パターン）
    $script:removePatterns = @(
        '^\[?CurrentVersion\]?$',
        '^\[?UpdateVersion\]?$',
        '^\[?Note\]?$'
    )

    # 外部モジュールの確認
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Throw "Module 'ImportExcel' is not installed. Please install it using 'Install-Module ImportExcel'."
    }
}

process {
    # パイプライン入力または引数指定の処理
    if ([string]::IsNullOrEmpty($CsvPath)) {
        Write-Warning "CSV Path is not specified."
        return
    }

    if (-not (Test-Path $CsvPath)) {
        Write-Error "File not found: $CsvPath"
        return
    }

    try {
        Write-Host "Reading CSV file: $CsvPath" -ForegroundColor Cyan

        # CSV読み込み (UTF-8)
        $csvData = Import-Csv -Path $CsvPath -Encoding UTF8

        $processedData = @()

        foreach ($row in $csvData) {
            # バージョン情報の取得 (Set-StrictMode対策としてプロパティ存在確認を行う)
            $current = ''
            if ($row.PSObject.Properties.Match('CurrentVersion').Count -gt 0) {
                $current = $row.CurrentVersion
            }

            $update = ''
            if ($row.PSObject.Properties.Match('UpdateVersion').Count -gt 0) {
                $update = $row.UpdateVersion
            }

            # 変換ルールの適用
            $versionText = ''
            if (-not [string]::IsNullOrEmpty($current) -and -not [string]::IsNullOrEmpty($update)) {
                $versionText = "$current / $update"
            }
            elseif (-not [string]::IsNullOrEmpty($current)) {
                $versionText = "$current /"
            }
            # Updateのみ、または両方なしの場合は空文字のまま

            # 新しいオブジェクトの作成（元のプロパティを保持しつつ、新しい列を追加）
            $newRow = $row | Select-Object *

            # 計算した列を追加
            $newRow | Add-Member -MemberType NoteProperty -Name '現行バージョン／更新版バージョン' -Value $versionText -Force

            $script:processedData += $newRow
        }
    }
    catch {
        Write-Error "An error occurred during processing: $_"
    }
}

end {
    if ($script:processedData.Count -eq 0) {
        Write-Warning "No data to export."
        return
    }

    Write-Host "Data processing complete. Total rows: $($script:processedData.Count)" -ForegroundColor Green

    # --- Step 2: Excelブック作成とデータ出力 ---
    try {
        # EPPlusのロード (ImportExcelモジュールに含まれるDLLを利用)
        $dllPath = Join-Path (Get-Module -ListAvailable ImportExcel | Select-Object -First 1).ModuleBase "EPPlus.dll"
        Add-Type -Path $dllPath

        # 出力ファイル名の設定
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $outputPath = ".\vns_report_$timestamp.xlsx"
        $fileInfo = New-Object System.IO.FileInfo $outputPath
        if ($fileInfo.Exists) { Remove-Item $outputPath }

        Write-Host "Creating Excel package: $outputPath" -ForegroundColor Cyan
        $pkg = New-Object OfficeOpenXml.ExcelPackage $fileInfo

        # ワークシート追加
        $ws = $pkg.Workbook.Worksheets.Add('vns_確認')

        # データの書き込み
        # ヘッダー行（プロパティ名）
        $headers = $script:processedData[0].PSObject.Properties.Name
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $ws.Cells[1, ($i + 1)].Value = $headers[$i]
        }

        # データ行 (LoadFromCollectionのオーバーロード問題を回避するためループ処理に変更)
        for ($r = 0; $r -lt $script:processedData.Count; $r++) {
            $rowItem = $script:processedData[$r]
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $ws.Cells[($r + 2), ($c + 1)].Value = $rowItem.($headers[$c])
            }
        }

        # 不要な列の削除 (Excelシート上で削除を実行)
        # 列インデックスがずれないように、後ろの列から削除するために降順で処理する
        $columnsToDelete = @()
        for ($col = 1; $col -le $ws.Dimension.End.Column; $col++) {
            $headerVal = $ws.Cells[1, $col].Text
            foreach ($pattern in $script:removePatterns) {
                if ($headerVal -match $pattern) {
                    $columnsToDelete += $col
                    break
                }
            }
        }
        $columnsToDelete | Sort-Object -Descending | ForEach-Object { $ws.DeleteColumn($_) }

        $pkg.Save()
        $pkg.Dispose()
        Write-Host "Excel file created successfully: $outputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create Excel file: $_"
    }
}