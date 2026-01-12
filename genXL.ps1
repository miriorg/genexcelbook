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

            # Package列が '-Not affected-' の場合の処理
            $pkgProp = $newRow.PSObject.Properties.Match('Package') | Select-Object -First 1
            if ($pkgProp -and $pkgProp.Value -eq '-Not affected-') {
                # Package列を空文字に設定
                $pkgProp.Value = ''

                # 処理結果列（影響有無）を '影響なし' に設定
                $statusProp = $newRow.PSObject.Properties.Match('処理結果') | Select-Object -First 1
                if ($statusProp) {
                    $statusProp.Value = '影響なし'
                }
            }

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
                $cellValue = $rowItem.($headers[$c])
                # D列(3)・E列(4)で数値として扱える場合は数値型に変換してセット
                if (($c -eq 3 -or $c -eq 4) -and ($cellValue -as [double])) {
                    $cellValue = $cellValue -as [double]
                }
                $ws.Cells[($r + 2), ($c + 1)].Value = $cellValue
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

        # ヘッダー行の文字列を generate_excel.ps1 の correctHeaders に変更
        $correctHeaders = @(
            "[JVN]`n区分", "[JVN]`n番号", "[JVN]`nタイトル", "[JVN]`nCVSS v3`n深刻度", "[JVN]`nCVSS v2`n深刻度",
            "[JVN]`nCVSS v3 URL", "[JVN]`nCVSS v2 URL", "[JVN]`n更新日", "[JVN]`nCVE", "[Gauge]`nURL",
            "影響有無", "対象パッケージ／ソフトウェア", "現行バージョン／`n更新版バージョン", ""
        )
        for ($i = 0; $i -lt $correctHeaders.Count; $i++) {
            $ws.Cells[1, ($i + 1)].Value = $correctHeaders[$i]
        }

        # --- Step 3: 書式設定と条件付き書式 ---
        Write-Host "Applying formatting..." -ForegroundColor Cyan

        $fontMSPGothic = 'ＭＳ Ｐゴシック'
        $fontYuGothic = '游ゴシック'
        $dataRowCount = $script:processedData.Count
        $lastCol = $ws.Dimension.End.Column

        # オートフィルター
        $ws.Cells["A1:N1"].AutoFilter = $true

        # 列幅設定 (generate_excel.ps1 準拠)
        $ws.Column(1).Width = 7.75
        $ws.Column(2).Width = 19.33
        $ws.Column(3).Width = 99.25
        $ws.Column(4).Width = 11
        $ws.Column(5).Width = 10.08
        $ws.Column(6).Width = 33.16
        $ws.Column(7).Width = 22.83
        $ws.Column(8).Width = 22.83
        $ws.Column(9).Width = 19.33
        $ws.Column(10).Width = 33.16
        $ws.Column(11).Width = 29
        $ws.Column(12).Width = 30
        $ws.Column(13).Width = 34.14

        # 行の高さ
        $ws.Row(1).Height = 39
        $ws.Row(1).CustomHeight = $true
        for ($i = 2; $i -le ($dataRowCount + 1); $i++) {
            $ws.Row($i).CustomHeight = $true
            $ws.Row($i).Height = 18
        }

        # ヘッダー書式
        $headerRange = $ws.Cells[1, 1, 1, $lastCol]
        $headerRange.Style.Font.Name = $fontMSPGothic
        $headerRange.Style.Font.Bold = $true
        $headerRange.Style.WrapText = $true
        $headerRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $headerRange.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#FFFFFF'))
        $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#2D7DCE'))
        $headerRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium

        # データエリア書式と罫線
        $allRange = $ws.Cells[1, 1, ($dataRowCount + 1), $lastCol]
        $allRange.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $allRange.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $allRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $allRange.Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

        $dataRange = $ws.Cells[2, 1, ($dataRowCount + 1), $lastCol]
        $dataRange.Style.Font.Name = $fontYuGothic
        $dataRange.Style.Font.Size = 11

        # 条件付き書式 (K列: 影響有無 を対象)
        $cfRule1 = $ws.ConditionalFormatting.AddExpression($dataRange); $cfRule1.Formula = '$K2="修正版配布待ち"'; $cfRule1.Style.Fill.PatternType = 'Solid'; $cfRule1.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#F1A983')
        $cfRule2 = $ws.ConditionalFormatting.AddExpression($dataRange); $cfRule2.Formula = '$K2="修正版適用待ち"'; $cfRule2.Style.Fill.PatternType = 'Solid'; $cfRule2.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#F1A983')
        $cfRule3 = $ws.ConditionalFormatting.AddExpression($dataRange); $cfRule3.Formula = '$K2="更新版適用済"'; $cfRule3.Style.Fill.PatternType = 'Solid'; $cfRule3.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#92D050')

        # --- Step 4: 凡例の追加 ---
        Write-Host "Adding footer legend..." -ForegroundColor Cyan

        # Legend data based on vns-sample.xlsx
        $legendItems = @(
            @{ Row = 0; Description = '脆弱性が修正された更新版パッケージが適用済みであるもの'; Status = '更新版適用済'; Color = '#92D050' },
            @{ Row = 1; Description = '脆弱性が存在するものの、修正パッチがまだ配布されておらず、配布を待っている状態のもの。'; Status = '修正版配布待ち'; Color = '#F1A983' },
            @{ Row = 2; Description = '修正パッチは配布されているが、まだ適用ができていないもの。（再起動を行わないと適用ができないものなど）'; Status = '修正版適用待ち'; Color = '#F1A983' },
            @{ Row = 3; Description = '脆弱性が情報が存在するパッケージであるものの、該当パッケージをインストールしていないため、影響が発生しないもの。'; Status = '影響なし（未インストール）'; Color = $null },
            @{ Row = 4; Description = '脆弱性が情報が存在するパッケージであるものの、該当OSには影響がないことを公式にて確認されているもの。'; Status = '影響なし（範囲外）'; Color = $null },
            @{ Row = 5; Description = '該当OSでは利用されないパッケージ、ソフトウェアであるため、とくに確認対象とはならないもの。'; Status = '影響なし'; Color = $null }
        )

        $TopRow = $dataRowCount + 4 + ($legendItems | Measure-Object -Property Row -Minimum | Select-Object -ExpandProperty Minimum)
        $BottomRow = $dataRowCount + 4 + ($legendItems | Measure-Object -Property Row -Maximum | Select-Object -ExpandProperty Maximum)
        $TitleRow = $TopRow - 1
        $range="A$($TitleRow)"
        $ws.Cells[$range].Value = "凡例"
        $ws.Cells[$range].Style.Font.Bold = $true
        $ws.Cells[$range].Style.Font.Name = $fontMSPGothic
        $ws.Cells[$range].Style.Font.Size = 11
        $ws.Row($TitleRow).CustomHeight = $true;
        $ws.Row($TitleRow).Height = 13

        foreach ($item in $legendItems) {
            $row = $dataRowCount + 4 + $item.Row

            $ws.Row($TopRow).CustomHeight = $true;
            $ws.Row($TopRow).Height = 13

            # Merge cells D:J for the description, set value and style for proper display
            $range="A$($row):M$($row)"
            $rangeLeft="A$($row):C$($row)"
            $rangeCenter="D$($row):J$($row)"
            $rangeRight="K$($row):M$($row)"
            $ws.Cells[$rangeCenter].Merge = $true
            $ws.Cells[$rangeLeft].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
            $ws.Cells[$rangeCenter].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
            $ws.Cells[$rangeRight].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
            $ws.Cells[$rangeRight].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

            $ws.Cells[$rangeCenter].Value = $item.Description
            $ws.Cells[$rangeCenter].Style.WrapText = $true
            $ws.Cells[$rangeCenter].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

            # Add the status text to column K
            $ws.Cells["K$row"].Value = $item.Status

            $ws.Cells[$range].Style.Font.Name = $fontMSPGothic
            $ws.Cells[$range].Style.Font.Size = 11
            $ws.Cells[$range].Style.Font.Bold = $false
            # Set the background color for the status cell
            if ($item.Color) {
                $ws.Cells[$range].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[$range].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($item.Color))
            }
        }

        $range="A$($TopRow):M$($BottomRow)"
        $ws.Cells[$range].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Medium)

        $pkg.Save()
        $pkg.Dispose()
        Write-Host "Excel file created successfully: $outputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create Excel file: $_"
    }
}