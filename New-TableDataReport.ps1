#Requires -Version 5.1
<#
.SYNOPSIS
  A5:SQL Mk-2 のテーブルCSVから、件数・分布・キー先頭1桁・全件検索用のExcelを作る。
.DESCRIPTION
  指定ディレクトリ直下の *.csv（1テーブル1ファイル、見出し行、RFC 4180）を読む。
  文字コードは UTF-8（BOMなし可）/ UTF-16 / Shift_JIS を自動判定する。
  NULL は空欄、(NULL)、\N（ゼロ幅スペース付きも可）として数える。
  キー列は -KeyColumn が優先。無いときは列名（ID / コード / 番号 / キー）から推定し、
  それも無ければ先頭列を推定キーにする。集計はキー値の先頭1文字。
.PARAMETER InputDir
  CSVのあるディレクトリ。
.PARAMETER OutputPath
  出力xlsx。省略時は入力ディレクトリへ テーブル分析_yyyyMMdd_HHmmss.xlsx。
.PARAMETER KeyColumn
  キー列名。複数可。表に無い名前だけのときは推定に戻る。
.PARAMETER WarnNullRate
  この欠損率以上を注意にする。既定は 0.2。
.EXAMPLE
  .\New-TableDataReport.ps1 -InputDir D:\export\csv
.EXAMPLE
  .\New-TableDataReport.ps1 -InputDir D:\export\csv -KeyColumn 顧客コード,id
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputDir,
    [Parameter(Position = 1)]
    [string]$OutputPath,
    [string[]]$KeyColumn,
    [ValidateRange(0.0, 1.0)]
    [double]$WarnNullRate = 0.2
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

$FontName = 'Meiryo UI'
$ExcelMaxRows = 1048575
$InkHex = '333333'
$NavyHex = '1A2B4C'
$WarnHex = '800020'

function ConvertTo-OleColor([string]$Hex) {
    $h = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return ($r -bor ($g -shl 8) -bor ($b -shl 16))
}

function Get-TextEncoding([string]$Path) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $head = New-Object byte[] 4
        $n = $fs.Read($head, 0, 4)
        if ($n -ge 3 -and $head[0] -eq 239 -and $head[1] -eq 187 -and $head[2] -eq 191) {
            return New-Object System.Text.UTF8Encoding $true
        }
        if ($n -ge 2 -and $head[0] -eq 255 -and $head[1] -eq 254) {
            return [Text.Encoding]::Unicode
        }
        $fs.Position = 0
        $all = New-Object byte[] $fs.Length
        [void]$fs.Read($all, 0, $all.Length)
        $strict = New-Object System.Text.UTF8Encoding $false, $true
        try {
            [void]$strict.GetString($all)
            return New-Object System.Text.UTF8Encoding $false
        }
        catch {
            return [Text.Encoding]::GetEncoding(932)
        }
    }
    finally { $fs.Dispose() }
}

function Get-Plain([string]$Value) {
    if ($null -eq $Value) { return '' }
    return (($Value -replace "[\u200B\uFEFF]", '').Trim())
}

function Test-Missing([string]$Value) {
    $t = Get-Plain $Value
    return ($t -eq '' -or $t -eq '(NULL)' -or $t -eq '\N')
}

function Test-NumericText([string]$Text) {
    if ($Text -match '^[+-]?0\d') { return $false }
    if ($Text -match '^[+-]?\d{16,}(\.\d+)?$') { return $false }
    $n = 0.0
    return [double]::TryParse($Text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)
}

function Test-DateText([string]$Text) {
    return $Text -match '^\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$'
}

function Test-KeyName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $n = $Name.Trim()
    if ($n -match 'コード$|番号$|キー$') { return $true }
    if ($n -match '(?i)^(id|pk)$') { return $true }
    if ($n -match '(?i)(^|_)(id|code|cd|no|key|pk)$') { return $true }
    if ($n -match '(Id|ID)$') { return $true }
    return $false
}

function ConvertTo-CellString([string]$Value) {
    if ($null -eq $Value -or $Value -eq '') { return '' }
    $s = $Value -replace "[\u0000-\u0008\u000B\u000C\u000E-\u001F]", ''
    if ($s.Length -gt 32767) { return $s.Substring(0, 32767) }
    return $s
}

function Format-InvariantNumber($Value) {
    return ([double]$Value).ToString('0.####', [Globalization.CultureInfo]::InvariantCulture)
}

function Read-A5Csv([string]$Path) {
    $enc = Get-TextEncoding $Path
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, $enc)
    try {
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        if ($parser.EndOfData) { throw '空のCSVです。' }
        $headers = @($parser.ReadFields())
        if ($headers.Count -eq 0) { throw 'ヘッダがありません。' }
        $headers[0] = ([string]$headers[0]) -replace '^\uFEFF', ''
        $seen = @{}
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $h = [string]$headers[$i]
            if ([string]::IsNullOrWhiteSpace($h)) { $h = '列{0}' -f ($i + 1) }
            $key = $h
            $seq = 2
            while ($seen.ContainsKey($key)) { $key = '{0}_{1}' -f $h, $seq; $seq++ }
            $seen[$key] = $true
            $headers[$i] = $key
        }
        $colCount = $headers.Count
        $rows = New-Object 'System.Collections.Generic.List[object]'
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -eq $fields) { continue }
            $line = New-Object string[] $colCount
            for ($c = 0; $c -lt $colCount; $c++) {
                $line[$c] = if ($c -lt $fields.Length -and $null -ne $fields[$c]) { [string]$fields[$c] } else { '' }
            }
            $rows.Add($line)
        }
        return [pscustomobject]@{ Headers = $headers; Rows = $rows; EncodingName = $enc.WebName }
    }
    finally { $parser.Close() }
}

function Get-SafeSheetName([string]$Name, $Used) {
    $s = ($Name -replace '[:\\/?*\[\]]', ' ').Trim().Trim("'")
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'table' }
    if ($s.Length -gt 31) { $s = $s.Substring(0, 31).Trim() }
    $base = $s
    $i = 2
    while ($Used.Contains($s)) {
        $suffix = '_' + $i
        $keep = [Math]::Max(1, 31 - $suffix.Length)
        $stem = $base
        if ($stem.Length -gt $keep) { $stem = $stem.Substring(0, $keep).Trim() }
        $s = $stem + $suffix
        $i++
    }
    [void]$Used.Add($s)
    return $s
}

function Get-Status([int]$RowCount, [int]$KeyNulls, [double]$MaxNull, [bool]$Truncated, [string]$ErrorMessage, [double]$Rate) {
    if ($ErrorMessage) { return '注意: 読込失敗' }
    $parts = New-Object System.Collections.Generic.List[string]
    if ($RowCount -eq 0) { $parts.Add('空テーブル') }
    if ($KeyNulls -gt 0) { $parts.Add('キー欠損') }
    if ($RowCount -gt 0 -and $MaxNull -ge $Rate) { $parts.Add('欠損率') }
    if ($Truncated) { $parts.Add('行数上限') }
    if ($parts.Count -eq 0) { return '正常' }
    return '注意: ' + ($parts -join '・')
}

function Build-ReportModel($Files, $ExplicitKeys, [double]$Rate) {
    $tables = New-Object System.Collections.Generic.List[object]
    $profiles = New-Object System.Collections.Generic.List[object]
    $keyAggs = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Files) {
        $base = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        try {
            $csv = Read-A5Csv $file.FullName
        }
        catch {
            $tables.Add([pscustomobject]@{
                    Name = $base; File = $file.Name; EncodingName = ''; Headers = @(); Rows = $null
                    Error = $_.Exception.Message; KeyLabel = $_.Exception.Message; KeyKindText = ''
                    KeyNulls = 0; MaxNullRate = 0.0; Status = '注意: 読込失敗'; Truncated = $false; Estimated = $false
                })
            continue
        }
        $headers = $csv.Headers
        $rows = $csv.Rows
        $rowCount = $rows.Count
        $colCount = $headers.Count
        $keyIdx = New-Object System.Collections.Generic.List[int]
        $estimated = $false
        if ($ExplicitKeys -and $ExplicitKeys.Count -gt 0) {
            for ($c = 0; $c -lt $colCount; $c++) {
                if ($ExplicitKeys.Contains([string]$headers[$c])) { $keyIdx.Add($c) }
            }
        }
        if ($keyIdx.Count -eq 0) {
            for ($c = 0; $c -lt $colCount; $c++) {
                if (Test-KeyName ([string]$headers[$c])) { $keyIdx.Add($c) }
            }
            if ($ExplicitKeys -and $ExplicitKeys.Count -gt 0) { $estimated = $true }
        }
        if ($keyIdx.Count -eq 0 -and $colCount -gt 0) {
            $keyIdx.Add(0)
            $estimated = $true
        }
        $mark = ''
        if ($estimated) { $mark = '※' }
        $keyNames = @()
        foreach ($k in $keyIdx) { $keyNames += ([string]$headers[$k] + $mark) }

        $prefixMaps = @{}
        $nulls = New-Object int[] $colCount
        $nonNull = New-Object int[] $colCount
        $numCount = New-Object int[] $colCount
        $dateCount = New-Object int[] $colCount
        $numSum = New-Object double[] $colCount
        $numMin = New-Object double[] $colCount
        $numMax = New-Object double[] $colCount
        $hasNum = New-Object bool[] $colCount
        $freq = New-Object object[] $colCount
        $strMin = New-Object string[] $colCount
        $strMax = New-Object string[] $colCount
        for ($c = 0; $c -lt $colCount; $c++) {
            $freq[$c] = New-Object 'System.Collections.Generic.Dictionary[string,int]'
            if ($keyIdx.Contains($c)) { $prefixMaps[$c] = New-Object 'System.Collections.Generic.Dictionary[string,int]' }
        }
        foreach ($line in $rows) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $raw = [string]$line[$c]
                if (Test-Missing $raw) {
                    $nulls[$c]++
                    if ($prefixMaps.ContainsKey($c)) {
                        $map = $prefixMaps[$c]
                        if ($map.ContainsKey('（空）')) { $map['（空）']++ } else { $map['（空）'] = 1 }
                    }
                    continue
                }
                $plain = Get-Plain $raw
                $nonNull[$c]++
                $bag = $freq[$c]
                if ($bag.ContainsKey($plain)) { $bag[$plain]++ } else { $bag[$plain] = 1 }
                if ($null -eq $strMin[$c] -or [string]::CompareOrdinal($plain, $strMin[$c]) -lt 0) { $strMin[$c] = $plain }
                if ($null -eq $strMax[$c] -or [string]::CompareOrdinal($plain, $strMax[$c]) -gt 0) { $strMax[$c] = $plain }
                if (Test-NumericText $plain) {
                    $numCount[$c]++
                    $val = [double]::Parse($plain, [Globalization.CultureInfo]::InvariantCulture)
                    $numSum[$c] += $val
                    if (-not $hasNum[$c] -or $val -lt $numMin[$c]) { $numMin[$c] = $val }
                    if (-not $hasNum[$c] -or $val -gt $numMax[$c]) { $numMax[$c] = $val }
                    $hasNum[$c] = $true
                }
                elseif (Test-DateText $plain) { $dateCount[$c]++ }
                if ($prefixMaps.ContainsKey($c)) {
                    $prefix = $plain.Substring(0, 1)
                    $map = $prefixMaps[$c]
                    if ($map.ContainsKey($prefix)) { $map[$prefix]++ } else { $map[$prefix] = 1 }
                }
            }
        }

        $maxNull = 0.0
        $keyNulls = 0
        $kindParts = @()
        for ($c = 0; $c -lt $colCount; $c++) {
            $rateCol = 0.0
            if ($rowCount -gt 0) { $rateCol = $nulls[$c] / $rowCount }
            if ($rateCol -gt $maxNull) { $maxNull = $rateCol }
            $isKey = $keyIdx.Contains($c)
            if ($isKey -and $nulls[$c] -gt $keyNulls) { $keyNulls = $nulls[$c] }
            $typeName = '文字'
            if ($nonNull[$c] -eq 0) { $typeName = '空' }
            elseif ($numCount[$c] -eq $nonNull[$c]) { $typeName = '数値' }
            elseif ($dateCount[$c] -eq $nonNull[$c]) { $typeName = '日付' }
            $mode = ''
            $modeN = $null
            foreach ($pair in $freq[$c].GetEnumerator()) {
                if ($null -eq $modeN -or $pair.Value -gt $modeN -or ($pair.Value -eq $modeN -and [string]::CompareOrdinal($pair.Key, $mode) -lt 0)) {
                    $mode = $pair.Key
                    $modeN = $pair.Value
                }
            }
            $minText = ''
            $maxText = ''
            $avg = $null
            if ($typeName -eq '数値') {
                $minText = Format-InvariantNumber $numMin[$c]
                $maxText = Format-InvariantNumber $numMax[$c]
                $avg = $numSum[$c] / $nonNull[$c]
            }
            elseif ($nonNull[$c] -gt 0) {
                $minText = $strMin[$c]
                $maxText = $strMax[$c]
            }
            $uniqueRate = $null
            if ($nonNull[$c] -gt 0) { $uniqueRate = $freq[$c].Count / [double]$nonNull[$c] }
            $keyMark = ''
            if ($isKey) { $keyMark = $(if ($estimated) { '推定' } else { 'キー' }) }
            $profiles.Add([pscustomobject]@{
                    Table = $base; Column = [string]$headers[$c]; KeyMark = $keyMark; TypeName = $typeName
                    RowCount = $rowCount; NonNull = $nonNull[$c]; Nulls = $nulls[$c]; NullRate = $rateCol
                    Distinct = $freq[$c].Count; UniqueRate = $uniqueRate; Min = $minText; Max = $maxText
                    Avg = $avg; Mode = $mode; ModeCount = $modeN
                })
            if ($isKey) {
                $kindParts += ('{0}:{1}' -f $headers[$c], $prefixMaps[$c].Count)
                foreach ($pair in $prefixMaps[$c].GetEnumerator()) {
                    $ratio = 0.0
                    if ($rowCount -gt 0) { $ratio = $pair.Value / [double]$rowCount }
                    $keyAggs.Add([pscustomobject]@{
                            Table = $base; Key = [string]$headers[$c]; Prefix = [string]$pair.Key
                            Count = [int]$pair.Value; Ratio = $ratio
                            Status = $(if ($pair.Key -eq '（空）') { '注意' } else { '正常' })
                        })
                }
            }
        }
        $truncated = $rowCount -gt $ExcelMaxRows
        $tables.Add([pscustomobject]@{
                Name = $base; File = $file.Name; EncodingName = $csv.EncodingName
                Headers = $headers; Rows = $rows; Error = $null
                KeyLabel = ($keyNames -join ', '); KeyKindText = ($kindParts -join ' / ')
                KeyNulls = $keyNulls; MaxNullRate = $maxNull
                Status = (Get-Status $rowCount $keyNulls $maxNull $truncated $null $Rate)
                Truncated = $truncated; Estimated = $estimated
            })
    }
    $orderedKeys = @($keyAggs | Sort-Object Table, Key, @{ Expression = 'Count'; Descending = $true }, Prefix)
    return [pscustomobject]@{ Tables = $tables; Profiles = $profiles; Keys = $orderedKeys }
}

function Set-GridBorder($Range, [int]$Rows, [int]$Cols, [int]$Color) {
    foreach ($index in 7, 8, 9, 10) {
        $edge = $Range.Borders.Item($index)
        $edge.LineStyle = 1
        $edge.Weight = 2
        $edge.Color = $Color
    }
    if ($Cols -gt 1) {
        $edge = $Range.Borders.Item(11)
        $edge.LineStyle = 1
        $edge.Weight = 2
        $edge.Color = $Color
    }
    if ($Rows -gt 1) {
        $edge = $Range.Borders.Item(12)
        $edge.LineStyle = 1
        $edge.Weight = 2
        $edge.Color = $Color
    }
}

function Set-HeaderStyle($Range, [int]$Fill, [int]$FontColor) {
    $Range.Font.Name = $script:FontName
    $Range.Font.Size = 10
    $Range.Font.Bold = $true
    $Range.Font.Color = $FontColor
    $Range.Interior.Color = $Fill
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
    $Range.WrapText = $true
}

function Set-Page($Sheet, [string]$TitleRows, [bool]$FitWidth) {
    $ps = $Sheet.PageSetup
    try { $ps.Orientation = 2 } catch {}
    try { $ps.PaperSize = 9 } catch {}
    try { $ps.PrintTitleRows = $TitleRows } catch {}
    try { $ps.CenterHorizontally = $true } catch {}
    try { $ps.PrintGridlines = $false } catch {}
    try { $ps.LeftFooter = '&"Meiryo UI"&8A5:SQL Mk-2 テーブルデータ分析' } catch {}
    try { $ps.RightFooter = '&"Meiryo UI"&8&P / &N' } catch {}
    try {
        if ($FitWidth) {
            $ps.Zoom = $false
            $ps.FitToPagesWide = 1
            $ps.FitToPagesTall = 0
        }
        else { $ps.Zoom = 100 }
    }
    catch {}
}

function Write-Grid($Sheet, [int]$HeaderRow, [string[]]$Headers, $Records, [string[]]$Formats, [int]$Navy, [int]$Ink, [int]$Line, [int]$Zebra) {
    $colCount = $Headers.Count
    $rowCount = 0
    if ($null -ne $Records) { $rowCount = $Records.Count }
    $head = $Sheet.Range($Sheet.Cells.Item($HeaderRow, 1), $Sheet.Cells.Item($HeaderRow, $colCount))
    $headArr = New-Object 'object[,]' 1, $colCount
    for ($c = 0; $c -lt $colCount; $c++) { $headArr[0, $c] = $Headers[$c] }
    $head.NumberFormat = '@'
    $head.Value2 = $headArr
    Set-HeaderStyle $head $Navy 16777215
    $Sheet.Rows.Item($HeaderRow).RowHeight = 22
    $body = $null
    if ($rowCount -gt 0) {
        $body = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1, 1), $Sheet.Cells.Item($HeaderRow + $rowCount, $colCount))
        $body.Font.Name = $script:FontName
        $body.Font.Size = 10
        $body.Font.Color = $Ink
        $body.VerticalAlignment = -4108
        for ($c = 0; $c -lt $colCount; $c++) {
            $colRange = $Sheet.Range($Sheet.Cells.Item($HeaderRow + 1, $c + 1), $Sheet.Cells.Item($HeaderRow + $rowCount, $c + 1))
            $fmt = $Formats[$c]
            if ($fmt -eq '@') { $colRange.NumberFormat = '@' }
            $colArr = New-Object 'object[,]' $rowCount, 1
            for ($r = 0; $r -lt $rowCount; $r++) { $colArr[$r, 0] = $Records[$r][$c] }
            $colRange.Value2 = $colArr
            if ($fmt -ne '@') { $colRange.NumberFormat = $fmt }
        }
        $zebraRule = $body.FormatConditions.Add(2, 0, '=MOD(ROW(),2)=0')
        $zebraRule.Interior.Color = $Zebra
        $all = $Sheet.Range($Sheet.Cells.Item($HeaderRow, 1), $Sheet.Cells.Item($HeaderRow + $rowCount, $colCount))
    }
    else { $all = $head }
    Set-GridBorder $all ($rowCount + 1) $colCount $Line
    $all.AutoFilter() | Out-Null
    $all.EntireColumn.AutoFit() | Out-Null
    for ($c = 1; $c -le $colCount; $c++) {
        $column = $Sheet.Columns.Item($c)
        if ($column.ColumnWidth -gt 46) { $column.ColumnWidth = 46 }
        elseif ($column.ColumnWidth -lt 9) { $column.ColumnWidth = 9 }
    }
    $Sheet.Activate() | Out-Null
    $window = $Sheet.Application.ActiveWindow
    if ($window.FreezePanes) { $window.FreezePanes = $false }
    $window.SplitColumn = 0
    $window.SplitRow = $HeaderRow
    $window.FreezePanes = $true
    $window.DisplayGridlines = $false
    $window.ScrollRow = 1
    $window.ScrollColumn = 1
    return $body
}

function Set-WarnFont($Cell, [int]$Color) {
    $Cell.Font.Name = $script:FontName
    $Cell.Font.Color = $Color
    $Cell.Font.Bold = $true
}

function Write-ReportBook($Model, [string]$OutPath, [string]$SourceDir, [double]$Rate) {
    $navy = ConvertTo-OleColor $script:NavyHex
    $ink = ConvertTo-OleColor $script:InkHex
    $warn = ConvertTo-OleColor $script:WarnHex
    $border = ConvertTo-OleColor 'D9D9D9'
    $zebra = ConvertTo-OleColor 'F4F5F7'
    $label = ConvertTo-OleColor 'EEF0F4'
    $muted = ConvertTo-OleColor '5C6570'
    $excel = $null
    $wb = $null
    $saved = $false
    $excelPid = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents = $false
        $excelPid = 0
        try {
            if (-not ('Win32.NativeMethods' -as [type])) {
                Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);'
            }
            $procId = [uint32]0
            [void][Win32.NativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$procId)
            $excelPid = [int]$procId
        }
        catch { $excelPid = 0 }
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $normal = $null
        foreach ($styleName in @('Normal', '標準')) {
            try { $normal = $wb.Styles.Item($styleName); break } catch { $normal = $null }
        }
        if ($normal) {
            try { $normal.Font.Name = $script:FontName } catch {}
            try { $normal.Font.Size = 10 } catch {}
            try { $normal.Font.Color = $ink } catch {}
            try { $normal.VerticalAlignment = -4108 } catch {}
        }

        $usedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $summary = $wb.Worksheets.Item(1)
        $summary.Name = '概要'
        [void]$usedNames.Add('概要')
        $dist = $wb.Worksheets.Add([Type]::Missing, $summary)
        $dist.Name = '分布'
        [void]$usedNames.Add('分布')
        $keySheet = $wb.Worksheets.Add([Type]::Missing, $dist)
        $keySheet.Name = 'キー集計'
        [void]$usedNames.Add('キー集計')

        $summaryHeaders = @('テーブル', 'ファイル', '文字コード', '行数', '列数', 'キー項目', '先頭1桁の種類', 'キー欠損', '最大欠損率', '状況')
        $summaryFormats = @('@', '@', '@', '#,##0', '#,##0', '@', '@', '#,##0', '0.0%', '@')
        $summaryRows = New-Object System.Collections.Generic.List[object]
        foreach ($t in $Model.Tables) {
            $nRows = 0
            $nCols = @($t.Headers).Count
            if ($t.Rows) { $nRows = $t.Rows.Count }
            $summaryRows.Add(@(
                    (ConvertTo-CellString $t.Name), (ConvertTo-CellString $t.File), (ConvertTo-CellString $t.EncodingName),
                    [double]$nRows, [double]$nCols, (ConvertTo-CellString $t.KeyLabel), (ConvertTo-CellString $t.KeyKindText),
                    [double]$t.KeyNulls, [double]$t.MaxNullRate, (ConvertTo-CellString $t.Status)
                ))
        }
        $totalRows = 0
        $warnTables = 0
        foreach ($t in $Model.Tables) {
            if ($t.Rows) { $totalRows += $t.Rows.Count }
            if ($t.Status -ne '正常') { $warnTables++ }
        }

        $headerRow = 8
        [void](Write-Grid $summary $headerRow $summaryHeaders $summaryRows $summaryFormats $navy $ink $border $zebra)
        for ($i = 0; $i -lt $summaryRows.Count; $i++) {
            $excelRow = $headerRow + 1 + $i
            $rec = $summaryRows[$i]
            if ([string]$rec[9] -like '*空テーブル*') { Set-WarnFont $summary.Cells.Item($excelRow, 4) $warn }
            if ([double]$rec[7] -gt 0) { Set-WarnFont $summary.Cells.Item($excelRow, 8) $warn }
            if ([double]$rec[8] -ge $Rate -and [double]$rec[3] -gt 0) { Set-WarnFont $summary.Cells.Item($excelRow, 9) $warn }
            if ([string]$rec[9] -ne '正常') { Set-WarnFont $summary.Cells.Item($excelRow, 10) $warn }
        }

        $summary.Rows.Item(1).RowHeight = 30
        $title = $summary.Range('A1:J1')
        $title.Merge() | Out-Null
        $title.Value2 = 'テーブルデータ分析レポート'
        $title.Font.Name = $script:FontName
        $title.Font.Size = 18
        $title.Font.Bold = $true
        $title.Font.Color = 16777215
        $title.Interior.Color = $navy
        $title.HorizontalAlignment = -4108
        $title.VerticalAlignment = -4108
        $sub = $summary.Range('A2:J2')
        $sub.Merge() | Out-Null
        $sub.Value2 = 'A5:SQL Mk-2 CSV / Aurora MySQL 8.4    件数・列の分布・キー先頭1桁・全件データ'
        $sub.Font.Name = $script:FontName
        $sub.Font.Size = 9
        $sub.Font.Color = $muted
        $sub.VerticalAlignment = -4108
        $summary.Rows.Item(2).RowHeight = 18

        $summary.Range('A3').Value2 = '見出し'
        Set-HeaderStyle $summary.Range('A3') $navy 16777215
        $summary.Range('B3').Value2 = '本文'
        $summary.Range('B3').Font.Name = $script:FontName
        $summary.Range('B3').Font.Color = $ink
        $summary.Range('B3').HorizontalAlignment = -4108
        $summary.Range('C3').Value2 = '注意'
        Set-WarnFont $summary.Range('C3') $warn
        $summary.Range('C3').HorizontalAlignment = -4108
        $paletteNote = $summary.Range('D3:J3')
        $paletteNote.Merge() | Out-Null
        $paletteNote.Value2 = 'データシートは全件をフィルター付きで収録。ブック全体は Ctrl+F の検索場所を「ブック」にする。'
        $paletteNote.Font.Name = $script:FontName
        $paletteNote.Font.Size = 9
        $paletteNote.Font.Color = $muted
        $paletteNote.VerticalAlignment = -4108

        $meta = @(
            @('A5', '作成日時', (Get-Date -Format 'yyyy-MM-dd HH:mm'), '@'),
            @('C5', '入力フォルダ', $SourceDir, '@'),
            @('A6', 'CSVファイル数', [double]$Model.Tables.Count, '#,##0'),
            @('C6', '総データ行数', [double]$totalRows, '#,##0'),
            @('E6', '注意テーブル', [double]$warnTables, '#,##0'),
            @('G6', '注意閾値', [double]$Rate, '0%')
        )
        foreach ($item in $meta) {
            $lab = $summary.Range($item[0])
            $lab.Value2 = $item[1]
            $lab.Font.Name = $script:FontName
            $lab.Font.Bold = $true
            $lab.Font.Color = $navy
            $lab.Interior.Color = $label
            $lab.HorizontalAlignment = -4108
            $valCell = $lab.Offset(0, 1)
            $valCell.NumberFormat = $item[3]
            $valCell.Value2 = $item[2]
            $valCell.Font.Name = $script:FontName
            $valCell.Font.Color = $ink
        }
        if ($warnTables -gt 0) { Set-WarnFont $summary.Range('F6') $warn }
        $summary.Range('D5:J5').Merge() | Out-Null
        $summary.Range('D5').Font.Name = $script:FontName
        $summary.Range('D5').Font.Color = $ink
        $noteRow = $headerRow + $summaryRows.Count + 2
        $note = $summary.Range($summary.Cells.Item($noteRow, 1), $summary.Cells.Item($noteRow, 10))
        $note.Merge() | Out-Null
        $thresholdText = $Rate.ToString('P0', [Globalization.CultureInfo]::GetCultureInfo('ja-JP'))
        $note.Value2 = "※は列名から推定したキー。-KeyColumn で指定できる。空テーブル、キー欠損、欠損率 $thresholdText 以上、Excelの行数上限を注意で示す。件数と分布とキー集計はCSVの全行が対象。行数上限を超えたデータシートだけ先頭 $ExcelMaxRows 行。"
        $note.Font.Name = $script:FontName
        $note.Font.Size = 9
        $note.Font.Color = $muted
        $note.WrapText = $true
        $summary.Rows.Item($noteRow).RowHeight = 32
        $summary.Tab.Color = $navy
        Set-Page $summary '$8:$8' $true

        $distHeaders = @('テーブル', '列名', 'キー', '型', '行数', '非欠損', '欠損', '欠損率', 'ユニーク', 'ユニーク率', '最小', '最大', '平均', '最頻値', '最頻件数')
        $distFormats = @('@', '@', '@', '@', '#,##0', '#,##0', '#,##0', '0.0%', '#,##0', '0.0%', '@', '@', '#,##0.00', '@', '#,##0')
        $distRows = New-Object System.Collections.Generic.List[object]
        foreach ($p in $Model.Profiles) {
            $distRows.Add(@(
                    (ConvertTo-CellString $p.Table), (ConvertTo-CellString $p.Column), (ConvertTo-CellString $p.KeyMark),
                    (ConvertTo-CellString $p.TypeName), [double]$p.RowCount, [double]$p.NonNull, [double]$p.Nulls, [double]$p.NullRate,
                    [double]$p.Distinct, $(if ($null -eq $p.UniqueRate) { $null } else { [double]$p.UniqueRate }),
                    (ConvertTo-CellString $p.Min), (ConvertTo-CellString $p.Max),
                    $(if ($null -eq $p.Avg) { $null } else { [double]$p.Avg }),
                    (ConvertTo-CellString $p.Mode),
                    $(if ($null -eq $p.ModeCount) { $null } else { [double]$p.ModeCount })
                ))
        }
        [void](Write-Grid $dist 1 $distHeaders $distRows $distFormats $navy $ink $border $zebra)
        if ($distRows.Count -gt 0) {
            $rateRange = $dist.Range($dist.Cells.Item(2, 8), $dist.Cells.Item(1 + $distRows.Count, 8))
            $rule = $rateRange.FormatConditions.Add(1, 7, $Rate)
            $rule.Font.Color = $warn
            $rule.Font.Bold = $true
            $nullRange = $dist.Range($dist.Cells.Item(2, 7), $dist.Cells.Item(1 + $distRows.Count, 7))
            $nullRule = $nullRange.FormatConditions.Add(2, 0, '=AND(H2>=' + $Rate.ToString([Globalization.CultureInfo]::InvariantCulture) + ',E2>0)')
            $nullRule.Font.Color = $warn
            $nullRule.Font.Bold = $true
        }
        $dist.Tab.Color = $muted
        Set-Page $dist '$1:$1' $true

        $keyHeaders = @('テーブル', 'キー項目', '先頭1桁', '件数', '構成比', '状況')
        $keyFormats = @('@', '@', '@', '#,##0', '0.0%', '@')
        $keyRows = New-Object System.Collections.Generic.List[object]
        foreach ($k in $Model.Keys) {
            $keyRows.Add(@(
                    (ConvertTo-CellString $k.Table), (ConvertTo-CellString $k.Key), (ConvertTo-CellString $k.Prefix),
                    [double]$k.Count, [double]$k.Ratio, (ConvertTo-CellString $k.Status)
                ))
        }
        [void](Write-Grid $keySheet 1 $keyHeaders $keyRows $keyFormats $navy $ink $border $zebra)
        for ($i = 0; $i -lt $keyRows.Count; $i++) {
            if ([string]$keyRows[$i][5] -eq '注意') {
                Set-WarnFont $keySheet.Cells.Item($i + 2, 3) $warn
                Set-WarnFont $keySheet.Cells.Item($i + 2, 6) $warn
            }
        }
        if ($keyRows.Count -gt 0) {
            $bar = $keySheet.Range($keySheet.Cells.Item(2, 4), $keySheet.Cells.Item(1 + $keyRows.Count, 4)).FormatConditions.AddDatabar()
            $bar.BarColor.Color = $navy
            $bar.ShowValue = $true
        }
        $keySheet.Tab.Color = $navy
        Set-Page $keySheet '$1:$1' $true

        foreach ($t in $Model.Tables) {
            if ($t.Error) { continue }
            $sheetName = Get-SafeSheetName $t.Name $usedNames
            $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
            $ws.Name = $sheetName
            $colCount = @($t.Headers).Count
            $formats = New-Object string[] $colCount
            $headers = New-Object string[] $colCount
            for ($c = 0; $c -lt $colCount; $c++) {
                $formats[$c] = '@'
                $headers[$c] = [string]$t.Headers[$c]
            }
            $source = $t.Rows
            if ($t.Truncated) { $source = $t.Rows.GetRange(0, $ExcelMaxRows) }
            $matrix = New-Object System.Collections.Generic.List[object]
            foreach ($record in $source) {
                $out = New-Object object[] $colCount
                for ($c = 0; $c -lt $colCount; $c++) { $out[$c] = ConvertTo-CellString ([string]$record[$c]) }
                $matrix.Add($out)
            }
            [void](Write-Grid $ws 1 $headers $matrix $formats $navy $ink $border $zebra)
            $ws.Tab.Color = $(if ($t.Status -eq '正常') { $border } else { $warn })
            Set-Page $ws '$1:$1' ($colCount -le 12)
        }

        $excel.ScreenUpdating = $true
        $summary.Activate() | Out-Null
        $summary.Range('A1').Select() | Out-Null
        $excel.ActiveWindow.ScrollRow = 1
        $excel.ActiveWindow.ScrollColumn = 1
        if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
        $wb.SaveAs($OutPath, 51)
        $saved = $true
        $wb.Close($true)
    }
    finally {
        if ($wb -and -not $saved) { try { $wb.Close($false) } catch {} }
        if ($excel) {
            try { $excel.ScreenUpdating = $true } catch {}
            try { $excel.Quit() } catch {}
        }
        foreach ($obj in @($wb, $excel)) {
            if ($obj) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch {} }
        }
        $summary = $null
        $dist = $null
        $keySheet = $null
        $ws = $null
        $normal = $null
        $wb = $null
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($excelPid) {
            $leftover = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($leftover -and $leftover.ProcessName -eq 'EXCEL' -and [string]::IsNullOrEmpty($leftover.MainWindowTitle)) {
                try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) {
    throw "入力ディレクトリが見つかりません: $InputDir"
}
$InputDir = (Resolve-Path -LiteralPath $InputDir).Path
$files = @(Get-ChildItem -LiteralPath $InputDir -Filter *.csv -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "CSVファイルがありません: $InputDir" }
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $InputDir ('テーブル分析_{0:yyyyMMdd_HHmmss}.xlsx' -f (Get-Date))
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$outDir = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$explicit = $null
if ($KeyColumn -and $KeyColumn.Count -gt 0) {
    $explicit = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $KeyColumn) { if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$explicit.Add($name.Trim()) } }
}
Write-Host ('読込: {0} ファイル' -f $files.Count)
$model = Build-ReportModel $files $explicit $WarnNullRate
Write-ReportBook $model $OutputPath $InputDir $WarnNullRate
$warnCount = @($model.Tables | Where-Object { $_.Status -ne '正常' }).Count
$rowCount = 0
foreach ($t in $model.Tables) { if ($t.Rows) { $rowCount += $t.Rows.Count } }
Write-Host ('テーブル: {0}  行: {1}  注意: {2}' -f $model.Tables.Count, $rowCount, $warnCount)
Write-Host ('出力: {0}' -f $OutputPath)
