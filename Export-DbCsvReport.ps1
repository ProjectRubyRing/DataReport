param(
    [string]$CsvDir = (Join-Path $PSScriptRoot 'csv'),
    [string]$OutFile = (Join-Path $PSScriptRoot 'MySQL_Data_Report.xlsx'),
    [string]$KeyColumn = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

if (-not (Test-Path $CsvDir)) { $CsvDir = $PSScriptRoot }
$csvFiles = @(Get-ChildItem -Path $CsvDir -Filter '*.csv' -File | Sort-Object Name)
if (-not $csvFiles) { throw "CSVファイルが見つかりません: $CsvDir" }

$utf8     = [Text.UTF8Encoding]::new($false, $true)
$cp932    = [Text.Encoding]::GetEncoding(932)
$IsNull   = { param($v) [string]::IsNullOrWhiteSpace([string]$v) -or [string]$v -ceq 'NULL' }
$MakeBar  = { param([double]$r) ('█' * [int][Math]::Round([Math]::Min(1.0, [Math]::Max(0.0, $r)) * 12)).PadRight(12, '░') }
$SortVals = { param($a) @($a | Sort-Object { $d = 0.0; if ([double]::TryParse($_, [ref]$d)) { $d } else { $_ } }) }
$EscXml   = { param([string]$s) [Security.SecurityElement]::Escape(($s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')) }

# 1. 全CSV読み込み (UTF-8 / Shift-JIS 自動判別)
$tables = @(foreach ($f in $csvFiles) {
    $b = [IO.File]::ReadAllBytes($f.FullName)
    $raw = $(try { $utf8.GetString($b) } catch { $cp932.GetString($b) }).TrimStart([char]0xFEFF)
    $rows = @(ConvertFrom-Csv $raw)
    $cols = @(if ($rows.Count) { $rows[0].PSObject.Properties.Name } elseif ($raw.Trim()) { (ConvertFrom-Csv ($raw.Trim() + "`n1")).PSObject.Properties.Name })
    $key = if ($KeyColumn -and $cols -contains $KeyColumn) { $KeyColumn } elseif ($cols) { $cols[0] } else { '' }
    [PSCustomObject]@{ Name = $f.BaseName; Rows = $rows; Cols = $cols; Key = $key }
})
$totalRows = [int](($tables | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum)

# 2. 集計データ構築 (1:テーブル件数・分布 / 2:カラム別分布 / 3:キー前1桁集計 / 4:全データ検索)
$tblSummary, $colDist, $prefixAgg, $allSearch = 1..4 | ForEach-Object { ,([Collections.Generic.List[object[]]]::new()) }

foreach ($t in $tables) {
    $n, $cCount, $totalNulls = $t.Rows.Count, $t.Cols.Count, 0
    $ratio = if ($totalRows) { $n / $totalRows } else { 0.0 }

    foreach ($c in $t.Cols) {
        $vals = @($t.Rows | ForEach-Object { [string]$_.$c } | Where-Object { -not (&$IsNull $_) })
        $nulls = $n - $vals.Count; $totalNulls += $nulls
        $nullRate = if ($n) { $nulls / $n } else { 0.0 }
        $sorted = @(&$SortVals $vals)
        $uniq = @($vals | Select-Object -Unique).Count
        $top = if ($vals) { $g = $vals | Group-Object | Sort-Object Count -Descending | Select-Object -First 1; "$($g.Name) ($($g.Count))" } else { '-' }
        $st = if ($n -eq 0) { '【警告】データ0件' } elseif ($nulls -eq $n) { '【警告】全件NULL' } elseif ($nullRate -ge 0.2) { '【注意】NULL率高' } elseif ($uniq -eq 1 -and $n -gt 1) { '【注意】単一値のみ' } else { '正常' }
        $colDist.Add(@($t.Name, $c, $n, $vals.Count, $nulls, $nullRate, $uniq, $(if ($sorted) { $sorted[0] } else { '-' }), $(if ($sorted) { $sorted[-1] } else { '-' }), $top, $st))
    }

    $keyVals = @(if ($t.Key) { $t.Rows | ForEach-Object { [string]$_.($t.Key) } | Where-Object { -not (&$IsNull $_) } })
    $keyUniq = @($keyVals | Select-Object -Unique).Count
    $keyNulls = $n - $keyVals.Count
    $tblNullRate = if ($n * $cCount) { $totalNulls / ($n * $cCount) } else { 0.0 }
    $tSt = if ($n -eq 0) { '【警告】データ0件' } elseif ($keyNulls -gt 0) { "【警告】キー欠損($keyNulls)" } elseif ($keyUniq -lt $n) { "【注意】キー重複($($n - $keyUniq))" } elseif ($tblNullRate -ge 0.2) { '【注意】NULL率高' } else { '正常' }
    $tblSummary.Add(@($t.Name, $n, $ratio, $cCount, $t.Key, $keyUniq, $tblNullRate, (&$MakeBar $ratio), $tSt))

    if ($n -eq 0) {
        $prefixAgg.Add(@($t.Name, $t.Key, '-', 0, 0.0, 0.0, (&$MakeBar 0), '-', '-', '【警告】データ0件'))
    } else {
        $groups = @($t.Rows | Group-Object { $kv = if ($t.Key) { [string]$_.($t.Key) } else { '' }; if (&$IsNull $kv) { '(NULL/空)' } else { $kv.Trim().Substring(0, 1) } } | Sort-Object Name)
        foreach ($g in $groups) {
            $rTbl, $rAll = ($g.Count / $n), $(if ($totalRows) { $g.Count / $totalRows } else { 0.0 })
            $kSorted = @(&$SortVals ($g.Group | ForEach-Object { [string]$_.($t.Key) } | Where-Object { -not (&$IsNull $_) }))
            $pSt = if ($g.Name -eq '(NULL/空)') { '【警告】キー欠損' } elseif ($rTbl -ge 0.8 -and $n -ge 5) { '【注意】特定カテゴリ偏重' } else { '正常' }
            $prefixAgg.Add(@($t.Name, $t.Key, $g.Name, $g.Count, $rTbl, $rAll, (&$MakeBar $rTbl), $(if ($kSorted) { $kSorted[0] } else { '-' }), $(if ($kSorted) { $kSorted[-1] } else { '-' }), $pSt))
        }
    }

    $idx = 0
    foreach ($r in $t.Rows) {
        $idx++
        $kv = if ($t.Key) { [string]$r.($t.Key) } else { '' }
        $pfx = if (&$IsNull $kv) { '(NULL/空)' } else { $kv.Trim().Substring(0, 1) }
        $rNulls = @($t.Cols | Where-Object { &$IsNull $r.$_ }).Count
        $joined = ($t.Cols | ForEach-Object { "$_=$($r.$_)" }) -join ' | '
        $rSt = if (&$IsNull $kv) { '【警告】キー欠損' } elseif ($rNulls -gt 0) { "【注意】NULLあり($rNulls)" } else { '正常' }
        $allSearch.Add(@($t.Name, $idx, $t.Key, $kv, $pfx, $rNulls, $joined, $rSt))
    }
}

# 3. OpenXML (.xlsx) 構築 (Meiryo UI / 文字:#333333 / アクセント:#1A2B4C / 警告:#800020 / モノトーン)
$sheets = @(
    @{ Name='1_データ件数・分布'; Title="テーブル別データ件数・分布状況 (総テーブル数: $($tables.Count) / 総レコード数: $($totalRows.ToString('#,##0')))";
       Headers=@('テーブル名','データ件数','構成比','カラム数','キー項目','キー種類数','NULL/空欄率','分布状況','判定');
       Rows=$tblSummary; Fmts=@('@','#,##0','0.0%','#,##0','@','#,##0','0.0%','@','@'); Center=@(5,8,9) },
    @{ Name='2_カラム別分布'; Title='カラム別データ分布・充足状況一覧';
       Headers=@('テーブル名','カラム名','総件数','有効件数','NULL/空欄数','NULL率','ユニーク数','最小値','最大値','最頻値(Top1)','判定');
       Rows=$colDist; Fmts=@('@','@','#,##0','#,##0','#,##0','0.0%','#,##0','@','@','@','@'); Center=@(11) },
    @{ Name='3_前1桁カテゴリ集計'; Title='キー項目・前1桁カテゴリ別件数集計';
       Headers=@('テーブル名','キー項目','前1桁カテゴリ','件数','テーブル内構成比','全体構成比','分布状況','最小キー','最大キー','判定');
       Rows=$prefixAgg; Fmts=@('@','@','@','#,##0','0.0%','0.0%','@','@','@','@'); Center=@(3,7,10) },
    @{ Name='4_全データ検索'; Title="全テーブル横断データ検索 (オートフィルタ・Ctrl+F 検索対応 / 総件数: $($totalRows.ToString('#,##0')))";
       Headers=@('テーブル名','行No','キー項目','キー値','前1桁カテゴリ','NULL項目数','全カラム統合データ (Key=Value検索用)','判定');
       Rows=$allSearch; Fmts=@('@','#,##0','@','@','@','#,##0','@','@'); Center=@(5,8) }
)

$outPath = [IO.Path]::GetFullPath($OutFile)
if (Test-Path $outPath) { Remove-Item $outPath -Force }
$zip = [IO.Compression.ZipFile]::Open($outPath, 'Create')
$AddXml = { param([string]$EntryPath, [string]$Xml)
    $e = $zip.CreateEntry($EntryPath)
    $w = [IO.StreamWriter]::new($e.Open(), $utf8)
    $w.Write($Xml); $w.Dispose()
}

try {
    &$AddXml '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + ((1..$sheets.Count | ForEach-Object { "<Override PartName=`"/xl/worksheets/sheet$_.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }) -join '') + '</Types>')
    &$AddXml '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    &$AddXml 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + ((0..($sheets.Count - 1) | ForEach-Object { "<sheet name=`"$(&$EscXml $sheets[$_].Name)`" sheetId=`"$($_ + 1)`" r:id=`"rId$($_ + 1)`"/>" }) -join '') + '</sheets></workbook>')
    &$AddXml 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ((1..$sheets.Count | ForEach-Object { "<Relationship Id=`"rId$_`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$_.xml`"/>" }) -join '') + "<Relationship Id=`"rId$($sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>")

    # スタイル定義: xf 0=標準, 1=タイトル(#1A2B4C), 2=ヘッダ(#1A2B4C背景+#FFFFFF文字), 3..18=データセル(Normal/Warn × 白/Zebra × LeftText/CenterText/Int/Pct)
    $xfs = @(
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment vertical="center"/></xf>',
        '<xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    )
    foreach ($fId in @(0, 2, 3)) { # 0:通常(#333333), 2:警告(#800020), 3:警告太字(#800020 Bold)
        foreach ($fillId in @(0, 3)) { # 0:白, 3:モノトーン淡色(#F5F6F8)
            $xfs += "<xf numFmtId=`"49`" fontId=`"$fId`" fillId=`"$fillId`" borderId=`"1`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"left`" vertical=`"center`"/></xf>"
            $xfs += "<xf numFmtId=`"49`" fontId=`"$fId`" fillId=`"$fillId`" borderId=`"1`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"center`" vertical=`"center`"/></xf>"
            $xfs += "<xf numFmtId=`"3`" fontId=`"$fId`" fillId=`"$fillId`" borderId=`"1`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"right`" vertical=`"center`"/></xf>"
            $xfs += "<xf numFmtId=`"164`" fontId=`"$fId`" fillId=`"$fillId`" borderId=`"1`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"right`" vertical=`"center`"/></xf>"
        }
    }
    &$AddXml 'xl/styles.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts count="5"><font><sz val="10"/><color rgb="FF333333"/><name val="Meiryo UI"/><charset val="128"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><charset val="128"/></font><font><sz val="10"/><color rgb="FF800020"/><name val="Meiryo UI"/><charset val="128"/></font><font><b/><sz val="10"/><color rgb="FF800020"/><name val="Meiryo UI"/><charset val="128"/></font><font><b/><sz val="13"/><color rgb="FF1A2B4C"/><name val="Meiryo UI"/><charset val="128"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1A2B4C"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF5F6F8"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + $xfs.Count + '">' + ($xfs -join '') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $s = $sheets[$i]; $cCount = $s.Headers.Count; $rCount = $s.Rows.Count
        $lastCol = [char]([int][char]'A' + $cCount - 1)
        $widths = @(for ($c = 0; $c -lt $cCount; $c++) { $cp932.GetByteCount([string]$s.Headers[$c]) })
        foreach ($row in $s.Rows) {
            for ($c = 0; $c -lt $cCount; $c++) {
                $txt = if ($s.Fmts[$c] -eq '0.0%') { ([double]$row[$c]).ToString('0.0%') } elseif ($s.Fmts[$c] -eq '#,##0') { ([double]$row[$c]).ToString('#,##0') } else { [string]$row[$c] }
                $len = $cp932.GetByteCount($txt)
                if ($len -gt $widths[$c]) { $widths[$c] = $len }
            }
        }
        $colsXml = (0..($cCount - 1) | ForEach-Object { "<col min=`"$($_ + 1)`" max=`"$($_ + 1)`" width=`"$([Math]::Min(65, [Math]::Max(12, $widths[$_] + 5)))`" customWidth=`"1`"/>" }) -join ''

        $sb = [Text.StringBuilder]::new()
        $null = $sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="19"/><cols>').Append($colsXml).Append('</cols><sheetData>')
        $null = $sb.Append("<row r=`"1`" ht=`"26`" customHeight=`"1`"><c r=`"A1`" s=`"1`" t=`"inlineStr`"><is><t>$(&$EscXml $s.Title)</t></is></c></row>")
        $fStr = "&quot;表示件数: &quot;&amp;TEXT(SUBTOTAL(103,A4:A$([Math]::Max(4, 3 + $rCount))),&quot;#,##0&quot;)&amp;&quot; / $($rCount.ToString('#,##0')) 件&quot;"
        $null = $sb.Append("<row r=`"2`" ht=`"18`" customHeight=`"1`"><c r=`"A2`" s=`"0`" t=`"str`"><f>$fStr</f></c></row>")

        $null = $sb.Append('<row r="3" ht="24" customHeight="1">')
        for ($c = 0; $c -lt $cCount; $c++) {
            $ref = "$([char]([int][char]'A' + $c))3"
            $null = $sb.Append("<c r=`"$ref`" s=`"2`" t=`"inlineStr`"><is><t>$(&$EscXml $s.Headers[$c])</t></is></c>")
        }
        $null = $sb.Append('</row>')

        for ($r = 0; $r -lt $rCount; $r++) {
            $rowNum = 4 + $r
            $rowObj = $s.Rows[$r]
            $isWarn = ([string]$rowObj[$cCount - 1] -match '^【(警告|注意)】')
            $zebra = ($r % 2)
            $null = $sb.Append("<row r=`"$rowNum`" ht=`"19`" customHeight=`"1`">")
            for ($c = 0; $c -lt $cCount; $c++) {
                $ref = "$([char]([int][char]'A' + $c))$rowNum"
                $fmt = $s.Fmts[$c]
                $typeOffset = if ($fmt -eq '#,##0') { 2 } elseif ($fmt -eq '0.0%') { 3 } elseif ($s.Center -contains ($c + 1)) { 1 } else { 0 }
                $fontGroup = if (-not $isWarn) { 0 } elseif ($c -eq ($cCount - 1)) { 2 } else { 1 }
                $styleIdx = 3 + ($fontGroup * 8) + ($zebra * 4) + $typeOffset
                if ($fmt -eq '@') {
                    $null = $sb.Append("<c r=`"$ref`" s=`"$styleIdx`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(&$EscXml ([string]$rowObj[$c]))</t></is></c>")
                } else {
                    $numVal = ([double]$rowObj[$c]).ToString([Globalization.CultureInfo]::InvariantCulture)
                    $null = $sb.Append("<c r=`"$ref`" s=`"$styleIdx`"><v>$numVal</v></c>")
                }
            }
            $null = $sb.Append('</row>')
        }
        $endRow = [Math]::Max(3, 3 + $rCount)
        $null = $sb.Append("</sheetData><autoFilter ref=`"A3:${lastCol}${endRow}`"/></worksheet>")
        &$AddXml "xl/worksheets/sheet$($i + 1).xml" $sb.ToString()
    }
} finally {
    $zip.Dispose()
}
Write-Output "出力完了: $outPath"
