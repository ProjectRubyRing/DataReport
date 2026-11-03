<#
.SYNOPSIS
  A5:SQL Mk-2 で一括エクスポートした CSV から、データ件数・項目別データ分布・キー先頭1桁集計・全データ検索の Excel レポートを作成する。
.EXAMPLE
  .\New-DataReport.ps1 -CsvDir C:\export\mydb
  .\New-DataReport.ps1 -CsvDir C:\export\mydb -OutFile C:\report\mydb.xlsx -KeyColumn @{ orders = 'order_no' }
.NOTES
  前提: 1 行目が列名 / ファイル名 = テーブル名 / 空値 = NULL。文字コード(UTF-8・Shift_JIS)は自動判定。Excel(Microsoft 365)が必要。
#>
param(
    [Parameter(Mandatory)][string]$CsvDir,
    [string]$OutFile = (Join-Path $CsvDir "DataReport_$(Get-Date -f yyyyMMdd_HHmmss).xlsx"),
    [hashtable]$KeyColumn = @{}   # キー項目の指定 @{ テーブル名 = '列名' }（未指定は先頭列）
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
$OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
[void](New-Item -ItemType Directory -Force (Split-Path $OutFile))
$NAVY, $BURG, $GRAY, $LIGHT = 0x4C2B1A, 0x200080, 0x333333, 0xF2F2F2   # Excel の色は BGR 順（#1A2B4C → 0x4C2B1A）
$LIMIT = 1048572                                                     # 全データの格納上限（Excel 最大行 - 見出し 4 行）
$sum, $dist, $pre, $all = 1..4 | ForEach-Object { , [Collections.Generic.List[object]]::new() }
$total = 0
function Cut([string]$s, $n = 40) { if ($s.Length -gt $n) { $s.Substring(0, $n) + '…' } else { $s } }   # 長い値は省略表示

foreach ($f in Get-ChildItem -LiteralPath $CsvDir -Filter *.csv | Sort-Object Name) {
    # 文字コード判定（UTF-8 として不正なら Shift_JIS）→ CSV 解析（"" 囲み内のカンマ・改行に対応）
    $enc = try { [void][Text.UTF8Encoding]::new($false, $true).GetString([IO.File]::ReadAllBytes($f.FullName)); [Text.Encoding]::UTF8 } catch { [Text.Encoding]::GetEncoding(932) }
    $p = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($f.FullName, $enc); $p.SetDelimiters(','); $p.TrimWhiteSpace = $false
    $t = [string]$f.BaseName; $hdr = $p.ReadFields(); $rows = [Collections.Generic.List[string[]]]::new()
    while (!$p.EndOfData) { $rows.Add($p.ReadFields()) }
    $p.Close()
    if (!$hdr) { Write-Warning "空ファイルのためスキップ: $($f.Name)"; continue }
    $n = $rows.Count; $total += $n; $ki = [Math]::Max(0, [Array]::IndexOf($hdr, $KeyColumn[$t]))
    Write-Host ('{0,-40}{1,12:N0} 件' -f $t, $n)

    # 項目別データ分布
    for ($c = 0; $c -lt $hdr.Count; $c++) {
        $cnt = [Collections.Generic.Dictionary[string, int]]::new(); $nul = 0; $len = 0
        foreach ($r in $rows) { $v = [string]$r[$c]; if ($v -eq '') { $nul++ } else { $cnt[$v]++; if ($v.Length -gt $len) { $len = $v.Length } } }
        $k = [string[]]$cnt.Keys; $avg = $null
        $type = if (!$k) { '-' } elseif (!($k -notmatch '^-?(0|[1-9]\d*)(\.\d+)?$')) { '数値' } elseif (!($k -notmatch '^\d{4}[-/]\d\d?[-/]\d\d?')) { '日時' } else { '文字列' }
        if ($type -eq '数値') { [Array]::Sort([double[]]$k, $k); $avg = 0; foreach ($e in $cnt.GetEnumerator()) { $avg += [double]$e.Key * $e.Value / ($n - $nul) } }
        else { [Array]::Sort($k, [StringComparer]::Ordinal) }
        $tk = [string[]]$cnt.Keys; $tv = [int[]]$cnt.Values; [Array]::Sort($tv, $tk)   # 出現回数の昇順
        $top = if (!$k) { $null } elseif ($k.Count -eq $n - $nul) { '(全件ユニーク)' } else { (($tk.Count - 1)..[Math]::Max(0, $tk.Count - 3) | ForEach-Object { "$(Cut $tk[$_]) ($($tv[$_]))" }) -join ', ' }
        $w = @(if ($n -and $nul -eq $n) { '全件NULL' }; if ($n -gt 1 -and $k.Count -eq 1) { '単一値' }
               if ($c -eq $ki -and $nul) { 'キー空白' }; if ($c -eq $ki -and $k.Count -lt $n - $nul) { 'キー重複' })
        if ($c -eq $ki) { $kw = $w; $ku = $k.Count }
        $dist.Add(@($t, $hdr[$c], $type, $nul, $(if ($n) { $nul / $n } else { $null }), $k.Count, (Cut $k[0]), (Cut $k[-1]), $avg, $len, $top, ($w -join '／')))
    }
    $sum.Add(@($t, $n, $hdr.Count, $hdr[$ki], $ku, ((@(if (!$n) { 'データなし' }) + $kw) -join '／')))

    # キー項目の先頭1桁カテゴリ別件数
    $pc = [Collections.Generic.Dictionary[string, int]]::new()
    foreach ($r in $rows) { $v = [string]$r[$ki]; $pc[$(if ($v) { $v.Substring(0, 1) } else { '(空白)' })]++ }
    $pk = [string[]]$pc.Keys; [Array]::Sort($pk, [StringComparer]::Ordinal)
    foreach ($x in $pk) { $pre.Add(@($t, $hdr[$ki], $x, $pc[$x], ($pc[$x] / $n))) }

    # 全データ（検索用）: 1 レコード =「列名=値 | 列名=値 …」
    $fmt = (0..($hdr.Count - 1) | ForEach-Object { ($hdr[$_] -replace '[{}]', '$0$0') + "={$_}" }) -join ' | '
    for ($i = 0; $i -lt $n -and $all.Count -lt $LIMIT; $i++) { $all.Add(@($t, ($i + 1), (Cut (($fmt -f $rows[$i]) -replace '[\r\n]+', ' ') 32000))) }
}
if (!$sum.Count) { throw "CSV ファイルがありません: $CsvDir" }

$xl = New-Object -ComObject Excel.Application
try {
    $xl.ScreenUpdating = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add(-4167)
    $fo = $wb.Styles.Item('Normal').Font; $fo.Name = 'Meiryo UI'; $fo.Size = 10; $fo.Color = $GRAY   # 標準スタイル = 全セル既定の書式
    $ts = $wb.TableStyles.Add('Mono'); $ts.TableStyleElements.Item(0).Font.Color = $GRAY; $ts.TableStyleElements.Item(5).Interior.Color = $LIGHT   # 表: 文字 #333333・縞 #F2F2F2

    # 注意: Value2 へは [string] 等の素の値か [object[,]] だけを代入する。PSObject（型なし引数など）を一度でも代入すると、
    #       PowerShell 5.1 の COM バインダーがその変換規則を使い回し、以降の 2 次元配列の一括代入が失敗するため
    function Set-Head($ws, $head) {   # 4 行目の見出し（ミッドナイトネイビー地・白太字）
        for ($c = 0; $c -lt $head.Count; $c++) { $ws.Cells.Item(4, $c + 1).Value2 = [string]$head[$c] }
        $r = $ws.Range('A4').Resize(1, $head.Count); $r.Interior.Color = $NAVY; $r.Font.Color = 0xFFFFFF; $r.Font.Bold = $true; $r.HorizontalAlignment = -4108; $r.RowHeight = 22
    }
    function New-Sheet([string]$name, [string]$title, [string]$note) {   # 1 行目タイトル / 2 行目説明 / 4 行目見出し（ここまで固定表示）
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $ws.Name = $name; $ws.Tab.Color = $NAVY
        $ws.Range('A1').Value2 = $title; $fo = $ws.Range('A1').Font; $fo.Size = 14; $fo.Bold = $true; $fo.Color = $NAVY
        $ws.Range('A2').Value2 = $note; $ws.Range('A2').Font.Size = 9
        $xl.ActiveWindow.DisplayGridlines = $false; $xl.ActiveWindow.SplitRow = 4; $xl.ActiveWindow.FreezePanes = $true
        $ws
    }
    function Add-Table($ws, $head, $rows, [string]$name) {   # 4 行目から表を出力（フィルター・縞模様付きテーブル）
        $m = $head.Count; $nr = [Math]::Max(1, $rows.Count)
        for ($c = 0; $c -lt $m; $c++) {   # 列の書式は書込み前に設定（文字列の数値・日付・数式への自動変換を防止）
            $ws.Columns.Item($c + 1).NumberFormat = switch -Regex ($head[$c]) { '率$|比$' { '0.0%' } '数$|長$|No$' { '#,##0' } '^平均' { '#,##0.00' } default { '@' } }
            $cell = $ws.Range('A5').Offset(0, $c).Resize($nr, 1)
            if ($head[$c] -eq '注意') { $cell.Font.Color = $BURG; $cell.Font.Bold = $true }
            if ($head[$c] -match '率$|比$') { $db = $cell.FormatConditions.AddDatabar(); $db.BarColor.Color = $NAVY; $db.MinPoint.Modify(0, 0); $db.MaxPoint.Modify(0, 1) }
        }
        Set-Head $ws $head
        for ($i = 0; $i -lt $rows.Count; $i += 10000) {   # 1 万行ずつ 2 次元配列で一括書込み
            $k = [Math]::Min(10000, $rows.Count - $i); $a = [object[,]]::new($k, $m)
            for ($y = 0; $y -lt $k; $y++) { for ($x = 0; $x -lt $m; $x++) { $a[$y, $x] = $rows[$i + $y][$x] } }
            $ws.Range('A5').Offset($i, 0).Resize($k, $m).Value2 = $a
        }
        $lo = $ws.ListObjects.Add(1, $ws.Range('A4').Resize($nr + 1, $m), $null, 1); $lo.TableStyle = 'Mono'; if ($name) { $lo.Name = $name }
        for ($c = 0; $c -lt $m; $c++) {   # 数値に見える文字列（最小値・最大値・先頭1桁）の緑三角を非表示（範囲一括不可・テーブル化後に設定）
            if ($head[$c] -match '^最.値$|桁$') { for ($y = 0; $y -lt $rows.Count; $y++) { if ($rows[$y][$c] -match '^[-+]?[\d.]') { $ws.Cells.Item($y + 5, $c + 1).Errors.Item(3).Ignore = $true } } }
        }
        [void]$ws.Range('A4').Resize([Math]::Min($nr + 1, 1000), $m).Columns.AutoFit()   # 列幅は先頭 1000 行で自動調整（上限 100）
        foreach ($col in $lo.Range.Columns) { $col.ColumnWidth = [Math]::Min(100, $col.ColumnWidth + 2) }
    }

    $ws = New-Sheet 'サマリー' 'データ件数サマリー' "対象: $CsvDir ／ テーブル数: $($sum.Count) ／ 総件数: $('{0:N0}' -f $total) 件 ／ 作成: $(Get-Date -f 'yyyy/MM/dd HH:mm')"
    Add-Table $ws ('テーブル', '件数', '項目数', 'キー項目', 'キーユニーク数', '注意') $sum
    $ws = New-Sheet 'データ分布' '項目別データ分布' '空文字は NULL として集計。推定型は値の形式で判定し、最小値・最大値は数値なら数値順、それ以外は文字コード順'
    Add-Table $ws ('テーブル', '項目名', '推定型', 'NULL件数', 'NULL率', 'ユニーク数', '最小値', '最大値', '平均', '最大長', '上位値(件数)', '注意') $dist
    $ws = New-Sheet 'キー先頭1桁' 'キー項目 先頭1桁カテゴリ別件数' 'キー項目は各テーブルの先頭列（-KeyColumn で変更可）'
    Add-Table $ws ('テーブル', 'キー項目', '先頭1桁', '件数', '構成比') $pre
    $wsS = New-Sheet '全データ検索' '全データ検索' 'C3 にキーワードを入力すると、全テーブル・全項目を部分一致で検索します（大文字小文字は区別なし／ワイルドカード * ? 可）'
    $ws = New-Sheet '全データ' '全データ（検索用）' $(if ($all.Count -lt $total) { "※ Excel の行数上限により先頭 $('{0:N0}' -f $LIMIT) 件のみ格納しています" } else { '1 レコード ＝「列名=値 | 列名=値 …」' })
    Add-Table $ws ('テーブル', '行No', '内容') $all 'T_ALL'
    if ($all.Count -lt $total) { $ws.Range('A2').Font.Color = $BURG }

    # 検索シート: キーワード → FILTER 関数で全データから抽出（動的配列: Microsoft 365 / Excel 2021 以降）
    $ws = $wsS; $ws.Range('B3').Value2 = '検索語'; $ws.Range('A3:B3').Font.Bold = $true
    $in = $ws.Range('C3'); $in.NumberFormat = '@'; $in.Interior.Color = $LIGHT; $in.Borders.Color = $NAVY
    $ws.Range('A3').Formula2 = '=IF(C3="","",COUNTIF(T_ALL[内容],"*"&C3&"*")&" 件")'
    Set-Head $ws ('テーブル', '行No', '内容')
    $ws.Range('A5').Formula2 = '=IF(C3="","",FILTER(T_ALL,ISNUMBER(SEARCH(C3,T_ALL[内容])),"該当なし"))'
    $ws.Range('A5:C1048576').FormatConditions.Add(2, [Type]::Missing, '=AND(INDEX($A:$A,ROW())<>"",MOD(ROW(),2)=1)').Interior.Color = $LIGHT
    $ws.Range('B:B').NumberFormat = '#,##0'; $ws.Range('A:A').ColumnWidth = 28; $ws.Range('B:B').ColumnWidth = 9; $ws.Range('C:C').ColumnWidth = 150

    [void]$wb.Worksheets.Item(1).Delete(); $wb.Worksheets.Item(1).Activate()   # 初期シートを削除しサマリーを表示
    $wb.SaveAs($OutFile, 51); $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
Write-Host "出力完了: $OutFile"
