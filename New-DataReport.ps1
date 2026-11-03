<#
.SYNOPSIS
  A5:SQL Mk-2 で一括エクスポートした CSV を解析し、グラフ付きのデータプロファイル Excel レポートを作成する。
.DESCRIPTION
  データ件数・カーディナリティ・値の分布をグラフで可視化し、CSV のデータから推定したインデックスの提案と、
  検索で問題になりやすいデータ分布（値の偏り・低カーディナリティ・表記ゆれ・型混在 等）の警告を出力する。

  出力シート
    ダッシュボード     KPI ／件数の多い・少ないテーブル／規模別テーブル数／カーディナリティ区分／カーディナリティマップ
                       ／検索リスクとインデックス提案の内訳／対応を優先すべき警告
    テーブル分析       テーブルと項目を選ぶと、項目別カーディナリティ・上位値の構成比・値の分布のグラフと項目一覧が切り替わる
    テーブル一覧       テーブル別の件数・主キー候補・提案数・警告数
    データ分布         項目別の推定型・NULL 率・ユニーク数・カーディナリティ区分・最頻値 等
    インデックス提案   優先度・根拠・想定する検索・DDL
    検索リスク警告     重要度・内容・検索への影響・対策
    値の分布           項目別の上位値と構成比
    キー先頭1桁 ／ 全データ検索 ／ 全データ ／ 判定基準
.EXAMPLE
  .\New-DataReport.ps1 -CsvDir C:\export\mydb
  .\New-DataReport.ps1 -CsvDir C:\export\mydb -OutFile C:\report\mydb.xlsx -KeyColumn @{ orders = 'order_no' }
.NOTES
  前提: 1 行目が列名 / ファイル名 = テーブル名 / 空値 = NULL。文字コード（UTF-8・Shift_JIS）は自動判定。Excel（Microsoft 365）が必要。
  インデックス提案は CSV のデータだけから推定したもの。既存のインデックスや実際の SQL は考慮しないため、作成前に EXPLAIN 等で確認すること。
  照合順序に関する判定は Aurora MySQL 8.x の既定（utf8mb4_0900_ai_ci）を想定している。
  このファイルは必ず「UTF-8 (BOM 付き)」で保存すること。BOM なしで保存すると PowerShell 5.1 が Shift_JIS として読み込み、
  日本語が文字化けして「終わりの '}' が存在しません」「トークン ')' を使用できません」等の構文エラーで起動できなくなる。
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
$Started = Get-Date

#region ── 設定 ─────────────────────────────────────────────────────────────

function ConvertTo-Bgr([string]$Hex) {   # '#RRGGBB' → Excel の色値（BGR 順の整数）
    $v = [Convert]::ToInt32($Hex.TrimStart('#'), 16)
    (($v -band 0xFF) -shl 16) -bor ($v -band 0xFF00) -bor (($v -shr 16) -band 0xFF)
}
# 配色: ミッドナイトネイビー（基調・順序のある段階は同系色の濃淡）／バーガンディ（警告）／グレー（文字・罫線）
$Color = @{}
foreach ($e in ([ordered]@{
            Navy = '#1A2B4C'; Navy2 = '#45577A'; Navy3 = '#7584A0'; Navy4 = '#A3AEC2'; NavyTint = '#E8EBF1'
            Burg = '#800020'; Burg2 = '#A64D61'; Burg3 = '#CC8C9A'; BurgTint = '#F6E9EC'
            Ink = '#333333'; Ink2 = '#666666'; Muted = '#999999'; Gray = '#BFBFBF'; Rule = '#D9D9D9'; Hair = '#E7E7E7'
            Chip = '#EBEBEB'; Stripe = '#F2F2F2'; Canvas = '#F3F4F6'; White = '#FFFFFF'
        }).GetEnumerator()) { $Color[$e.Key] = ConvertTo-Bgr $e.Value }

# 判定のしきい値
$TH = @{
    LowCard  = 10        # ユニーク数がこれ未満 → 低カーディナリティ（1 値あたり平均で全体の 10% 超）
    HighCard = 100       # ユニーク数がこれ以上 → 高
    SkewWarn = 0.5       # 最頻値が非 NULL 件数のこの割合以上 → 値の偏り（警告）
    SkewNote = 0.2       # この割合以上かつ 1 値あたり平均件数の SkewTimes 倍以上 → 値の偏り（注意）
    SkewTimes = 20
    NullWarn = 0.5       # NULL 率
    MinRows  = 1000      # 検索性能の評価対象とする件数（未満のテーブルは性能系の重要度を 1 段階下げる）
    BigRows  = 100000    # 大規模テーブル
    IndexBytes = 3072    # InnoDB のインデックスキー長の上限（バイト）
    TopN     = 10        # 値の分布に出す上位値の数
    HistBins = 20        # 数値の分布の区間数（目安）
}
$LIMIT = 1048572          # 全データの格納上限（Excel 最大行 - 見出し 4 行）
$SEP = [string][char]1    # 値の連結用の区切り文字（データ中に現れない制御文字）
$Utf8 = [Text.UTF8Encoding]::new($false)
$Inv = [Globalization.CultureInfo]::InvariantCulture

# 値の形式・列名のパターン（数字は [0-9] で判定する。\d は全角数字等にも一致し、数値変換で停止するため）
$Re = @{
    Num          = '^-?(0|[1-9][0-9]*)(\.[0-9]+)?$'
    NumLoose     = '^-?[0-9]+(\.[0-9]+)?$'
    Date         = '^[0-9]{4}[-/][0-9]{1,2}[-/][0-9]{1,2}'
    Time         = '[ T][0-9]{1,2}:[0-9]{2}'
    PseudoNull   = '^(?:null|\(null\)|<null>|n/a|#n/a|none|nil|undefined|-|－|―|なし|無し|不明|未設定|未入力|\s+)$'
    SentinelDate = '^(?:(?:9999|29[0-9]{2}|3000|0000|0001)[-/]|1900[-/]0?1[-/]0?1|1899[-/]12[-/]3[01]|1753[-/]0?1[-/]0?1)'
    SentinelNum  = '^-?9{4,}(\.9+)?$'
    Space        = '^\s|\s$'
    Ctrl         = '[\r\n\t]'
    Surrogate    = '[\uD800-\uDBFF]'
    IdName       = '(^|_)(id|cd|code|no|num|number|key|kbn)$|(?-i:[a-z](Id|Cd|Code|No|Num|Key)$|ID$)|(番号|コード|ＩＤ)$'
    UniqName     = '(^|_)(id|cd|code|no|num|number|key|uuid|guid|mail|email|account|login)$|(?-i:[a-z](Id|Cd|Code|No|Mail|Email)$|ID$)|(番号|コード|ＩＤ|メール)$'
    CatName      = '(^|_)(status|state|type|kind|kbn|kubun|flag|flg|class|cls|category|level|lvl|rank|div)(_|$)|(?-i:[a-z](Status|Type|Kind|Flag|Class|Category|Level|Rank)$)|区分|種別|状態|フラグ|ステータス|分類|ランク'
    DateName     = '(date|day|_at$|_on$|dt$|time|日|時)'
    DateSkip     = '(^|_)(upd|update|updated|modified|changed|deleted)|(?-i:[a-z](Upd|Update|Modified))|更新|削除'
    SearchName   = '(^|_)(name|kana|mail|email|tel|phone|fax|zip|postal)(_|$)|(?-i:[a-z](Name|Kana|Mail|Email|Tel|Phone)$)|氏名|名前|名称|カナ|メール|電話|郵便'
    TextName     = '(^|_)(title|note|notes|memo|comment|comments|desc|description|body|text|content|contents|remark|remarks|message|msg|address|addr)(_|$)|(?-i:[a-z](Title|Note|Memo|Comment|Description|Body|Text|Content|Remarks?|Message|Address)$)|件名|本文|内容|説明|備考|メモ|コメント|住所|メッセージ|タイトル'
}

# 警告の分類（重要度・判定条件・検索への影響・対策）。Demote = 小規模テーブルでは重要度を 1 段階下げる（性能面の問題）
$Rule = [ordered]@{
    'キー重複'          = @{ Sev = '重大'; Demote = $false; Cond = 'キー項目に同じ値が複数ある'; Impact = '主キー・一意制約を設定できない。キーで 1 件を特定する検索が複数件を返す'; Fix = '重複データを調査・解消する。業務上重複するなら複合キーかサロゲートキーを採用する' }
    'キー空白'          = @{ Sev = '重大'; Demote = $false; Cond = 'キー項目に NULL（空値）がある'; Impact = '主キーに設定できない。NULL は = や IN の検索で一致しない'; Fix = '値を補完して NOT NULL 制約を付ける' }
    '主キー候補なし'    = @{ Sev = '警告'; Demote = $false; Cond = '一意かつ NULL なしの列（2 列の組合せを含む）がない'; Impact = 'InnoDB は内部の行 ID でデータを管理し、行の特定・更新・レプリケーションの性能が低下する'; Fix = 'サロゲートキー（BIGINT AUTO_INCREMENT）を追加する' }
    'ほぼユニーク'      = @{ Sev = '警告'; Demote = $false; Cond = 'ID・コード・メール等の列が 99% 以上一意だが重複がある'; Impact = '一意制約を設定できず、ID の検索で複数件が返る'; Fix = '重複データの混入を調査する' }
    '参照整合性'        = @{ Sev = '警告'; Demote = $false; Cond = '結合候補の列に、参照先テーブルのキーに存在しない値がある'; Impact = 'INNER JOIN で行が欠落する。外部キー制約を設定できない'; Fix = '参照先にない値（孤立データ）を調査・修正する' }
    '値の偏り'          = @{ Sev = '警告'; Demote = $true; Cond = '最頻値が非 NULL 件数の 50% 以上（注意: 20% 以上かつ 1 値あたり平均件数の 20 倍以上）'; Impact = '最頻値での検索はインデックスを使わず全件走査になる。インデックス統計は 1 値あたりの平均件数で見積もるため、結合時の行数見積りを誤り実行計画が悪化しやすい'; Fix = 'ヒストグラムを作成する（ANALYZE TABLE … UPDATE HISTOGRAM）。最頻値での検索を避け、複合インデックスで絞り込む' }
    '低カーディナリティ' = @{ Sev = '情報'; Demote = $false; Cond = 'ユニーク数 2～9（非 NULL 1,000 件以上）'; Impact = '1 値あたり平均で全体の 10% 超が該当し、単独のインデックスはほぼ使われない'; Fix = '単独のインデックスは作らず、複合インデックスの先頭列として使う' }
    '表記ゆれ'          = @{ Sev = '警告'; Demote = $false; Cond = '全角/半角・大文字/小文字・ひらがな/カタカナの違いを除くと同じになる値がある'; Impact = '照合順序が _bin・_as_cs では検索漏れ、_ai_ci では想定外の一致になる。集計や一意制約も不整合になる'; Fix = '登録時に表記を正規化し、既存データをクレンジングする' }
    '照合順序で同一視'  = @{ Sev = '注意'; Demote = $false; Cond = 'utf8mb4_0900_ai_ci で同じと判定される別の値がある（濁点・半濁点・アクセントの違い。例: ハハ／パパ）'; Impact = '= の検索で別の値も一致する。UNIQUE 制約で重複エラーになる'; Fix = '列の照合順序を utf8mb4_ja_0900_as_cs 等に変更する' }
    '前後空白'          = @{ Sev = '警告'; Demote = $false; Cond = '先頭・末尾に空白（全角を含む）がある値'; Impact = 'utf8mb4_0900 系の照合順序（NO PAD）では「abc」と「abc␣」が一致せず検索漏れになる'; Fix = 'TRIM で除去し、登録時にも除去する' }
    '先頭ゼロ'          = @{ Sev = '警告'; Demote = $false; Cond = '数字のみで、先頭が 0 の値がある（例: 00123）'; Impact = '数値（WHERE col = 123）で検索すると暗黙の型変換でインデックスが使われず全件走査になり、「123」と「00123」も一致する'; Fix = '文字列（WHERE col = ''00123''）で検索する。桁数・型を統一する' }
    '型混在'            = @{ Sev = '警告'; Demote = $false; Cond = '数値・日付の形式が大半だが、一部に異なる形式の値がある'; Impact = '数値・日付として比較すると変換エラーや暗黙の型変換でインデックスが使われない。範囲検索が文字列順になる'; Fix = '不正な値を修正し、適切なデータ型（INT・DATE 等）に変更する' }
    '疑似NULL'          = @{ Sev = '注意'; Demote = $false; Cond = '「NULL」「N/A」「-」空白のみ 等、NULL の代わりの値がある'; Impact = 'IS NULL で検索できず、件数や選択度の見積りが狂う'; Fix = '本来の NULL に置き換える' }
    '番兵値'            = @{ Sev = '注意'; Demote = $false; Cond = '9999-12-31・1900-01-01 等のダミー日付、9999… の数値がある'; Impact = '範囲検索・並べ替え・統計（ヒストグラム）が歪み、期間検索の件数が想定とずれる'; Fix = 'NULL を使うか、検索条件で除外する' }
    'NULL多数'          = @{ Sev = '注意'; Demote = $true; Cond = 'NULL 率 50% 以上'; Impact = 'NULL は = や IN で一致せず検索漏れの原因になる。IS NULL 検索の選択度が低い'; Fix = '既定値の設定や列の要否を見直す' }
    '改行・制御文字'    = @{ Sev = '注意'; Demote = $false; Cond = '改行・タブを含む値がある'; Impact = '完全一致の検索や CSV 連携で一致しない'; Fix = '用途を確認し、不要なら除去する' }
    '長大値'            = @{ Sev = '注意'; Demote = $false; Cond = '最大バイト数が 3,072 を超える（utf8mb4 換算）'; Impact = '列全体の B-tree インデックスは作成できない（InnoDB の上限 3,072 バイト）'; Fix = 'プレフィックスインデックス col(n) か FULLTEXT インデックスを使う' }
    '4バイト文字'       = @{ Sev = '情報'; Demote = $false; Cond = '絵文字・𠮷 等の 4 バイト文字を含む'; Impact = 'utf8mb3（utf8）の列には格納できない'; Fix = '列の文字セットを utf8mb4 にする' }
    '単一値'            = @{ Sev = '情報'; Demote = $false; Cond = '全件が同じ値'; Impact = '検索条件として機能しない（常に全件が一致）'; Fix = '列の要否を見直す' }
    '全NULL'            = @{ Sev = '情報'; Demote = $false; Cond = '全件 NULL'; Impact = '検索条件として機能しない'; Fix = '未使用の列なら削除を検討する' }
    'データなし'        = @{ Sev = '情報'; Demote = $false; Cond = '0 件'; Impact = '件数・分布を評価できない'; Fix = 'エクスポートの条件を確認する' }
}
# インデックス提案の種別（判定条件・ねらい・データ分布シートでの表示名）
$IdxKind = [ordered]@{
    '主キー'             = @{ Tag = '主キー'; Cond = '一意かつ NULL なしの列（なければ 2 列の組合せ）'; Aim = '行の特定・結合の基点。InnoDB は主キー順にデータを格納する' }
    '主キー（サロゲート）' = @{ Tag = '主キー'; Cond = '主キーにできる列・組合せがない'; Aim = '行を一意に特定できるようにする' }
    '一意インデックス'   = @{ Tag = '一意'; Cond = 'ID・コード・メール等の名前で、全件一意な列'; Aim = '重複の混入を防ぎ、等価検索を高速化する' }
    '結合・外部キー'     = @{ Tag = '索引(結合)'; Cond = '他テーブルのキーと列名が対応し、値の 80% 以上が一致する列'; Aim = '結合と外部キー制約の検査を高速化する' }
    '複合インデックス'   = @{ Tag = '複合索引'; Cond = '区分系の列（ユニーク数 2～30）と日付列があるテーブル（1,000 件以上）'; Aim = '「区分で絞り込み＋期間指定」の検索。等価条件の列を先頭、範囲条件の列を後ろに置く' }
    '期間検索'           = @{ Tag = '索引(期間)'; Cond = 'ユニーク数 10 以上・NULL 率 50% 未満の日付列（1,000 件以上）'; Aim = '期間指定（BETWEEN）・日付順の並べ替え' }
    '少数値検索'         = @{ Tag = '索引(少数値)'; Cond = 'ユニーク数 2～9 で、全体の 5% 以下の値がある列（1 万件以上）'; Aim = '少数派の値（エラー・未処理 等）の抽出に限り有効' }
    '検索キー'           = @{ Tag = '索引(検索)'; Cond = '名称・カナ・連絡先等の列、または参照先を特定できない ID・コード系の列で、ユニーク数 100 以上（1,000 件以上）'; Aim = '完全一致・前方一致（LIKE ''abc%''）の検索' }
    '全文検索'           = @{ Tag = '全文索引'; Cond = '文章系の列（平均 10 文字以上の備考・本文等、または平均 50 文字以上）'; Aim = '部分一致（LIKE ''%語%''）は B-tree を使えないため、ngram の全文検索で代替する' }
    'ヒストグラム'       = @{ Tag = 'ヒストグラム'; Cond = '値の偏りの警告がある列（1,000 件以上）'; Aim = 'インデックスのない条件でも行数の見積りを正確にし、結合順序の誤りを防ぐ' }
}
$SevOrder = @{ '重大' = 1; '警告' = 2; '注意' = 3; '情報' = 4 }
$PriOrder = @{ '高' = 1; '中' = 2; '低' = 3 }
$Classes = 'ユニーク', '高', '中', '低', '単一値', '全NULL'

#endregion
#region ── 共通関数 ─────────────────────────────────────────────────────────

function Cut([string]$s, [int]$n = 40) {   # 長い値は省略表示（𠮷 等のサロゲートペアの途中では切らない）
    if ($s.Length -gt $n) { $s.Substring(0, $n - [char]::IsHighSurrogate($s[$n - 1])) + '…' } else { $s }
}
function Limit-Width([string]$s, [int]$units) {   # 表示幅（全角 = 2・半角 = 1）で切り詰める
    $w = 0
    for ($i = 0; $i -lt $s.Length; $i++) {
        $ch = [int]$s[$i]; $w += $(if ($ch -lt 0x100 -or ($ch -ge 0xFF61 -and $ch -le 0xFF9F)) { 1 } else { 2 })
        if ($w -gt $units) { $cut = $i - 1; if ($cut -gt 0 -and [char]::IsHighSurrogate($s[$cut - 1])) { $cut-- }; return $s.Substring(0, [Math]::Max(0, $cut)) + '…' }
    }
    $s
}
function Format-Cell([string]$s, [int]$n = 40) { Cut ($s -replace '
?
', '↵' -replace '	', '→') $n }   # セル表示用（改行で行が崩れないように）
function Format-N($x) { '{0:N0}' -f $x }
function Format-Pct($x) { '{0:0.0}%' -f (100 * $x) }
function Format-Val([string]$s, [int]$n = 24) {   # 例示用: 空白・改行を記号で見せ、「」で囲む
    '「' + (Cut ($s -replace ' ', '␣' -replace '　', '□' -replace '\r?\n', '↵' -replace '\t', '→') $n) + '」'
}
function Format-Id([string]$id) { '`' + $id.Replace('`', '``') + '`' }   # MySQL の識別子
function Get-IndexName([string]$prefix, [string]$table, [string[]]$cols) {
    $s = ($prefix + '_' + $table + '_' + ($cols -join '_')) -replace '[^\w]', '_'
    if ($s.Length -gt 64) { $s.Substring(0, 64) } else { $s }
}
function ConvertTo-CellValue([string]$s) {   # 数値として損失なく表せる値は数値で格納（「文字列として保存された数値」の緑三角を出さない）
    if ($s -match '^-?(0|[1-9][0-9]{0,14})(\.[0-9]{1,6})?$') {
        $d = [double]$s; if ($d.ToString('R', $Inv) -eq $s) { return $d }
    }
    $s
}
function Get-RowSum($dict, [string[]]$keys) { $s = 0; foreach ($x in $keys) { $s += $dict[$x] }; $s }
function Get-Examples($dict, [string[]]$keys, [int]$take = 3) {   # 件数の多い順に例示
    $arr = if ($keys.Count -gt 2000) { $keys[0..1999] } else { $keys }
    ($arr | Sort-Object { - $dict[$_] } | Select-Object -First $take | ForEach-Object { (Format-Val $_) + '×' + (Format-N $dict[$_]) }) -join '、'
}
function Get-VariantGroups([string[]]$norm) {   # 正規化後に同じになる値のグループ（要素の添字の配列）を返す
    $g = [Collections.Generic.Dictionary[string, Collections.Generic.List[int]]]::new()
    for ($i = 0; $i -lt $norm.Count; $i++) {
        $l = $null
        if (!$g.TryGetValue($norm[$i], [ref]$l)) { $l = [Collections.Generic.List[int]]::new(); $g[$norm[$i]] = $l }
        $l.Add($i)
    }
    foreach ($l in $g.Values) { if ($l.Count -gt 1) { , $l.ToArray() } }
}
function Format-Groups($groups, [string[]]$raw, $dict, [int]$take = 2) {   # 例: 「ABC商事」/「ＡＢＣ商事」
    $sorted = $groups | Select-Object -First 500 | Sort-Object { - (Get-RowSum $dict ([string[]]($_ | ForEach-Object { $raw[$_] }))) } | Select-Object -First $take
    ($sorted | ForEach-Object { (($_ | Sort-Object { - $dict[$raw[$_]] } | Select-Object -First 3 | ForEach-Object { Format-Val $raw[$_] 16 }) -join '/') }) -join '、'
}
function Get-Step([double]$range, [int]$bins) {   # 区間幅を 1・2・2.5・5 × 10^n に丸める
    if ($range -le 0) { return 1 }
    $raw = $range / $bins; $mag = [Math]::Pow(10, [Math]::Floor([Math]::Log10($raw)))
    foreach ($m in 1, 2, 2.5, 5, 10) { if ($raw -le $m * $mag) { return $m * $mag } }
}

# 数値の分布: 1～99 パーセンタイルの範囲を等間隔の区間に分け、範囲外は「<」「≥」の区間にまとめる
function Get-NumericHistogram([double[]]$x, [int[]]$w, [double]$lo, [double]$hi, [bool]$isInt) {
    if ($isInt -and $hi - $lo -le 40) { $step = 1.0; $start = $lo }
    else {
        $step = Get-Step ($hi - $lo) $TH.HistBins; if ($isInt) { $step = [Math]::Max(1, [Math]::Ceiling($step)) }
        $start = [Math]::Floor($lo / $step) * $step
    }
    $nb = [int][Math]::Floor(($hi - $start) / $step) + 1
    $bins = [double[]]::new($nb); $under = 0; $over = 0
    for ($i = 0; $i -lt $x.Count; $i++) {
        if ($x[$i] -lt $start) { $under += $w[$i]; continue }
        $b = [int][Math]::Floor(($x[$i] - $start) / $step)
        if ($b -ge $nb) { $over += $w[$i] } else { $bins[$b] += $w[$i] }
    }
    $dec = if ($step -ge 1) { 0 } else { [int][Math]::Ceiling(-[Math]::Log10($step)) }
    $f = "{0:N$dec}"
    $out = [Collections.Generic.List[object]]::new()
    if ($under) { $out.Add(@(('<' + ($f -f $start)), $under)) }
    for ($b = 0; $b -lt $nb; $b++) { $v = $start + $b * $step; $out.Add(@($(if ($step -eq 1 -and $isInt) { $f -f $v } else { ($f -f $v) + '～' }), $bins[$b])) }
    if ($over) { $out.Add(@(('≥' + ($f -f ($start + $nb * $step))), $over)) }
    @{ Unit = '数値の区間'; Rows = $out }
}
# 日付の分布: 期間の長さに応じて日別・月別・年別に集計し、欠けた期間も 0 件で表示する
function Get-DateHistogram([string[]]$d, [int[]]$w, [string]$lo, [string]$hi) {
    $months = ([int]$hi.Substring(0, 4) * 12 + [int]$hi.Substring(5, 2)) - ([int]$lo.Substring(0, 4) * 12 + [int]$lo.Substring(5, 2))
    if ($months -le 2) { $len = 10; $unit = '日別' } elseif ($months -le 60) { $len = 7; $unit = '月別' } else { $len = 4; $unit = '年別' }
    $kLo = $lo.Substring(0, $len); $kHi = $hi.Substring(0, $len)
    $seq = [Collections.Generic.List[string]]::new()
    try {
        switch ($len) {
            10 { $t = [datetime]::ParseExact($kLo, 'yyyy-MM-dd', $Inv); $e = [datetime]::ParseExact($kHi, 'yyyy-MM-dd', $Inv); while ($t -le $e) { $seq.Add($t.ToString('yyyy-MM-dd', $Inv)); $t = $t.AddDays(1) } }
            7 { $t = [datetime]::ParseExact($kLo, 'yyyy-MM', $Inv); $e = [datetime]::ParseExact($kHi, 'yyyy-MM', $Inv); while ($t -le $e) { $seq.Add($t.ToString('yyyy-MM', $Inv)); $t = $t.AddMonths(1) } }
            4 { for ($y = [int]$kLo; $y -le [int]$kHi; $y++) { $seq.Add('{0:0000}' -f $y) } }
        }
    } catch { $seq.Clear() }   # 0000-00-00 等の不正な日付が範囲の端にある場合は、出現した期間だけを並べる
    $cnt = [Collections.Generic.Dictionary[string, double]]::new(); $under = 0; $over = 0
    for ($i = 0; $i -lt $d.Count; $i++) {
        $k = $d[$i].Substring(0, $len)
        if ([string]::CompareOrdinal($k, $kLo) -lt 0) { $under += $w[$i] } elseif ([string]::CompareOrdinal($k, $kHi) -gt 0) { $over += $w[$i] } else { $cnt[$k] += $w[$i] }
    }
    if (!$seq.Count) { $ks = [string[]]::new($cnt.Count); $cnt.Keys.CopyTo($ks, 0); [Array]::Sort($ks, [StringComparer]::Ordinal); $seq.AddRange($ks) }
    $sfx = if ($len -eq 4) { '年' } else { '' }
    $out = [Collections.Generic.List[object]]::new()
    if ($under) { $out.Add(@(('<' + $kLo + $sfx), $under)) }
    foreach ($k in $seq) { $v = 0.0; [void]$cnt.TryGetValue($k, [ref]$v); $out.Add(@(($k + $sfx), $v)) }
    if ($over) { $out.Add(@(('>' + $kHi + $sfx), $over)) }
    @{ Unit = $unit; Rows = $out }
}
# 文字列の分布: 文字数の分布（コード値の桁ぶれ等が分かる）
function Get-LengthHistogram($lenCnt) {
    $ls = [int[]]::new($lenCnt.Count); $lenCnt.Keys.CopyTo($ls, 0); [Array]::Sort($ls)
    if ($ls[-1] - $ls[0] -le 40) {
        $out = [Collections.Generic.List[object]]::new()
        for ($L = $ls[0]; $L -le $ls[-1]; $L++) { $v = 0; [void]$lenCnt.TryGetValue($L, [ref]$v); $out.Add(@("$L 文字", $v)) }
        return @{ Unit = '文字数'; Rows = $out }
    }
    $w = [int[]]::new($ls.Count); $tot = 0; for ($i = 0; $i -lt $ls.Count; $i++) { $w[$i] = $lenCnt[$ls[$i]]; $tot += $w[$i] }
    $cum = 0; $lo = $ls[0]; $hi = $ls[-1]; $gl = $false
    for ($i = 0; $i -lt $ls.Count; $i++) { $cum += $w[$i]; if (!$gl -and $cum -ge 0.01 * $tot) { $lo = $ls[$i]; $gl = $true }; if ($cum -ge 0.99 * $tot) { $hi = $ls[$i]; break } }
    $h = Get-NumericHistogram ([double[]]$ls) $w $lo $hi $true
    $h.Unit = '文字数'; $h
}

function Add-Warn([string]$table, [string]$col, [string]$cat, [string]$detail, [int]$rows, [string]$sev = '') {
    $r = $Rule[$cat]; if (!$sev) { $sev = $r.Sev }
    if ($r.Demote -and $rows -lt $TH.MinRows) { $sev = switch ($sev) { '重大' { '重大' } '警告' { '注意' } default { '情報' } } }
    $Warns.Add([pscustomobject]@{ Sev = $sev; Table = $table; Col = $col; Cat = $cat; Detail = $detail; Impact = $r.Impact; Fix = $r.Fix })
}
function Add-Index([string]$pri, $tb, [string]$kind, [string[]]$cols, [string]$why, [string]$query, [string]$ddl) {
    $Idx.Add([pscustomobject]@{ Pri = $pri; Table = $tb.Name; Rows = $tb.Rows; Kind = $kind; Cols = $cols; Why = $why; Query = $query; Ddl = $ddl })
}

function Get-ValueType([string[]]$keys) {   # 推定型: 整数・小数・日付・日時・数字文字列（先頭ゼロ付きのコード）・文字列
    if (!$keys.Count) { return '文字列' }
    $notNum = @($keys -notmatch $Re.Num)
    if (!$notNum.Count) { if (@($keys -match '\.').Count) { return '小数' } else { return '整数' } }
    if (!@($keys -notmatch $Re.Date).Count) { if (@($keys -match $Re.Time).Count) { return '日時' } else { return '日付' } }
    if (!@($notNum -notmatch '^[0-9]+$').Count) { return '数字文字列' }
    '文字列'
}
function Get-CardClass([int]$k, [int]$nn, [int]$n) {   # カーディナリティ区分
    if ($n -eq 0) { '-' } elseif ($k -eq 0) { '全NULL' } elseif ($k -eq 1) { '単一値' } elseif ($k -eq $nn) { 'ユニーク' }
    elseif ($k -ge $TH.HighCard -or ($k -ge $TH.LowCard -and $k / $nn -ge 0.5)) { '高' } elseif ($k -ge $TH.LowCard) { '中' } else { '低' }
}

#endregion
#region ── 解析: 項目のプロファイル ─────────────────────────────────────────

function Measure-Column($rows, [int]$c, [string]$table, [string]$name, [bool]$isKey, [bool]$keepKeys) {
    $n = $rows.Count
    $cnt = [Collections.Generic.Dictionary[string, int]]::new(); $nul = 0
    foreach ($r in $rows) { $v = $r[$c]; if (!$v) { $nul++ } else { $cnt[$v]++ } }
    $k = $cnt.Count; $nn = $n - $nul
    $p = @{ Table = $table; Name = $name; Index = $c; N = $n; Null = $nul; NonNull = $nn; Distinct = $k; Class = (Get-CardClass $k $nn $n)
        Type = '-'; Min = $null; Max = $null; Avg = $null; MaxLen = 0; MaxBytes = 0; AvgLen = 0.0
        Mode = $null; ModeCount = 0; ModeShare = $null; Rare = $null; RareCount = 0; Top = $null; TopText = ''
        Hist = $null; Dict = $cnt; FkKeys = $null; FkVals = $null; Prefix = 0; CiDup = 0; Skew = $false; Fk = $null
    }
    if ($k -eq 0) { if ($n) { Add-Warn $table $name '全NULL' "全 $(Format-N $n) 件が NULL" $n }; return $p }

    $keys = [string[]]::new($k); $vals = [int[]]::new($k); $cnt.Keys.CopyTo($keys, 0); $cnt.Values.CopyTo($vals, 0)
    if ($keepKeys) { $p.FkKeys = $keys; $p.FkVals = $vals }   # 外部キー判定用に値を保持

    # 上位値（出現回数の降順）
    $tk = [string[]]$keys.Clone(); $tv = [int[]]$vals.Clone(); [Array]::Sort($tv, $tk)
    $p.Mode = $tk[$k - 1]; $p.ModeCount = $tv[$k - 1]; $p.ModeShare = $tv[$k - 1] / $nn; $p.Rare = $tk[0]; $p.RareCount = $tv[0]
    $top = [Collections.Generic.List[object]]::new()
    for ($i = $k - 1; $i -ge [Math]::Max(0, $k - $TH.TopN); $i--) { $top.Add(@($tk[$i], $tv[$i])) }
    $p.Top = $top
    $p.TopText = if ($k -eq $nn) { '(全件ユニーク)' } else { (($k - 1)..[Math]::Max(0, $k - 3) | ForEach-Object { "$(Format-Cell $tk[$_]) ($($tv[$_]))" }) -join ', ' }

    # 推定型（疑似 NULL を除いて判定。数値・日付が大半で一部だけ異なる形式なら「混在」とし、異なる値を除いて分布を出す）
    $pn = @($keys -match $Re.PseudoNull)
    $core = $keys; if ($pn.Count) { $core = [string[]]@($keys -notmatch $Re.PseudoNull) }   # if 式の出力にすると要素 1 個の配列がスカラーになるため直接代入
    $type = Get-ValueType $core; $good = $core
    if ($type -eq '文字列' -and $core.Count -ge 5) {
        foreach ($try in @(@('数値', $Re.NumLoose), @('日付', $Re.Date))) {
            $bad = [string[]]@($core -notmatch $try[1])
            if ($bad.Count -and $bad.Count -le 0.2 * $core.Count) {
                $br = Get-RowSum $cnt $bad
                if ($br -le 0.2 * $nn) {
                    $good = [string[]]@($core -match $try[1]); $type = (Get-ValueType $good) + '（混在）'
                    Add-Warn $table $name '型混在' "$($try[0])の形式が大半だが、異なる形式の値が $(Format-N $bad.Count) 種類・$(Format-N $br) 件（例: $(Get-Examples $cnt $bad)）" $n
                    break
                }
            }
        }
    }
    $p.Type = $type; $base = $type -replace '（混在）$', ''

    # 文字数・バイト数（文字数はサロゲートペアを 1 文字 = MySQL の CHAR_LENGTH と同じ）
    $maxLen = 0; $maxB = 0; $sumLen = 0.0; $isStr = $base -in '文字列', '数字文字列'
    $lenCnt = [Collections.Generic.Dictionary[int, int]]::new()
    for ($i = 0; $i -lt $k; $i++) {
        $s = $keys[$i]; $L = $s.Length; $sumLen += $L * $vals[$i]
        if ($isStr) { $lenCnt[$L] += $vals[$i] }
        if ($L -gt $maxLen) { $maxLen = [Math]::Max($maxLen, $L - [regex]::Matches($s, $Re.Surrogate).Count) }
        if ($L * 3 -gt $maxB) { $b = $Utf8.GetByteCount($s); if ($b -gt $maxB) { $maxB = $b } }
    }
    $p.MaxLen = $maxLen; $p.MaxBytes = $maxB; $p.AvgLen = $sumLen / $nn

    # 分布・最小値・最大値・平均（対象: 型に合う値）
    if ($good.Count -eq $k) { $hk = $keys; $hv = $vals } else { $hk = $good; $hv = [int[]]::new($good.Count); for ($i = 0; $i -lt $good.Count; $i++) { $hv[$i] = $cnt[$good[$i]] } }
    $m = $hk.Count; $sent = @()
    if ($m) {
        $tot = 0; foreach ($x in $hv) { $tot += $x }
        if ($base -in '整数', '小数') {
            $dk = [double[]]$hk; $ix = [int[]](0..($m - 1)); [Array]::Sort($dk, $ix)
            $sw = [int[]]::new($m); $sum = 0.0; $cum = 0; $lo = $dk[0]; $hi = $dk[$m - 1]; $gl = $false; $gh = $false
            for ($i = 0; $i -lt $m; $i++) {
                $wt = $hv[$ix[$i]]; $sw[$i] = $wt; $sum += $dk[$i] * $wt; $cum += $wt
                if (!$gl -and $cum -ge 0.01 * $tot) { $lo = $dk[$i]; $gl = $true }
                if (!$gh -and $cum -ge 0.99 * $tot) { $hi = $dk[$i]; $gh = $true }
            }
            $p.Min = $hk[$ix[0]]; $p.Max = $hk[$ix[$m - 1]]; $p.Avg = $sum / $tot
            $p.Hist = Get-NumericHistogram $dk $sw $lo $hi ($base -eq '整数')
            $sent = [string[]]@($hk -match $Re.SentinelNum)
        } elseif ($base -in '日付', '日時') {
            # 'yyyy/M/d' 等を 'yyyy-MM-dd' に正規化して並べ替え・集計する
            $j = [string]::Join($SEP, $hk)
            $j = [regex]::Replace($j, '(?<=(?:^|\x01)[0-9]{4})[-/]([0-9]{1,2})[-/]', '-$1-')
            $j = [regex]::Replace($j, '(?<=(?:^|\x01)[0-9]{4}-)([0-9])(?=-)', '0$1')
            $j = [regex]::Replace($j, '(?<=(?:^|\x01)[0-9]{4}-[0-9]{2}-)([0-9])(?![0-9])', '0$1')
            $dn = $j.Split([char]1)
            if ($dn.Count -eq $m) {
                $ix = [int[]](0..($m - 1)); [Array]::Sort($dn, $ix, [StringComparer]::Ordinal)
                $sw = [int[]]::new($m); $cum = 0; $lo = $dn[0]; $hi = $dn[$m - 1]; $gl = $false; $gh = $false
                for ($i = 0; $i -lt $m; $i++) {
                    $wt = $hv[$ix[$i]]; $sw[$i] = $wt; $cum += $wt
                    if (!$gl -and $cum -ge 0.01 * $tot) { $lo = $dn[$i]; $gl = $true }
                    if (!$gh -and $cum -ge 0.99 * $tot) { $hi = $dn[$i]; $gh = $true }
                }
                $p.Min = $hk[$ix[0]]; $p.Max = $hk[$ix[$m - 1]]
                $p.Hist = Get-DateHistogram $dn $sw $lo $hi
            }
            $sent = [string[]]@($hk -match $Re.SentinelDate)
        } else {
            $ss = [string[]]$hk.Clone(); [Array]::Sort($ss, [StringComparer]::Ordinal); $p.Min = $ss[0]; $p.Max = $ss[-1]
            if ($lenCnt.Count) { $p.Hist = Get-LengthHistogram $lenCnt }
            $sent = @()
        }
        if ($sent.Count) {
            $sr = Get-RowSum $cnt $sent
            Add-Warn $table $name '番兵値' "$(Format-N $sent.Count) 種類・$(Format-N $sr) 件（$(Format-Pct ($sr / $nn))）（例: $(Get-Examples $cnt $sent)）" $n
        }
    }

    # 値の形式に関する検査
    if ($pn.Count) { $r = Get-RowSum $cnt $pn; Add-Warn $table $name '疑似NULL' "$(Format-N $r) 件（例: $(Get-Examples $cnt $pn)）" $n }
    if ($base -notin '整数', '小数') {
        $sp = [string[]]@($core -match $Re.Space)
        if ($sp.Count) { $r = Get-RowSum $cnt $sp; Add-Warn $table $name '前後空白' "$(Format-N $sp.Count) 種類・$(Format-N $r) 件（例: $(Get-Examples $cnt $sp)）" $n }
    }
    if ($isStr) {
        $cc = [string[]]@($core -match $Re.Ctrl)
        if ($cc.Count) { $r = Get-RowSum $cnt $cc; Add-Warn $table $name '改行・制御文字' "$(Format-N $cc.Count) 種類・$(Format-N $r) 件（例: $(Get-Examples $cnt $cc 2)）" $n }
        $sg = [string[]]@($core -match $Re.Surrogate)
        if ($sg.Count) { $r = Get-RowSum $cnt $sg; Add-Warn $table $name '4バイト文字' "$(Format-N $sg.Count) 種類・$(Format-N $r) 件（例: $(Get-Examples $cnt $sg 2)）" $n }
        if ($base -eq '数字文字列') {
            $lz = [string[]]@($core -match '^0[0-9]+$')
            if ($lz.Count) { $r = Get-RowSum $cnt $lz; Add-Warn $table $name '先頭ゼロ' "数字のみの値で、先頭が 0 のものが $(Format-N $lz.Count) 種類・$(Format-N $r) 件（例: $(Get-Examples $cnt $lz)）" $n }
        }
        # 表記ゆれ・照合順序（1 本の文字列に連結して一括で正規化し、HashSet で重複の有無だけを先に判定する）
        if ($core.Count -ge 2) {
            $j = [string]::Join($SEP, $core)
            $n1 = [Microsoft.VisualBasic.Strings]::StrConv($j.Normalize([Text.NormalizationForm]::FormKC).ToLowerInvariant(), [Microsoft.VisualBasic.VbStrConv]::Hiragana, 1041).Split([char]1)
            $n2 = [Microsoft.VisualBasic.Strings]::StrConv([regex]::Replace($j.Normalize([Text.NormalizationForm]::FormKD), '\p{Mn}', '').ToLowerInvariant(), [Microsoft.VisualBasic.VbStrConv]::Hiragana, 1041).Split([char]1)
            if ($n1.Count -eq $core.Count -and $n2.Count -eq $core.Count) {
                $s1 = [Collections.Generic.HashSet[string]]::new([string[]]$n1).Count
                $s2 = [Collections.Generic.HashSet[string]]::new([string[]]$n2).Count
                $p.CiDup = $core.Count - $s2
                if ($s1 -lt $core.Count) {
                    $gs = @(Get-VariantGroups $n1)
                    $r = 0; foreach ($g in $gs) { $sm = 0; $mx = 0; foreach ($ix in $g) { $v = $cnt[$core[$ix]]; $sm += $v; if ($v -gt $mx) { $mx = $v } }; $r += $sm - $mx }   # 少数派の表記の件数
                    Add-Warn $table $name '表記ゆれ' "$(Format-N $gs.Count) 組・少数派の表記 $(Format-N $r) 件（例: $(Format-Groups $gs $core $cnt)）" $n
                }
                if ($s2 -lt $s1) {
                    $gs = @(Get-VariantGroups $n2 | Where-Object { [Collections.Generic.HashSet[string]]::new([string[]]@($_ | ForEach-Object { $n1[$_] })).Count -gt 1 })
                    if ($gs.Count) { Add-Warn $table $name '照合順序で同一視' "$(Format-N $gs.Count) 組（例: $(Format-Groups $gs $core $cnt)）" $n }
                }
            }
        }
        # 検索キー向けのプレフィックス長（先頭 n 文字でユニーク数の 95% 以上を区別できる最小の n）
        if ($name -match $Re.SearchName -and $maxLen -gt 64) {
            $j = [string]::Join($SEP, $core); $p.Prefix = 64
            foreach ($len in 8, 12, 16, 20, 24, 32, 48) {
                $tj = [regex]::Replace($j, "(?<=(?:^|\x01)[^\x01]{$len})[^\x01]+", '')
                if ([Collections.Generic.HashSet[string]]::new([string[]]$tj.Split([char]1)).Count -ge 0.95 * $core.Count) { $p.Prefix = $len; break }
            }
        }
    }
    if ($maxB -gt $TH.IndexBytes) { Add-Warn $table $name '長大値' "最大 $(Format-N $maxB) バイト（$(Format-N $maxLen) 文字）" $n }

    # 分布に関する検査
    $share = $p.ModeShare
    if ($k -eq 1 -and $nul -eq 0 -and $n -gt 1) { Add-Warn $table $name '単一値' "全 $(Format-N $n) 件が$(Format-Val $p.Mode)" $n }
    if ($nul -and $nul / $n -ge $TH.NullWarn) { Add-Warn $table $name 'NULL多数' "NULL $(Format-N $nul) 件（$(Format-Pct ($nul / $n))）" $n }
    if ($k -ge 2 -and $nn -ge 100) {
        $avgPer = $nn / $k
        $sev = if ($share -ge $TH.SkewWarn) { '警告' } elseif ($k -ge $TH.LowCard -and $share -ge $TH.SkewNote -and $p.ModeCount -ge $TH.SkewTimes * $avgPer) { '注意' } else { '' }
        if ($sev) {
            $p.Skew = $true
            Add-Warn $table $name '値の偏り' "最頻値$(Format-Val $p.Mode)が $(Format-Pct $share)（$(Format-N $p.ModeCount) 件）を占める。ユニーク数 $(Format-N $k)・1 値あたり平均 $('{0:N1}' -f $avgPer) 件の $('{0:N1}' -f ($p.ModeCount / $avgPer)) 倍" $n $sev
        } elseif ($k -lt $TH.LowCard -and $nn -ge $TH.MinRows) {
            Add-Warn $table $name '低カーディナリティ' "ユニーク数 $k（1 値あたり平均 $(Format-Pct (1 / $k))）。値: $((($p.Top | Select-Object -First 4) | ForEach-Object { (Format-Val $_[0] 12) + ' ' + (Format-Pct ($_[1] / $nn)) }) -join '、')" $n
        }
    }
    if (!$isKey -and $k -lt $nn -and $k -ge 0.99 * $nn -and $nn -ge 100 -and ($name -match $Re.IdName -or $name -match $Re.UniqName)) {
        $dups = [string[]]@(for ($i = $k - 1; $i -ge 0 -and $tv[$i] -gt 1 -and $i -ge $k - 3; $i--) { $tk[$i] })
        Add-Warn $table $name 'ほぼユニーク' "非 NULL $(Format-N $nn) 件に対しユニーク数 $(Format-N $k)。重複 $(Format-N ($nn - $k)) 件（例: $(Get-Examples $cnt $dups)）" $n
    }
    $p
}

#endregion
#region ── 解析: CSV の読込 ─────────────────────────────────────────────────

$Tables = [Collections.Generic.List[object]]::new()
$Warns = [Collections.Generic.List[object]]::new()
$Idx = [Collections.Generic.List[object]]::new()
$Pre = [Collections.Generic.List[object]]::new()
$All = [Collections.Generic.List[object]]::new()
$total = 0

function Get-Encoding([string]$path) {   # 文字コード判定（UTF-8 として不正なら Shift_JIS）
    try { [void][Text.UTF8Encoding]::new($false, $true).GetString([IO.File]::ReadAllBytes($path)); [Text.Encoding]::UTF8 } catch { [Text.Encoding]::GetEncoding(932) }
}
function New-Parser($source) { $p = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($source); $p.SetDelimiters(','); $p.TrimWhiteSpace = $false; $p }

$files = @(Get-ChildItem -LiteralPath $CsvDir -Filter *.csv | Sort-Object Name)
if (!$files) { throw "CSV ファイルがありません: $CsvDir" }

# 事前に全ファイルの見出しを読み、各テーブルのキー項目名を集める（外部キー候補の列を見分けるため）
$KeyNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$EncOf = @{}
foreach ($f in $files) {
    $EncOf[$f.FullName] = Get-Encoding $f.FullName
    $sr = [IO.StreamReader]::new($f.FullName, $EncOf[$f.FullName]); $line = $sr.ReadLine(); $sr.Close()
    if ($line) { $h = (New-Parser ([IO.StringReader]::new($line))).ReadFields(); $kc = [string]$KeyColumn[$f.BaseName]; [void]$KeyNames.Add($(if ($kc -and $h -contains $kc) { $kc } else { $h[0] })) }
}

Write-Host ''
Write-Host ('  {0,-34}{1,12}{2,8}{3,10}' -f 'テーブル', '件数', '項目', '秒')
foreach ($f in $files) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $p = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($f.FullName, $EncOf[$f.FullName]); $p.SetDelimiters(','); $p.TrimWhiteSpace = $false   # "" 囲み内のカンマ・改行に対応
    $t = [string]$f.BaseName; $hdr = $p.ReadFields(); $rows = [Collections.Generic.List[string[]]]::new()
    if ($hdr) {
        $m = $hdr.Count
        while (!$p.EndOfData) {   # 列が足りない行は不足分を NULL とする
            $fl = $p.ReadFields(); if ($fl.Count -lt $m) { $tmp = [string[]]::new($m); [Array]::Copy($fl, $tmp, $fl.Count); $fl = $tmp }
            $rows.Add($fl)
        }
    }
    $p.Close()
    if (!$hdr) { Write-Warning "空ファイルのためスキップ: $($f.Name)"; continue }
    $n = $rows.Count; $total += $n; $ki = [Math]::Max(0, [Array]::IndexOf($hdr, [string]$KeyColumn[$t]))
    $tb = @{ Name = $t; Rows = $n; Header = $hdr; KeyIdx = $ki; KeyName = $hdr[$ki]; Profiles = $null; PkCols = @(); PkSet = $null }

    # 項目別のプロファイル
    $tb.Profiles = @(for ($c = 0; $c -lt $m; $c++) {
            Measure-Column $rows $c $t $hdr[$c] ($c -eq $ki) ($hdr[$c] -match $Re.IdName -or $KeyNames.Contains($hdr[$c]))
        })

    # 主キーの推定: 一意かつ NULL なしの列（キー項目を優先）→ なければ 2 列の組合せ
    $pk = @()
    if ($n) {
        # 1) キー項目または ID らしい単独列（ID 系の名前・整数・数字コード）→ 2) 2 列の組合せ → 3) その他の一意な単独列（名称等の偶然の一意を後回しにする）
        $uniq = @($tb.Profiles | Where-Object { $_.Null -eq 0 -and $_.Distinct -eq $n } |
                Sort-Object @{ Expression = { [int]($_.Index -eq $ki) * 8 + [int]($_.Name -match $Re.IdName) * 4 + [int]($_.Type -in '整数', '数字文字列') * 2 + [int]($_.MaxLen -le 36) }; Descending = $true }, @{ Expression = { $_.Index } })
        $strong = @($uniq | Where-Object { $_.Index -eq $ki -or $_.Name -match $Re.IdName -or $_.Type -in '整数', '数字文字列' })
        if ($strong.Count) { $pk = @($strong[0].Name) }
        else {
            $cand = @($tb.Profiles | Where-Object { $_.Null -eq 0 -and $_.Distinct -ge 2 -and $_.Distinct -lt $n -and $_.Type -ne '小数' -and ($_.Index -lt 5 -or $_.Name -match $Re.IdName) } | Select-Object -First 8)
            $pairs = [Collections.Generic.List[object]]::new()
            foreach ($withKey in $true, $false) {
                for ($i = 0; $i -lt $cand.Count; $i++) { for ($j = $i + 1; $j -lt $cand.Count; $j++) { if ((($cand[$i].Index -eq $ki) -or ($cand[$j].Index -eq $ki)) -eq $withKey) { $pairs.Add(@($cand[$i], $cand[$j])) } } }
            }
            foreach ($pr in $pairs) {
                if ([double]$pr[0].Distinct * $pr[1].Distinct -lt $n) { continue }
                $hs = [Collections.Generic.HashSet[string]]::new(); $ok = $true; $ia = $pr[0].Index; $ib = $pr[1].Index
                foreach ($r in $rows) { if (!$hs.Add($r[$ia] + $SEP + $r[$ib])) { $ok = $false; break } }
                if ($ok) { $pk = @($pr[0].Name, $pr[1].Name); break }
            }
            if (!$pk.Count -and $uniq.Count) { $pk = @($uniq[0].Name) }
        }
    }
    $tb.PkCols = $pk
    if ($pk.Count) { [void]$Warns.RemoveAll([Predicate[object]] { param($w) $w.Table -eq $t -and $w.Cat -eq '低カーディナリティ' -and $pk -contains $w.Col }) }
    if ($pk.Count -eq 1) { $tb.PkSet = ($tb.Profiles | Where-Object { $_.Name -eq $pk[0] } | Select-Object -First 1).Dict }
    foreach ($q in $tb.Profiles) { $q.Dict = $null }   # 主キー以外の値の集合は解放

    # キー項目の検査（主キー候補が別にある場合は重要度を下げる）
    $kp = $tb.Profiles[$ki]
    if (!$n) { Add-Warn $t '(テーブル)' 'データなし' 'CSV にデータ行がない' 0 }
    else {
        $pkText = if ($pk.Count) { "（主キー候補は $($pk -join ' + ')）" } else { '' }
        if ($kp.Null) { Add-Warn $t $kp.Name 'キー空白' "NULL $(Format-N $kp.Null) 件（$(Format-Pct ($kp.Null / $n))）$pkText" $n $(if ($pk.Count -and $pk -notcontains $kp.Name) { '注意' } else { '' }) }
        if ($kp.Distinct -lt $kp.NonNull) {
            $dups = [string[]]@($kp.Top | Where-Object { $_[1] -gt 1 } | Select-Object -First 3 | ForEach-Object { $_[0] })
            $ex = ($dups | ForEach-Object { Format-Val $_ }) -join '、'
            if ($pk.Count -gt 1 -and $pk[0] -eq $kp.Name) { Add-Warn $t $kp.Name 'キー重複' "重複 $(Format-N ($kp.NonNull - $kp.Distinct)) 件（例: $ex）。複合主キー候補（$($pk -join ' + ')）の先頭列のため想定どおり" $n '情報' }
            else { Add-Warn $t $kp.Name 'キー重複' "重複 $(Format-N ($kp.NonNull - $kp.Distinct)) 件（例: $ex）$pkText" $n $(if ($pk.Count) { '注意' } else { '' }) }
        }
        if (!$pk.Count) { Add-Warn $t '(テーブル)' '主キー候補なし' "全 $(Format-N $m) 列を調べたが、一意かつ NULL なしの列・2 列の組合せがない" $n }
    }

    # キー項目の先頭1桁カテゴリ別件数（サロゲートペア・異体字セレクタ付きの字も 1 文字として扱う）
    $pc = [Collections.Generic.Dictionary[string, int]]::new()
    foreach ($r in $rows) { $v = [string]$r[$ki]; $pc[$(if ($v) { [Globalization.StringInfo]::GetNextTextElement($v) } else { '(空白)' })]++ }
    $pk2 = [string[]]$pc.Keys; [Array]::Sort($pk2, [StringComparer]::Ordinal)
    foreach ($x in $pk2) { $Pre.Add(@($t, $hdr[$ki], (ConvertTo-CellValue $x), $pc[$x], ($pc[$x] / $n))) }

    # 全データ（検索用）: 1 レコード =「列名=値 | 列名=値 …」
    $fmt = (0..($m - 1) | ForEach-Object { ($hdr[$_] -replace '[{}]', '$0$0') + "={$_}" }) -join ' | '
    for ($i = 0; $i -lt $n -and $All.Count -lt $LIMIT; $i++) { $All.Add(@($t, ($i + 1), (Cut (($fmt -f $rows[$i]) -replace '[\r\n]+', ' ') 32000))) }

    $Tables.Add($tb); $rows = $null
    Write-Host ('  {0,-34}{1,12:N0}{2,8}{3,10:N1}' -f (Cut $t 32), $n, $m, $sw.Elapsed.TotalSeconds)
}
if (!$Tables.Count) { throw "読み込めるデータがありません: $CsvDir" }

#endregion
#region ── 解析: 外部キー・インデックス提案 ─────────────────────────────────

function Get-Stem([string]$t) {   # テーブル名の語幹（接頭辞 m_ 等と複数形の s を除く）
    $s = $t.ToLowerInvariant() -replace '^(m|mst|t|tbl|trn|v)_', ''
    if ($s -match 'ies$') { $s -replace 'ies$', 'y' } elseif ($s -match '(ss|x|ch|sh)es$') { $s -replace 'es$', '' } elseif ($s -match '[^s]s$') { $s -replace 's$', '' } else { $s }
}
function Test-FkName([string]$col, $par) {   # 列名が参照先テーブルのキーに対応するか（customer_id ↔ customers.customer_id / customers.id）
    $c = $col.ToLowerInvariant(); $pk = ([string]$par.PkCols[0]).ToLowerInvariant(); $t = $par.Name.ToLowerInvariant(); $st = Get-Stem $par.Name
    if ($pk -notin 'id', 'no', 'cd', 'code', 'key' -and $c -eq $pk) { return $true }
    $c -in "$($st)_$pk", "$st$pk", "$($t)_$pk", "$t$pk"
}
foreach ($tb in $Tables) {
    foreach ($p in $tb.Profiles) {
        if (!$p.FkKeys) { continue }
        $best = $null
        foreach ($par in $Tables) {
            if ($par -eq $tb -or $par.PkCols.Count -ne 1 -or !$par.PkSet -or !(Test-FkName $p.Name $par)) { continue }
            $orph = [Collections.Generic.List[string]]::new()
            foreach ($x in $p.FkKeys) { if (!$par.PkSet.ContainsKey($x)) { $orph.Add($x) } }
            $rate = 1 - $orph.Count / $p.FkKeys.Count
            if ($rate -ge 0.8 -and (!$best -or $rate -gt $best.Rate)) { $best = @{ Table = $par.Name; Col = $par.PkCols[0]; Rate = $rate; Orphans = $orph.ToArray() } }
        }
        if ($best) {
            $p.Fk = $best
            if ($best.Orphans.Count) {
                $dict = [Collections.Generic.Dictionary[string, int]]::new(); for ($i = 0; $i -lt $p.FkKeys.Count; $i++) { $dict[$p.FkKeys[$i]] = $p.FkVals[$i] }
                $r = Get-RowSum $dict $best.Orphans
                Add-Warn $tb.Name $p.Name '参照整合性' "参照先 $($best.Table).$($best.Col) にない値が $(Format-N $best.Orphans.Count) 種類・$(Format-N $r) 件（例: $(Get-Examples $dict $best.Orphans)）" $tb.Rows
            }
        }
        $p.FkKeys = $null; $p.FkVals = $null
    }
}
foreach ($tb in $Tables) { $tb.PkSet = $null }

foreach ($tb in $Tables) {
    $n = $tb.Rows; if (!$n) { continue }
    $t = $tb.Name; $tq = Format-Id $t; $ps = $tb.Profiles
    $sizePri = if ($n -ge $TH.BigRows) { '高' } elseif ($n -ge $TH.MinRows) { '中' } else { '低' }
    $pk1 = if ($tb.PkCols.Count -eq 1) { $tb.PkCols[0] } else { '' }
    $lead = [Collections.Generic.HashSet[string]]::new()   # 提案済みインデックスの先頭列（重複した提案を防ぐ）
    $ci = { param($q) if ($q.CiDup) { " ※ utf8mb4_0900_ai_ci では大文字小文字・全角半角等を同一視するため、$(Format-N $q.CiDup) 件が重複扱いになる（照合順序の確認が必要）" } else { '' } }

    # 1. 主キー
    if ($tb.PkCols.Count) {
        $why = if ($tb.PkCols.Count -gt 1) { "単独で一意な列がなく、$($tb.PkCols -join ' + ') の組合せが全 $(Format-N $n) 件で一意・NULL なし" } else { "全 $(Format-N $n) 件で一意・NULL なし" }
        $why += '（定義済みなら対応不要）' + $(if ($tb.PkCols.Count -eq 1) { & $ci ($ps | Where-Object { $_.Name -eq $pk1 }) } else { '' })
        Add-Index '高' $tb '主キー' $tb.PkCols $why "WHERE $($tb.PkCols[0]) = ? ／ 結合の基点" "ALTER TABLE $tq ADD PRIMARY KEY ($(($tb.PkCols | ForEach-Object { Format-Id $_ }) -join ', '));"
        [void]$lead.Add($tb.PkCols[0])
    } else {
        $idc = if ($tb.Header -contains 'id') { 'row_id' } else { 'id' }
        Add-Index '高' $tb '主キー（サロゲート）' @($idc) '一意かつ NULL なしの列・2 列の組合せがない。InnoDB は主キーがないと内部の行 ID でデータを管理する' '行の特定・更新・レプリケーション' "ALTER TABLE $tq ADD COLUMN $(Format-Id $idc) BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY FIRST;"
    }
    # 2. 一意インデックス（ID・コード・メール等の名前で全件一意）
    foreach ($p in $ps) {
        if ($p.Name -eq $pk1 -or $p.Distinct -ne $p.NonNull -or $p.NonNull -lt 2 -or $p.Type -in '小数', '日時' -or $p.MaxLen -gt 255 -or $p.Name -notmatch $Re.UniqName) { continue }
        Add-Index '中' $tb '一意インデックス' @($p.Name) ("全 $(Format-N $p.NonNull) 件（NULL 除く）で一意。重複の混入を防ぎ、等価検索を高速化する" + (& $ci $p)) "WHERE $($p.Name) = ?" "ALTER TABLE $tq ADD UNIQUE INDEX $(Format-Id (Get-IndexName 'uq' $t @($p.Name))) ($(Format-Id $p.Name));"
        [void]$lead.Add($p.Name)
    }
    # 3. 結合・外部キー（複合主キーの先頭列は主キーで代用できるため除く）
    foreach ($p in $ps) {
        if (!$p.Fk -or $lead.Contains($p.Name)) { continue }
        $low = if ($p.Distinct -lt $TH.LowCard) { "。ユニーク数 $($p.Distinct) と少なく絞り込み効果は小さいが、結合・外部キー制約の検査に必要" } else { '' }
        Add-Index $sizePri $tb '結合・外部キー' @($p.Name) "$($p.Fk.Table).$($p.Fk.Col) を参照（値の一致率 $(Format-Pct $p.Fk.Rate)）$low" "JOIN $($p.Fk.Table) ON $($p.Fk.Table).$($p.Fk.Col) = $t.$($p.Name)" "ALTER TABLE $tq ADD INDEX $(Format-Id (Get-IndexName 'ix' $t @($p.Name))) ($(Format-Id $p.Name));"
        [void]$lead.Add($p.Name)
    }
    # 4. 区分＋日付の複合インデックス／日付の単独インデックス
    if ($n -ge $TH.MinRows) {
        $dateCol = $ps | Where-Object { $_.Type -in '日付', '日時' -and $_.Distinct -ge 10 -and $_.Null / $n -lt 0.5 -and $_.Name -notmatch $Re.DateSkip } |
            Sort-Object { - [int]($_.Name -match $Re.DateName) }, { - $_.Distinct } | Select-Object -First 1
        $catCol = $ps | Where-Object { $_.Distinct -ge 2 -and $_.Distinct -le 30 -and $_.NonNull -ge 0.5 * $n -and ($_.Type -replace '（混在）$', '') -in '文字列', '整数', '数字文字列' -and $_.Name -ne $pk1 } |
            Sort-Object { - [int]($_.Name -match $Re.CatName) }, { $_.ModeShare }, { $_.Distinct } | Select-Object -First 1
        if ($dateCol -and $catCol) {
            $cols = @($catCol.Name, $dateCol.Name)
            Add-Index $(if ($sizePri -eq '高') { '高' } else { '中' }) $tb '複合インデックス' $cols "$(Format-Val $catCol.Name 20)（ユニーク数 $($catCol.Distinct)）で絞り込み、$(Format-Val $dateCol.Name 20)で期間指定・並べ替えする検索向け。等価条件の列を先頭、範囲条件の列を後ろに置く" "WHERE $($catCol.Name) = ? AND $($dateCol.Name) BETWEEN ? AND ? ORDER BY $($dateCol.Name)" "ALTER TABLE $tq ADD INDEX $(Format-Id (Get-IndexName 'ix' $t $cols)) ($(($cols | ForEach-Object { Format-Id $_ }) -join ', '));"
            [void]$lead.Add($catCol.Name)
        }
        if ($dateCol -and !$lead.Contains($dateCol.Name)) {
            Add-Index $(if ($catCol) { '低' } else { '中' }) $tb '期間検索' @($dateCol.Name) "ユニーク数 $(Format-N $dateCol.Distinct)・期間 $($dateCol.Min) ～ $($dateCol.Max)。期間指定・日付順の並べ替えに有効$(if ($catCol) { '（区分を指定しない期間検索がある場合）' } else { '' })" "WHERE $($dateCol.Name) BETWEEN ? AND ?" "ALTER TABLE $tq ADD INDEX $(Format-Id (Get-IndexName 'ix' $t @($dateCol.Name))) ($(Format-Id $dateCol.Name));"
            [void]$lead.Add($dateCol.Name)
        }
    }
    foreach ($p in $ps) {
        if ($p.Name -eq $pk1 -or !$p.Distinct) { continue }
        $base = $p.Type -replace '（混在）$', ''
        # 5. 少数値の検索用（偏った区分値のうち、少数派の値の抽出）
        if ($n -ge 10000 -and $p.Distinct -ge 2 -and $p.Distinct -lt $TH.LowCard -and $p.RareCount / $p.NonNull -le 0.05 -and !$lead.Contains($p.Name)) {
            Add-Index '低' $tb '少数値検索' @($p.Name) "$(Format-Val $p.Rare)は全体の $(Format-Pct ($p.RareCount / $p.NonNull))（$(Format-N $p.RareCount) 件）のみ。この値の検索に限り有効（最頻値$(Format-Val $p.Mode)の検索では使われない）" "WHERE $($p.Name) = '$($p.Rare)'" "ALTER TABLE $tq ADD INDEX $(Format-Id (Get-IndexName 'ix' $t @($p.Name))) ($(Format-Id $p.Name));"
            [void]$lead.Add($p.Name)
        }
        # 6. 検索キー（名称・カナ・連絡先）
        $isName = $base -in '文字列', '数字文字列' -and $p.Name -match $Re.SearchName
        $isId = $base -in '整数', '文字列', '数字文字列' -and $p.Name -match $Re.IdName -and !$p.Fk
        if ($n -ge $TH.MinRows -and ($isName -or $isId) -and $p.Distinct -ge $TH.HighCard -and !$lead.Contains($p.Name)) {
            $colDef = if ($p.Prefix) { "$(Format-Id $p.Name)($($p.Prefix))" } else { Format-Id $p.Name }
            $preTxt = if ($p.Prefix) { "。最大 $($p.MaxLen) 文字と長いため、先頭 $($p.Prefix) 文字のプレフィックスインデックスとする（95% 以上を区別可能）" } else { '' }
            $what = if ($isName) { "名称・連絡先などの検索キー（ユニーク数 $(Format-N $p.Distinct)）。完全一致・前方一致に有効" } else { "ID・コード系の列（ユニーク数 $(Format-N $p.Distinct)・1 値あたり平均 $('{0:N1}' -f ($p.NonNull / $p.Distinct)) 件）。参照先は CSV から特定できないが、等価検索・結合に使われやすい" }
            Add-Index $(if ($n -ge $TH.BigRows) { '中' } else { '低' }) $tb '検索キー' @($p.Name) "$what$preTxt" $(if ($isName) { "WHERE $($p.Name) = ? ／ $($p.Name) LIKE 'abc%'" } else { "WHERE $($p.Name) = ?" }) "ALTER TABLE $tq ADD INDEX $(Format-Id (Get-IndexName 'ix' $t @($p.Name))) ($colDef);"
            [void]$lead.Add($p.Name)
        }
        # 7. 全文検索（文章系の列）
        if ($n -ge $TH.MinRows -and $base -eq '文字列' -and $p.Distinct -ge $TH.HighCard -and (($p.Name -match $Re.TextName -and $p.AvgLen -ge 10) -or $p.AvgLen -ge 50)) {
            Add-Index '低' $tb '全文検索' @($p.Name) "平均 $('{0:N0}' -f $p.AvgLen) 文字・最大 $(Format-N $p.MaxLen) 文字の文章系の列。部分一致（LIKE '%語%'）は B-tree インデックスを使えず全件走査になる" "WHERE MATCH($($p.Name)) AGAINST('語' IN BOOLEAN MODE)" "ALTER TABLE $tq ADD FULLTEXT INDEX $(Format-Id (Get-IndexName 'ft' $t @($p.Name))) ($(Format-Id $p.Name)) WITH PARSER ngram;"
        }
        # 8. ヒストグラム（値の偏りがある列）
        if ($p.Skew -and $n -ge $TH.MinRows) {
            $bk = if ($p.Distinct -le 1024) { $p.Distinct } else { 100 }
            Add-Index $(if ($n -ge $TH.BigRows) { '中' } else { '低' }) $tb 'ヒストグラム' @($p.Name) "最頻値$(Format-Val $p.Mode)が $(Format-Pct $p.ModeShare) を占める。値ごとの件数を統計に持たせ、行数の見積り誤りによる実行計画の悪化を防ぐ" "WHERE $($p.Name) = ?（インデックスのない条件・結合）" "ANALYZE TABLE $tq UPDATE HISTOGRAM ON $(Format-Id $p.Name) WITH $bk BUCKETS;"
        }
    }
}

# 項目別の「インデックス」「注意」の要約
$IdxTag = @{}; $WarnTag = @{}
foreach ($x in $Idx) {
    for ($i = 0; $i -lt $x.Cols.Count; $i++) {
        $key = "$($x.Table)`t$($x.Cols[$i])"; $tag = $IdxKind[$x.Kind].Tag + $(if ($x.Cols.Count -gt 1) { "($($i + 1)列目)" } else { '' })
        if (!$IdxTag[$key]) { $IdxTag[$key] = [Collections.Generic.List[string]]::new() }; if (!$IdxTag[$key].Contains($tag)) { $IdxTag[$key].Add($tag) }
    }
}
$WarnsSorted = @($Warns | Sort-Object { $SevOrder[$_.Sev] }, Table, Col)
foreach ($w in $WarnsSorted) { $key = "$($w.Table)`t$($w.Col)"; if (!$WarnTag[$key]) { $WarnTag[$key] = [Collections.Generic.List[string]]::new() }; if (!$WarnTag[$key].Contains($w.Cat)) { $WarnTag[$key].Add($w.Cat) } }
$kindOrder = @{}; $i = 0; foreach ($k in $IdxKind.Keys) { $kindOrder[$k] = $i++ }
$IdxSorted = @($Idx | Sort-Object { $PriOrder[$_.Pri] }, @{ Expression = { $_.Rows }; Descending = $true }, Table, { $kindOrder[$_.Kind] })

#endregion
#region ── 出力データの組み立て ─────────────────────────────────────────────

$maxRows = ($Tables | ForEach-Object { $_.Rows } | Measure-Object -Maximum).Maximum
$logMax = [Math]::Max(1.0, [Math]::Log10($maxRows + 1))
$Sum = [Collections.Generic.List[object]]::new(); $Dist = [Collections.Generic.List[object]]::new()
$Top = [Collections.Generic.List[object]]::new(); $Hist = [Collections.Generic.List[object]]::new()
foreach ($tb in $Tables) {
    $t = $tb.Name; $n = $tb.Rows
    $tw = @($Warns | Where-Object { $_.Table -eq $t })
    $sc = @{}; foreach ($s in '重大', '警告', '注意') { $sc[$s] = @($tw | Where-Object { $_.Sev -eq $s }).Count }
    $find = (@($tw | Where-Object { $_.Sev -in '重大', '警告' }) | Group-Object Cat | Sort-Object { $SevOrder[$_.Group[0].Sev] }, { - $_.Count } | ForEach-Object { $_.Name + $(if ($_.Count -gt 1) { "×$($_.Count)" } else { '' }) }) -join '／'
    $pkText = if ($tb.PkCols.Count) { $tb.PkCols -join ' + ' } elseif ($n) { '（なし）' } else { '' }
    $Sum.Add(@($t, $n, [Math]::Log10($n + 1), $tb.Header.Count, $tb.KeyName, $tb.Profiles[$tb.KeyIdx].Distinct, $pkText, @($Idx | Where-Object { $_.Table -eq $t }).Count, $sc['重大'], $sc['警告'], $sc['注意'], $find))
    foreach ($p in $tb.Profiles) {
        $key = "$t`t$($p.Name)"; $nn = $p.NonNull; $k = $p.Distinct; $uq = $k -eq $nn -and $nn -gt 1
        $Dist.Add(@($t, $p.Name, $p.Type, $nn, $p.Null, $(if ($n) { $p.Null / $n } else { $null }), $k, [Math]::Log10($k + 1), $p.Class,
                $(if ($nn) { $k / $nn } else { $null }), $(if ($k) { $nn / $k } else { $null }), $(if ($uq) { '' } else { ConvertTo-CellValue (Format-Cell $p.Mode) }), $(if ($uq -or $null -eq $p.ModeShare) { '' } else { $p.ModeShare }),
                (ConvertTo-CellValue (Format-Cell $p.Min)), (ConvertTo-CellValue (Format-Cell $p.Max)), $p.Avg, $p.MaxLen, $p.MaxBytes, $p.TopText,
                $(if ($IdxTag[$key]) { $IdxTag[$key] -join '／' } else { '' }), $(if ($WarnTag[$key]) { $WarnTag[$key] -join '／' } else { '' })))
        # 上位値（構成比は NULL を含む全件に対する比率。最頻値が非 NULL の 50% 以上なら「偏り」）
        $cum = 0.0; $rank = 0; $shown = 0
        if ($uq) { $shown = $nn; $cum = $nn / $n; $Top.Add(@($t, $p.Name, $null, "(全件ユニーク $(Format-N $k) 種類)", $nn, ($nn / $n), $cum, '')) }
        elseif ($p.Top) {
            foreach ($e in $p.Top) {
                $rank++; $shown += $e[1]; $cum += $e[1] / $n
                $Top.Add(@($t, $p.Name, $rank, (ConvertTo-CellValue (Format-Cell $e[0])), $e[1], ($e[1] / $n), $cum, $(if ($rank -eq 1 -and $p.Skew -and $p.ModeShare -ge $TH.SkewWarn) { '偏り' } else { '' })))
            }
        }
        if ($nn -gt $shown) { $cum += ($nn - $shown) / $n; $Top.Add(@($t, $p.Name, $null, "(その他 $(Format-N ($k - $rank)) 種類)", ($nn - $shown), (($nn - $shown) / $n), $cum, '')) }
        if ($p.Null) { $cum += $p.Null / $n; $Top.Add(@($t, $p.Name, $null, '(NULL)', $p.Null, ($p.Null / $n), $cum, '')) }
        if ($p.Hist) { $no = 0; foreach ($h in $p.Hist.Rows) { $no++; $Hist.Add(@($t, $p.Name, $p.Hist.Unit, $no, [string]$h[0], $h[1], $(if ($nn) { $h[1] / $nn } else { 0 }))) } }
    }
}
$Profiles = @($Tables | ForEach-Object { $_.Profiles })
$WarnRows = @($WarnsSorted | ForEach-Object { , @($_.Sev, $_.Table, $_.Col, $_.Cat, $_.Detail, $_.Impact, $_.Fix) })
$IdxRows = @($IdxSorted | ForEach-Object { , @($_.Pri, $_.Table, $_.Rows, $_.Kind, ($_.Cols -join ', '), $_.Why, $_.Query, $_.Ddl) })

# 判定基準
$RuleRows = [Collections.Generic.List[object]]::new()
foreach ($e in @(
        @('ユニーク', 'ユニーク数 = 非 NULL 件数', '1 つの値で 1 件に絞り込める', '主キー・一意制約の候補'),
        @('高', "ユニーク数 $($TH.HighCard) 以上、または $($TH.LowCard) 以上かつユニーク率 50% 以上", '等価検索の選択度が高く、インデックスが有効', '検索条件・結合キーならインデックスを作成'),
        @('中', "ユニーク数 $($TH.LowCard)～$($TH.HighCard - 1)", '1 値あたり平均で全体の 1～10% 程度が該当', '単独よりも複合インデックスの一部として使う'),
        @('低', "ユニーク数 2～$($TH.LowCard - 1)", '1 値あたり平均で全体の 10% 超が該当し、単独のインデックスはほぼ使われない', '複合インデックスの先頭列、または少数派の値の検索だけに使う'),
        @('単一値', '非 NULL の値が 1 種類', '検索条件として機能しない', '列の要否を見直す'),
        @('全NULL', '全件 NULL', '検索条件として機能しない', '未使用の列なら削除を検討'))) { $RuleRows.Add(@('カーディナリティ', $e[0], '', $e[1], $e[2], $e[3])) }
foreach ($e in $Rule.GetEnumerator()) { $RuleRows.Add(@('検索リスク警告', $e.Key, $e.Value.Sev, $e.Value.Cond, $e.Value.Impact, $e.Value.Fix)) }
foreach ($e in $IdxKind.GetEnumerator()) { $RuleRows.Add(@('インデックス提案', $e.Key, '', $e.Value.Cond, $e.Value.Aim, '')) }
foreach ($e in @(
        @('推定の前提', 'CSV のデータだけから推定している。既存のインデックス・実際の SQL・データ型定義は考慮しないため、作成前に EXPLAIN で効果を確認する'),
        @('重要度の調整', "性能面の警告（値の偏り・NULL 多数）は、$(Format-N $TH.MinRows) 件未満のテーブルでは 1 段階下げる（全件走査でも高速なため）"),
        @('照合順序', 'Aurora MySQL 8.x の既定 utf8mb4_0900_ai_ci（大文字小文字・全角半角・ひらがなカタカナ・濁点を区別しない、末尾空白を無視しない）を想定'),
        @('NULL の扱い', 'CSV の空値を NULL として集計する。「NULL」等の文字列は疑似 NULL として警告する'),
        @('値の分布', "数値は 1～99 パーセンタイルを約 $($TH.HistBins) 区間、日付は期間に応じて日別・月別・年別、文字列は文字数で集計。範囲外の値は「<」「≥」の区間にまとめる"))) { $RuleRows.Add(@('前提', $e[0], '', $e[1], '', '')) }

#endregion
#region ── Excel 出力 ────────────────────────────────────────────────────────

function Invoke-Retry([scriptblock]$sb) {   # Excel が処理中で呼出しを拒否した場合（RPC_E_CALL_REJECTED 等）は待って再試行
    for ($i = 0; ; $i++) {
        try { return & $sb } catch {
            $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
            if ($i -ge 40 -or $e.HResult -notin 0x80010001, 0x8001010A) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

Write-Host ''
Write-Host '  Excel を作成しています…'
$xl = New-Object -ComObject Excel.Application
try {
    $xl.ScreenUpdating = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add(-4167); $init = $wb.Worksheets.Item(1)
    $xl.Calculation = -4135   # 作成中は手動計算（最後に自動へ戻す）
    foreach ($fs in $wb.Theme.ThemeFontScheme.MajorFont, $wb.Theme.ThemeFontScheme.MinorFont) { foreach ($i in 1, 3) { $fs.Item($i).Name = 'Meiryo UI' } }   # グラフ・図形の既定フォント
    $fo = $wb.Styles.Item('Normal').Font; $fo.Name = 'Meiryo UI'; $fo.Size = 10; $fo.Color = $Color.Ink   # 標準スタイル = 全セル既定の書式
    $ts = $wb.TableStyles.Add('Report'); $ts.TableStyleElements.Item(0).Font.Color = $Color.Ink; $ts.TableStyleElements.Item(5).Interior.Color = $Color.Stripe   # 表: 文字 #333333・縞 #F2F2F2
    $script:IgnoreBudget = 1500   # 緑三角の非表示はセル単位で約 10 ms かかるため上限を設ける

    #── セル・表の部品 ──────────────────────────────────────────────
    # 注意: Value2 へは [string] 等の素の値か [object[,]] だけを代入する。PSObject（型なし引数など）を一度でも代入すると、
    #       PowerShell 5.1 の COM バインダーがその変換規則を使い回し、以降の 2 次元配列の一括代入が失敗するため
    function Set-Head($ws, [string[]]$head, [int]$row = 4, [int]$col = 1) {   # 見出し（ミッドナイトネイビー地・白太字）
        for ($c = 0; $c -lt $head.Count; $c++) { $ws.Cells.Item($row, $col + $c).Value2 = [string]$head[$c] }
        $r = $ws.Cells.Item($row, $col).Resize(1, $head.Count)
        $r.Interior.Color = $Color.Navy; $r.Font.Color = $Color.White; $r.Font.Bold = $true; $r.HorizontalAlignment = -4108; $r.VerticalAlignment = -4108; $r.WrapText = $true; $r.RowHeight = 26
    }
    function Add-Link($ws, [string]$cell, [string]$text, [string]$sheet) {
        $r = $ws.Range($cell); [void]$ws.Hyperlinks.Add($r, '', "'$sheet'!A1", "「$sheet」シートへ移動", $text)
        $r.Font.Color = $Color.Navy; $r.Font.Underline = -4142; $r.Font.Size = 9
    }
    function New-Sheet([string]$name, [string]$title, [string]$note, [int]$tab = $Color.Navy, [int]$freezeCol = 0) {   # 1 行目タイトル / 2 行目説明 / 4 行目見出し（ここまで固定表示）
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $ws.Name = $name; $ws.Tab.Color = $tab
        $ws.Range('A1').Value2 = $title; $f = $ws.Range('A1').Font; $f.Size = 16; $f.Bold = $true; $f.Color = $Color.Navy; $ws.Rows.Item(1).RowHeight = 32; $ws.Rows.Item(1).VerticalAlignment = -4108
        $ws.Range('A2').Value2 = $note; $ws.Range('A2').Font.Size = 9; $ws.Range('A2').Font.Color = $Color.Ink2
        $w = $xl.ActiveWindow; $w.DisplayGridlines = $false; $w.SplitRow = 4; $w.SplitColumn = $freezeCol; $w.FreezePanes = $true
        $ws
    }
    function Add-Chips($rng, [hashtable]$map) {   # 値ごとの色分け（条件付き書式）
        foreach ($e in $map.GetEnumerator()) {
            $fc = $rng.FormatConditions.Add(1, 3, "=""$($e.Key)""")
            if ($null -ne $e.Value.Fill) { $fc.Interior.Color = $e.Value.Fill }
            if ($null -ne $e.Value.Font) { $fc.Font.Color = $e.Value.Font }
            if ($e.Value.Bold) { $fc.Font.Bold = $true }
        }
    }
    function Add-DataBar($rng, [int]$barColor, $max = $null, [bool]$showValue = $true) {   # 単色のデータバー
        $db = $rng.FormatConditions.AddDatabar(); $db.BarColor.Color = $barColor; $db.MinPoint.Modify(0, 0)
        if ($null -ne $max) { $db.MaxPoint.Modify(0, $max) }
        $db.ShowValue = $showValue; try { $db.BarFillType = 0; $db.BarBorder.Type = 0 } catch {}
    }
    $Chip = @{
        Sev = @{ '重大' = @{ Fill = $Color.Burg; Font = $Color.White; Bold = 1 }; '警告' = @{ Fill = $Color.BurgTint; Font = $Color.Burg; Bold = 1 }; '注意' = @{ Fill = $Color.Chip; Font = $Color.Ink; Bold = 1 }; '情報' = @{ Font = $Color.Muted } }
        Pri = @{ '高' = @{ Fill = $Color.Navy; Font = $Color.White; Bold = 1 }; '中' = @{ Fill = $Color.NavyTint; Font = $Color.Navy; Bold = 1 }; '低' = @{ Font = $Color.Muted } }
        Cls = @{ 'ユニーク' = @{ Font = $Color.Navy; Bold = 1 }; '高' = @{ Font = $Color.Navy2; Bold = 1 }; '中' = @{ Font = $Color.Navy3 }; '低' = @{ Font = $Color.Burg; Bold = 1 }; '単一値' = @{ Font = $Color.Muted }; '全NULL' = @{ Font = $Color.Muted } }
        Skew = @{ '偏り' = @{ Font = $Color.Burg; Bold = 1 } }
    }
    # 列の定義: H 見出し / F 表示形式 / W 列幅 / A 配置(L・C・R) / Wrap 折返し / Font / Bar データバー色 / BarMax / NoVal 値を隠す / Chip 色分け / Pos 正の値を強調 / Val 数値に見える文字列 / Color 文字色
    function Add-Table($ws, [object[]]$spec, $rows, [string]$name, [int]$top = 4, [int]$left = 1) {
        $m = $spec.Count; $nr = [Math]::Max(1, $rows.Count); $body = $top + 1
        for ($c = 0; $c -lt $m; $c++) {   # 列の書式は書込み前に設定（文字列の数値・日付・数式への自動変換を防止）
            $s = $spec[$c]; $col = $ws.Cells.Item($body, $left + $c).Resize($nr, 1)
            $col.NumberFormat = if ($s.F) { $s.F } else { '@' }
            if ($s.A) { $col.HorizontalAlignment = @{ L = -4131; C = -4108; R = -4152 }[$s.A] }
            if ($s.Font) { $col.Font.Name = $s.Font }
            if ($s.Wrap) { $col.WrapText = $true }
            if ($s.Color) { $col.Font.Color = $s.Color }
        }
        Set-Head $ws ([string[]]($spec | ForEach-Object { $_.H })) $top $left
        for ($i = 0; $i -lt $rows.Count; $i += 10000) {   # 1 万行ずつ 2 次元配列で一括書込み
            $k = [Math]::Min(10000, $rows.Count - $i); $a = [object[,]]::new($k, $m)
            for ($y = 0; $y -lt $k; $y++) { $row = $rows[$i + $y]; for ($x = 0; $x -lt $m; $x++) { $v = $row[$x]; if ($v -is [psobject]) { $v = $v.psobject.BaseObject }; $a[$y, $x] = $v } }   # パイプライン経由の値は PSObject で包まれているため素の値に戻す
            $ws.Cells.Item($body + $i, $left).Resize($k, $m).Value2 = $a
        }
        $lo = $ws.ListObjects.Add(1, $ws.Cells.Item($top, $left).Resize($nr + 1, $m), $null, 1); $lo.TableStyle = 'Report'; if ($name) { $lo.Name = $name }
        for ($c = 0; $c -lt $m; $c++) {
            $s = $spec[$c]; $cr = $lo.ListColumns.Item($c + 1).DataBodyRange
            if ($s.Bar) { Add-DataBar $cr $s.Bar $s.BarMax (!$s.NoVal) }
            if ($s.Chip) { Add-Chips $cr $s.Chip }
            if ($s.Pos) { $fc = $cr.FormatConditions.Add(1, 5, '=0'); $fc.Font.Color = $s.Pos; $fc.Font.Bold = $true }   # 5 = xlGreater
            if ($s.Val) {   # 数値に見える文字列の緑三角を非表示（範囲一括不可・テーブル化後に設定）
                for ($y = 0; $y -lt $rows.Count -and $script:IgnoreBudget -gt 0; $y++) {
                    $v = $rows[$y][$c]; if ($v -is [string] -and $v -match '^[-+]?[0-9.]') { $ws.Cells.Item($body + $y, $left + $c).Errors.Item(3).Ignore = $true; $script:IgnoreBudget-- }
                }
            }
        }
        for ($c = 0; $c -lt $m; $c++) {   # 列幅: 指定値、なければ先頭 1000 行で自動調整（上限 60）
            $s = $spec[$c]; $colR = $ws.Columns.Item($left + $c)
            if ($s.W) { $colR.ColumnWidth = $s.W } else { [void]$ws.Cells.Item($top, $left + $c).Resize([Math]::Min($nr + 1, 1000), 1).Columns.AutoFit(); $colR.ColumnWidth = [Math]::Min(60, $colR.ColumnWidth + 2) }
        }
        $lo.DataBodyRange.VerticalAlignment = -4108
        if (@($spec | Where-Object { $_.Wrap }).Count -and $rows.Count -le 20000) {
            [void]$lo.DataBodyRange.Rows.AutoFit()   # 折返しの行は自動調整し、Meiryo UI で最終行が欠けないよう少し余白を足す
            if ($rows.Count -le 3000) { foreach ($row in $lo.DataBodyRange.Rows) { $row.RowHeight = [double]([Math]::Min(409.0, $row.RowHeight + 4)) } }
        } elseif ($rows.Count -le 200000) { $lo.DataBodyRange.RowHeight = 18 }
        $lo
    }
    function Set-Print($ws, [string]$titleRows = '$4:$4') {   # 印刷: 横向き・横 1 ページに収める（プリンター未設定の環境では省略）
        try {
            $xl.PrintCommunication = $false
            $ps = $ws.PageSetup; $ps.Orientation = 2; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
            if ($titleRows) { $ps.PrintTitleRows = $titleRows }
            $ps.CenterFooter = '&9&P / &N'; $ps.LeftMargin = 28; $ps.RightMargin = 28; $ps.TopMargin = 36; $ps.BottomMargin = 36
        } catch {} finally { try { $xl.PrintCommunication = $true } catch {} }
    }

    #── 図形・グラフの部品 ──────────────────────────────────────────
    function New-Canvas([string]$name, [int]$tab) {   # ダッシュボード用シート（薄いグレーの地に白いカードを並べる）
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $ws.Name = $name; $ws.Tab.Color = $tab
        $ws.Cells.Interior.Color = $Color.Canvas
        $w = $xl.ActiveWindow; $w.DisplayGridlines = $false; $w.DisplayHeadings = $false
        $ws
    }
    function Add-Text($ws, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [double]$size = 9, [int]$fg = 0x333333, [bool]$bold = $false, [string]$link = '', [int]$align = 1) {
        $tb = $ws.Shapes.AddTextbox(1, $x, $y, $w, $h); $tb.Fill.Visible = 0; $tb.Line.Visible = 0; $tb.Placement = 3
        $tf = $tb.TextFrame2; $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0; $tf.WordWrap = -1; $tf.AutoSize = 0
        if ($link) { $tb.DrawingObject.Formula = $link } else { $tf.TextRange.Text = $text }   # link: セルの値を表示（選択に応じて変わる）
        $f = $tf.TextRange.Font; $f.Size = $size; $f.Bold = $(if ($bold) { -1 } else { 0 }); $f.Fill.ForeColor.RGB = $fg
        $tf.TextRange.ParagraphFormat.Alignment = $align
        $tb
    }
    function Add-Card($ws, [double]$x, [double]$y, [double]$w, [double]$h, [string]$label, [string]$value, [string]$sub, [int]$accent, [string]$valueLink = '', [string]$subLink = '', [double]$valueSize = 20) {   # KPI カード
        $bg = $ws.Shapes.AddShape(1, $x, $y, $w, $h); $bg.Fill.ForeColor.RGB = $Color.White; $bg.Line.ForeColor.RGB = $Color.Hair; $bg.Line.Weight = 0.75; $bg.Shadow.Visible = 0
        $ac = $ws.Shapes.AddShape(1, $x, $y, 4, $h); $ac.Fill.ForeColor.RGB = $accent; $ac.Line.Visible = 0; $ac.Shadow.Visible = 0
        $t1 = Add-Text $ws ($x + 16) ($y + 9) ($w - 24) 14 $label 9 $Color.Ink2
        $t2 = Add-Text $ws ($x + 16) ($y + 24) ($w - 24) ($valueSize + 10) $value $valueSize $Color.Navy $true $valueLink
        $t3 = Add-Text $ws ($x + 16) ($y + $h - 19) ($w - 24) 13 $sub 8 $Color.Muted $false $subLink
        $g = $ws.Shapes.Range([object[]]@($bg.Name, $ac.Name, $t1.Name, $t2.Name, $t3.Name)).Group(); $g.Placement = 3
    }
    function Add-NavButton($ws, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [string]$sheet) {   # シートへのリンクボタン
        $b = $ws.Shapes.AddShape(1, $x, $y, $w, $h); $b.Fill.ForeColor.RGB = $Color.White; $b.Line.ForeColor.RGB = $Color.Rule; $b.Line.Weight = 0.75; $b.Shadow.Visible = 0; $b.Placement = 3
        $tf = $b.TextFrame2; $tf.MarginLeft = 2; $tf.MarginRight = 2; $tf.MarginTop = 0; $tf.MarginBottom = 0; $tf.VerticalAnchor = 3; $tf.WordWrap = 0
        $tf.TextRange.Text = $text; $tf.TextRange.Font.Size = 9; $tf.TextRange.Font.Fill.ForeColor.RGB = $Color.Navy; $tf.TextRange.ParagraphFormat.Alignment = 2
        [void]$ws.Hyperlinks.Add($b, '', "'$sheet'!A1", "「$sheet」シートへ移動")
    }
    function New-Chart($ws, [double]$x, [double]$y, [double]$w, [double]$h) { $co = $ws.ChartObjects().Add($x, $y, $w, $h); $co.Placement = 3; $co.Chart }
    function Add-Series($c, [string]$name, [string]$values, [string]$cats = '') {
        $s = $c.SeriesCollection().NewSeries(); $s.Name = $name; $s.Values = $values; if ($cats) { $s.XValues = $cats }; $s
    }
    function Format-Chart($c, [string]$title, [string]$sub = '', [string]$titleLink = '') {   # 白地・枠なし・左寄せのタイトル（2 行目は補足）
        $c.ChartArea.Format.Fill.Visible = -1; $c.ChartArea.Format.Fill.Solid(); $c.ChartArea.Format.Fill.ForeColor.RGB = $Color.White; $c.ChartArea.Format.Line.Visible = 0
        $c.ChartArea.Font.Size = 9; $c.ChartArea.Font.Color = $Color.Ink2
        $c.PlotArea.Format.Fill.Visible = 0; $c.PlotArea.Format.Line.Visible = 0
        $c.HasTitle = $true
        if ($titleLink) {   # セルに連動するタイトル: TextFrame2 で書式を変えるとリンクが切れるため従来の Font で設定
            $c.ChartTitle.Formula = $titleLink; $c.ChartTitle.Font.Size = 11; $c.ChartTitle.Font.Bold = $true; $c.ChartTitle.Font.Color = $Color.Ink
        } else {
            $c.ChartTitle.Text = $(if ($sub) { "$title`n$sub" } else { $title })
            $tr = $c.ChartTitle.Format.TextFrame2.TextRange
            $tr.Font.Size = 11; $tr.Font.Bold = -1; $tr.Font.Fill.ForeColor.RGB = $Color.Ink; $tr.ParagraphFormat.Alignment = 1
            if ($sub) { $sr = $tr.Characters($title.Length + 2, $sub.Length); $sr.Font.Size = 8.5; $sr.Font.Bold = 0; $sr.Font.Fill.ForeColor.RGB = $Color.Muted }
        }
        $c.ChartTitle.Left = 10; $c.ChartTitle.Top = 6
    }
    function Format-Axis($ax, [string]$fmt = '', [bool]$hide = $false, [bool]$grid = $false, [bool]$line = $false, [bool]$log = $false) {   # 目盛は細く控えめに
        $ax.MajorTickMark = -4142; $ax.MinorTickMark = -4142
        if ($line) { $ax.Format.Line.Visible = -1; $ax.Format.Line.ForeColor.RGB = $Color.Rule; $ax.Format.Line.Weight = 0.75 } else { $ax.Format.Line.Visible = 0 }
        if ($hide) { $ax.TickLabelPosition = -4142 } else { $ax.TickLabels.Font.Size = 8.5; $ax.TickLabels.Font.Color = $Color.Ink2; if ($fmt) { $ax.TickLabels.NumberFormat = $fmt } }
        $ax.HasMajorGridlines = $grid; $ax.HasMinorGridlines = $false
        if ($grid) { $gl = $ax.MajorGridlines.Format.Line; $gl.ForeColor.RGB = $Color.Hair; $gl.Weight = 0.75; $gl.DashStyle = 1 }
        if ($log) { $ax.ScaleType = -4133; $ax.LogBase = 10 }
    }
    function Set-Fill($s, [int]$fill, [bool]$gap = $false) {   # gap: 積み上げの区切りに白い細線
        $s.Format.Fill.Visible = -1; $s.Format.Fill.Solid(); $s.Format.Fill.ForeColor.RGB = $fill
        if ($gap) { $s.Format.Line.Visible = -1; $s.Format.Line.ForeColor.RGB = $Color.White; $s.Format.Line.Weight = 1.5 } else { $s.Format.Line.Visible = 0 }
    }
    function Set-Labels($s, [string]$fmt, [int]$pos = 2) {   # 値ラベル（2 = 外側の端）
        $s.HasDataLabels = $true; $dl = $s.DataLabels(); $dl.NumberFormat = $fmt; $dl.Font.Size = 8.5; $dl.Font.Color = $Color.Ink2; try { $dl.Position = $pos } catch {}
    }
    function Set-Legend($c, [int]$pos = -4107) { $c.HasLegend = $true; $c.Legend.Position = $pos; $c.Legend.Font.Size = 8.5; $c.Legend.Font.Color = $Color.Ink2 }
    function Set-Gap($c, [int]$count, [double]$plotH) {   # 棒の太さを最大 16pt 程度に抑える
        $slot = $plotH / [Math]::Max(1, $count); $thick = [Math]::Min(16.0, $slot * 0.62)
        $c.ChartGroups(1).GapWidth = [int][Math]::Max(40, [Math]::Min(400, ($slot / $thick - 1) * 100))
    }
    function Add-Empty($ws, [double]$x, [double]$y, [double]$w, [double]$h, [string]$title, [string]$msg) {   # データがないときのカード
        $bg = $ws.Shapes.AddShape(1, $x, $y, $w, $h); $bg.Fill.ForeColor.RGB = $Color.White; $bg.Line.Visible = 0; $bg.Shadow.Visible = 0; $bg.Placement = 3
        [void](Add-Text $ws ($x + 10) ($y + 8) ($w - 20) 18 $title 11 $Color.Ink $true)
        [void](Add-Text $ws ($x + 10) ($y + $h / 2 - 8) ($w - 20) 16 $msg 9 $Color.Muted $false '' 2)
    }

    #── データのシート ──────────────────────────────────────────────
    Write-Host '    テーブル一覧・データ分布'
    $ws = New-Sheet 'テーブル一覧' 'テーブル一覧' "対象: $CsvDir ／ テーブル数: $($Tables.Count) ／ 総件数: $(Format-N $total) 件 ／ 作成: $(Get-Date -f 'yyyy/MM/dd HH:mm')"
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    $loSum = Add-Table $ws @(
        @{ H = 'テーブル'; W = 26 }, @{ H = '件数'; F = '#,##0'; W = 12 }, @{ H = '件数(対数)'; F = '0.0'; W = 16; Bar = $Color.Navy; BarMax = $logMax; NoVal = 1 },
        @{ H = '項目数'; F = '#,##0'; W = 8 }, @{ H = 'キー項目'; W = 18 }, @{ H = 'キーユニーク数'; F = '#,##0'; W = 11 }, @{ H = '主キー候補'; W = 24 },
        @{ H = '提案'; F = '#,##0;;'; W = 7; A = 'C'; Color = $Color.Navy }, @{ H = '重大'; F = '#,##0;;'; W = 7; A = 'C'; Pos = $Color.Burg },
        @{ H = '警告'; F = '#,##0;;'; W = 7; A = 'C'; Pos = $Color.Burg }, @{ H = '注意'; F = '#,##0;;'; W = 7; A = 'C' }, @{ H = '主な所見（重大・警告）'; W = 56; Color = $Color.Burg }
    ) $Sum 'T_SUM'
    Set-Print $ws

    $ws = New-Sheet 'データ分布' '項目別データ分布' '空値は NULL として集計。推定型は値の形式から判定。カーディナリティ: ユニーク ＞ 高（ユニーク数 100 以上）＞ 中（10～99）＞ 低（2～9）' $Color.Navy 2
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    $DistHead = 'テーブル', '項目名', '推定型', '非NULL件数', 'NULL件数', 'NULL率', 'ユニーク数', 'ユニーク数(対数)', 'カーディナリティ', 'ユニーク率', '1値平均件数', '最頻値', '最頻値比率', '最小値', '最大値', '平均', '最大長', '最大バイト', '上位値(件数)', 'インデックス', '注意'
    [void](Add-Table $ws @(
            @{ H = 'テーブル'; W = 22 }, @{ H = '項目名'; W = 20 }, @{ H = '推定型'; W = 12; A = 'C' }, @{ H = '非NULL件数'; F = '#,##0'; W = 11 }, @{ H = 'NULL件数'; F = '#,##0'; W = 10 },
            @{ H = 'NULL率'; F = '0.0%'; W = 10; Bar = $Color.Gray; BarMax = 1 }, @{ H = 'ユニーク数'; F = '#,##0'; W = 11 }, @{ H = 'ユニーク数(対数)'; F = '0.0'; W = 14; Bar = $Color.Navy3; BarMax = $logMax; NoVal = 1 },
            @{ H = 'カーディナリティ'; W = 13; A = 'C'; Chip = $Chip.Cls }, @{ H = 'ユニーク率'; F = '0.0%'; W = 11 }, @{ H = '1値平均件数'; F = '#,##0.0'; W = 11 },
            @{ H = '最頻値'; W = 18; Val = 1; A = 'L' }, @{ H = '最頻値比率'; F = '0.0%'; W = 11; Bar = $Color.Burg3; BarMax = 1 }, @{ H = '最小値'; W = 18; Val = 1; A = 'L' }, @{ H = '最大値'; W = 18; Val = 1; A = 'L' },
            @{ H = '平均'; F = '#,##0.00'; W = 12 }, @{ H = '最大長'; F = '#,##0'; W = 8 }, @{ H = '最大バイト'; F = '#,##0'; W = 9 }, @{ H = '上位値(件数)'; W = 40 },
            @{ H = 'インデックス'; W = 20; Color = $Color.Navy }, @{ H = '注意'; W = 40; Color = $Color.Burg }
        ) $Dist 'T_DIST')
    Set-Print $ws

    Write-Host '    インデックス提案・検索リスク警告'
    $ws = New-Sheet 'インデックス提案' 'インデックス提案' 'CSV のデータから推定した候補。既存のインデックス・実際の SQL は考慮していないため、作成前に EXPLAIN で効果を確認してください（判定条件は「判定基準」シート）'
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    [void](Add-Table $ws @(
            @{ H = '優先度'; W = 8; A = 'C'; Chip = $Chip.Pri }, @{ H = 'テーブル'; W = 20 }, @{ H = '件数'; F = '#,##0'; W = 11 }, @{ H = '種別'; W = 18 }, @{ H = '対象列'; W = 22 },
            @{ H = '根拠'; W = 58; Wrap = 1 }, @{ H = '想定する検索'; W = 36; Wrap = 1 }, @{ H = 'DDL'; W = 64; Wrap = 1; Font = 'Consolas'; Color = $Color.Navy }
        ) $IdxRows 'T_IDX')
    Set-Print $ws

    $ws = New-Sheet '検索リスク警告' '検索リスク警告' '検索で問題になりやすいデータ分布・値の形式。重要度: 重大 ＞ 警告 ＞ 注意 ＞ 情報（判定条件は「判定基準」シート）' $(if (@($Warns | Where-Object { $_.Sev -in '重大', '警告' }).Count) { $Color.Burg } else { $Color.Navy })
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    [void](Add-Table $ws @(
            @{ H = '重要度'; W = 8; A = 'C'; Chip = $Chip.Sev }, @{ H = 'テーブル'; W = 20 }, @{ H = '項目名'; W = 20 }, @{ H = '分類'; W = 16 },
            @{ H = '内容'; W = 60; Wrap = 1 }, @{ H = '検索への影響'; W = 52; Wrap = 1 }, @{ H = '対策'; W = 40; Wrap = 1 }
        ) $WarnRows 'T_WARN')
    Set-Print $ws

    Write-Host '    値の分布・キー先頭1桁'
    $ws = New-Sheet '値の分布' '項目別 上位値の分布' "各項目の上位 $($TH.TopN) 値と構成比（NULL を含む全件に対する比率）。判定「偏り」= 最頻値が非 NULL 件数の 50% 以上" $Color.Navy 2
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    [void](Add-Table $ws @(
            @{ H = 'テーブル'; W = 22 }, @{ H = '項目名'; W = 20 }, @{ H = '順位'; F = '#,##0'; W = 6; A = 'C' }, @{ H = '値'; W = 40; Val = 1; A = 'L' },
            @{ H = '件数'; F = '#,##0'; W = 11 }, @{ H = '構成比'; F = '0.0%'; W = 16; Bar = $Color.Navy3; BarMax = 1 }, @{ H = '累積比'; F = '0.0%'; W = 9 }, @{ H = '判定'; W = 7; A = 'C'; Chip = $Chip.Skew }
        ) $Top 'T_TOP')

    $ws = New-Sheet 'キー先頭1桁' 'キー項目 先頭1桁カテゴリ別件数' 'キー項目は各テーブルの先頭列（-KeyColumn で変更可）'
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    [void](Add-Table $ws @(
            @{ H = 'テーブル'; W = 24 }, @{ H = 'キー項目'; W = 20 }, @{ H = '先頭1桁'; W = 9; A = 'C'; Val = 1 }, @{ H = '件数'; F = '#,##0'; W = 12 }, @{ H = '構成比'; F = '0.0%'; W = 18; Bar = $Color.Navy3; BarMax = 1 }
        ) $Pre 'T_PRE')

    Write-Host '    全データ'
    $wsS = New-Sheet '全データ検索' '全データ検索' 'C3 にキーワードを入力すると、全テーブル・全項目を部分一致で検索します（大文字小文字は区別なし／ワイルドカード * ? 可）' $Color.Navy3
    $ws = New-Sheet '全データ' '全データ（検索用）' $(if ($All.Count -lt $total) { "※ Excel の行数上限により先頭 $(Format-N $LIMIT) 件のみ格納しています" } else { '1 レコード ＝「列名=値 | 列名=値 …」' }) $Color.Navy3
    [void](Add-Table $ws @(@{ H = 'テーブル'; W = 28 }, @{ H = '行No'; F = '#,##0'; W = 9 }, @{ H = '内容'; W = 150 }) $All 'T_ALL')
    if ($All.Count -lt $total) { $ws.Range('A2').Font.Color = $Color.Burg }

    # 検索シート: キーワード → FILTER 関数で全データから抽出（動的配列: Microsoft 365 / Excel 2021 以降）
    $ws = $wsS; $ws.Range('B3').Value2 = '検索語'; $ws.Range('B3').Font.Bold = $true; $ws.Range('B3').HorizontalAlignment = -4152
    $in = $ws.Range('C3'); $in.NumberFormat = '@'; $in.Interior.Color = $Color.White; $in.BorderAround(1, -4138, [Type]::Missing, $Color.Navy) | Out-Null; $in.Font.Bold = $true; $ws.Rows.Item(3).RowHeight = 22
    $ws.Range('A3').Formula2 = '=IF(C3="","",COUNTIF(T_ALL[内容],"*"&C3&"*")&" 件")'; $ws.Range('A3').Font.Color = $Color.Navy; $ws.Range('A3').Font.Bold = $true
    Set-Head $ws ([string[]]('テーブル', '行No', '内容'))
    $ws.Range('A5').Formula2 = '=IF(C3="","",FILTER(T_ALL,ISNUMBER(SEARCH(C3,T_ALL[内容])),"該当なし"))'
    $ws.Range('A5:C1048576').FormatConditions.Add(2, [Type]::Missing, '=AND(INDEX($A:$A,ROW())<>"",MOD(ROW(),2)=1)').Interior.Color = $Color.Stripe
    $ws.Range('B:B').NumberFormat = '#,##0'; $ws.Range('A:A').ColumnWidth = 28; $ws.Range('B:B').ColumnWidth = 9; $ws.Range('C:C').ColumnWidth = 150
    Add-Link $ws 'E1' '◀ ダッシュボード' 'ダッシュボード'

    $ws = New-Sheet '判定基準' '判定基準' 'カーディナリティ区分・検索リスク警告・インデックス提案の判定条件と、レポートの前提' $Color.Navy3
    Add-Link $ws 'A3' '◀ ダッシュボード' 'ダッシュボード'
    [void](Add-Table $ws @(
            @{ H = '区分'; W = 16 }, @{ H = '名称'; W = 20; Color = $Color.Navy }, @{ H = '重要度'; W = 8; A = 'C'; Chip = $Chip.Sev }, @{ H = '判定条件'; W = 52; Wrap = 1 },
            @{ H = '検索への影響・ねらい'; W = 56; Wrap = 1 }, @{ H = '対策'; W = 44; Wrap = 1 }
        ) $RuleRows 'T_RULE')
    Set-Print $ws

    #── グラフ用データ（非表示シート）─────────────────────────────
    Write-Host '    グラフ'
    $GD = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $GD.Name = 'グラフデータ'; $GD.Tab.Color = $Color.Gray
    $GN = "'グラフデータ'!"
    function Write-Block($row, $col, [string[]]$head, $rows) {   # 見出し + 2 次元配列で一括書込み
        for ($c = 0; $c -lt $head.Count; $c++) { $GD.Cells.Item($row, $col + $c).Value2 = [string]$head[$c] }
        if (!$rows.Count) { return }
        $a = [object[,]]::new($rows.Count, $head.Count)
        for ($y = 0; $y -lt $rows.Count; $y++) { for ($x = 0; $x -lt $head.Count; $x++) { $v = $rows[$y][$x]; if ($v -is [psobject]) { $v = $v.psobject.BaseObject }; $a[$y, $x] = $v } }
        $GD.Cells.Item($row + 1, $col).Resize($rows.Count, $head.Count).Value2 = $a
    }
    function GRange([int]$row, [int]$col, [int]$count) { '=' + $GN + $GD.Cells.Item($row, $col).Resize([Math]::Max(1, $count), 1).Address() }

    # 値の分布（区間）: テーブル分析のグラフ用
    [void](Add-Table $GD @(@{ H = 'テーブル' }, @{ H = '項目名' }, @{ H = '単位' }, @{ H = 'No'; F = '0' }, @{ H = '区間' }, @{ H = '件数'; F = '#,##0' }, @{ H = '構成比'; F = '0.0%' }) $Hist 'T_HIST' 4 79)

    # ダッシュボードの元データ
    $byRows = @($Tables | Sort-Object @{ Expression = { $_.Rows }; Descending = $true }, @{ Expression = { $_.Name } })
    $NT = $byRows.Count; $topN = [Math]::Min(15, [Math]::Ceiling($NT / 2)); $botN = [Math]::Min(15, $NT - $topN)
    Write-Block 1 1 @('テーブル', '件数') @($byRows | Select-Object -First $topN | ForEach-Object { , @($_.Name, $_.Rows) })
    Write-Block 1 4 @('テーブル', '件数') @($byRows | Select-Object -Last $botN | Sort-Object @{ Expression = { $_.Rows } }, @{ Expression = { $_.Name } } | ForEach-Object { , @($_.Name, $_.Rows) })
    $bands = @(@('0 件', 0, 0), @('1～99', 1, 99), @('100～999', 100, 999), @('1千～1万', 1000, 9999), @('1万～10万', 10000, 99999), @('10万～100万', 100000, 999999), @('100万～', 1000000, [double]::MaxValue))
    Write-Block 1 7 @('規模', 'テーブル数') @($bands | ForEach-Object { $b = $_; , @($b[0], @($Tables | Where-Object { $_.Rows -ge $b[1] -and $_.Rows -le $b[2] }).Count) })
    $clsCount = @($Classes | ForEach-Object { $cl = $_; , @($cl, @($Profiles | Where-Object { $_.Class -eq $cl }).Count) })
    Write-Block 1 10 @('区分', '項目数') $clsCount
    $pts = @($Profiles | Where-Object { $_.NonNull -gt 0 -and $_.Distinct -gt 0 })
    $ptHi = @($pts | Where-Object { $_.Distinct -ge $TH.LowCard } | ForEach-Object { , @($_.NonNull, $_.Distinct) })
    $ptLo = @($pts | Where-Object { $_.Distinct -lt $TH.LowCard } | ForEach-Object { , @($_.NonNull, $_.Distinct) })
    $axMax = [Math]::Pow(10, [Math]::Max(1, [Math]::Ceiling([Math]::Log10([Math]::Max(10, $maxRows)))))
    Write-Block 1 13 @('非NULL件数', 'ユニーク数') $ptHi
    Write-Block 1 15 @('非NULL件数', 'ユニーク数') $ptLo
    Write-Block 1 17 @('x', 'y') @(@(1, 1), @($axMax, $axMax))                   # 一意の対角線
    Write-Block 1 19 @('x', 'y') @(@(1, $TH.LowCard), @($axMax, $TH.LowCard))   # 低カーディナリティの境界
    $wCats = @($Warns | Group-Object Cat | Sort-Object { - $_.Count }, Name)
    Write-Block 1 22 @('分類', '重大', '警告', '注意', '情報') @($wCats | ForEach-Object { $grp = $_.Group; , @($_.Name, @($grp | Where-Object Sev -eq '重大').Count, @($grp | Where-Object Sev -eq '警告').Count, @($grp | Where-Object Sev -eq '注意').Count, @($grp | Where-Object Sev -eq '情報').Count) })
    $iKinds = @($Idx | Group-Object Kind | Sort-Object { - $_.Count }, Name)
    Write-Block 1 28 @('種別', '高', '中', '低') @($iKinds | ForEach-Object { $grp = $_.Group; , @($_.Name, @($grp | Where-Object Pri -eq '高').Count, @($grp | Where-Object Pri -eq '中').Count, @($grp | Where-Object Pri -eq '低').Count) })

    #── テーブル分析（選択に応じて変わるグラフと項目一覧）──────────
    Write-Host '    テーブル分析'
    $An = New-Canvas 'テーブル分析' $Color.Navy
    $maxCols = ($Tables | ForEach-Object { $_.Header.Count } | Measure-Object -Maximum).Maximum
    $widths = 2, 22, 11, 9, 11, 12, 18, 10, 22, 46, 2
    for ($i = 0; $i -lt $widths.Count; $i++) { $An.Columns.Item($i + 1).ColumnWidth = $widths[$i] }
    foreach ($e in @(@(1, 8), @(2, 30), @(3, 16), @(4, 10), @(5, 24), @(6, 10))) { $An.Rows.Item($e[0]).RowHeight = $e[1] }
    $An.Range('B2').Value2 = 'テーブル分析'; $f = $An.Range('B2').Font; $f.Size = 18; $f.Bold = $true; $f.Color = $Color.Navy
    $An.Range('B3').Value2 = 'C5 でテーブル、G5 で項目を選ぶと（セルを選択して ▼）、グラフと項目一覧が切り替わります'; $An.Range('B3').Font.Size = 9; $An.Range('B3').Font.Color = $Color.Ink2
    foreach ($e in @(@('B5', 'テーブル'), @('F5', '項目'))) { $r = $An.Range($e[0]); $r.Value2 = [string]$e[1]; $r.HorizontalAlignment = -4152; $r.Font.Color = $Color.Ink2; $r.IndentLevel = 1 }
    $defT = $byRows[0]
    $defC = ($defT.Profiles | Where-Object { $_.Skew } | Select-Object -First 1); if (!$defC) { $defC = $defT.Profiles[0] }
    foreach ($e in @(@('C5:E5', $defT.Name), @('G5:H5', $defC.Name))) {
        $r = $An.Range($e[0]); $r.Merge(); $r.NumberFormat = '@'; $r.Interior.Color = $Color.White; $r.Font.Bold = $true; $r.Font.Size = 11; $r.Font.Color = $Color.Navy; $r.IndentLevel = 1
        $r.BorderAround(1, 2, [Type]::Missing, $Color.Navy) | Out-Null; $An.Range($e[0].Split(':')[0]).Value2 = [string]$e[1]
    }

    # 補助セル（グラフデータ BA 列以降）: 選択値・KPI・グラフの系列（FILTER で抽出）
    $GD.Range('BA1').Formula2 = "='テーブル分析'!`$C`$5"
    $GD.Range('BA2').Formula2 = '=SUM(--(T_DIST[テーブル]=$BA$1))'
    $GD.Range('BA3').Formula2 = "=LET(sel,'テーブル分析'!`$G`$5,IF(SUM((T_DIST[テーブル]=`$BA`$1)*(T_DIST[項目名]=sel))>0,sel,IFERROR(INDEX(FILTER(T_DIST[項目名],T_DIST[テーブル]=`$BA`$1),1),`"`")))"
    $GD.Range('BA4').Formula2 = '=SUM((T_TOP[テーブル]=$BA$1)*(T_TOP[項目名]=$BA$3))'
    $GD.Range('BA5').Formula2 = '=SUM((T_HIST[テーブル]=$BA$1)*(T_HIST[項目名]=$BA$3))'
    $GD.Range('BA6').Formula2 = '="上位値の構成比 ─ "&$BA$3'
    $GD.Range('BA7').Formula2 = '="値の分布 ─ "&$BA$3&IFERROR("（"&INDEX(FILTER(T_HIST[単位],(T_HIST[テーブル]=$BA$1)*(T_HIST[項目名]=$BA$3)),1)&"）","")'
    $GD.Range('BA8').Formula2 = '="項目別カーディナリティ ─ "&$BA$1'
    $kpi = [ordered]@{
        BA10 = @('=IFERROR(XLOOKUP($BA$1,T_SUM[テーブル],T_SUM[件数]),0)', '#,##0')
        BA11 = @('="件数の順位 "&(SUM(--(T_SUM[件数]>$BA$10))+1)&" ／ "&ROWS(T_SUM[件数])&" テーブル"', 'General')
        BA12 = @('=$BA$2', '#,##0')
        BA13 = @('="ユニーク "&SUM((T_DIST[テーブル]=$BA$1)*(T_DIST[カーディナリティ]="ユニーク"))&" ／ 低 "&SUM((T_DIST[テーブル]=$BA$1)*(T_DIST[カーディナリティ]="低"))&" 列"', 'General')
        BA14 = @('=IFERROR(XLOOKUP($BA$1,T_SUM[テーブル],T_SUM[主キー候補]),"")', 'General')
        BA15 = @('=IF($BA$14="（なし）","サロゲートキーの追加を推奨",IF($BA$14="","","一意・NULL なし"))', 'General')
        BA16 = @('=SUM(--(T_IDX[テーブル]=$BA$1))', '#,##0')
        BA17 = @('="優先度 高 "&SUM((T_IDX[テーブル]=$BA$1)*(T_IDX[優先度]="高"))&" ／ 中 "&SUM((T_IDX[テーブル]=$BA$1)*(T_IDX[優先度]="中"))', 'General')
        BA18 = @('=SUM((T_WARN[テーブル]=$BA$1)*((T_WARN[重要度]="重大")+(T_WARN[重要度]="警告")))', '#,##0')
        BA19 = @('="重大 "&SUM((T_WARN[テーブル]=$BA$1)*(T_WARN[重要度]="重大"))&" ／ 警告 "&SUM((T_WARN[テーブル]=$BA$1)*(T_WARN[重要度]="警告"))&" ／ 注意 "&SUM((T_WARN[テーブル]=$BA$1)*(T_WARN[重要度]="注意"))', 'General')
    }
    foreach ($e in $kpi.GetEnumerator()) { if ($e.Value[1] -ne 'General') { $GD.Range($e.Key).NumberFormat = $e.Value[1] }; $GD.Range($e.Key).Formula2 = $e.Value[0] }   # 文字列を返す数式は既定の書式のまま（文字列書式だと数式が文字列として入る）
    foreach ($e in @(@('BA20', '項目名'), @('BB20', 'ユニーク'), @('BC20', '高'), @('BD20', '中'), @('BE20', '低'), @('BG20', '値'), @('BH20', '構成比'), @('BI20', '偏り'), @('BK20', '区間'), @('BL20', '件数'))) { $GD.Range($e[0]).Value2 = [string]$e[1] }
    $GD.Range('BA21').Formula2 = '=FILTER(T_DIST[項目名],T_DIST[テーブル]=$BA$1,"")'
    # 区分ごとに系列を分け、該当しない項目は #N/A（棒を描かない）→ 1 本の棒が区分の色になる
    foreach ($col in 'BB', 'BC', 'BD', 'BE') { $GD.Range("$($col)21").Formula2 = "=LET(flt,T_DIST[テーブル]=`$BA`$1,uq,FILTER(T_DIST[ユニーク数],flt,0),cls,FILTER(T_DIST[カーディナリティ],flt,`"`"),IF((cls=$col`$20)*(uq>0),uq,NA()))" }
    $GD.Range('BG21').Formula2 = '=FILTER(T_TOP[値],(T_TOP[テーブル]=$BA$1)*(T_TOP[項目名]=$BA$3),"")'
    $GD.Range('BH21').Formula2 = '=LET(flt,(T_TOP[テーブル]=$BA$1)*(T_TOP[項目名]=$BA$3),sh,FILTER(T_TOP[構成比],flt,0),jd,FILTER(T_TOP[判定],flt,""),IF(jd="偏り",0,sh))'
    $GD.Range('BI21').Formula2 = '=LET(flt,(T_TOP[テーブル]=$BA$1)*(T_TOP[項目名]=$BA$3),sh,FILTER(T_TOP[構成比],flt,0),jd,FILTER(T_TOP[判定],flt,""),IF(jd="偏り",sh,0))'
    $GD.Range('BK21').Formula2 = '=FILTER(T_HIST[区間],(T_HIST[テーブル]=$BA$1)*(T_HIST[項目名]=$BA$3),"")'
    $GD.Range('BL21').Formula2 = '=FILTER(T_HIST[件数],(T_HIST[テーブル]=$BA$1)*(T_HIST[項目名]=$BA$3),0)'
    # 可変長の範囲名（グラフの系列・入力規則が参照）
    # 横棒グラフは行数を一定以上に揃え（空の行は棒なし）、項目の少ないテーブルでも棒の太さを変えない
    $cs = @($Tables | ForEach-Object { $_.Header.Count } | Sort-Object)
    $slots = [Math]::Min(40, [Math]::Max(12, $cs[[int][Math]::Floor(0.9 * ($cs.Count - 1))]))   # 列数の 90 パーセンタイル
    foreach ($e in @(@('pcCat', 'BA', '$BA$2', $slots), @('pcU', 'BB', '$BA$2', $slots), @('pcH', 'BC', '$BA$2', $slots), @('pcM', 'BD', '$BA$2', $slots), @('pcL', 'BE', '$BA$2', $slots),
            @('tvCat', 'BG', '$BA$4', 12), @('tvN', 'BH', '$BA$4', 12), @('tvD', 'BI', '$BA$4', 12), @('hsCat', 'BK', '$BA$5', 1), @('hsN', 'BL', '$BA$5', 1))) {
        [void]$GD.Names.Add($e[0], "=OFFSET($GN`$$($e[1])`$21,0,0,MAX($($e[3]),$GN$($e[2])),1)")   # シート単位の名前（ブック名に依存しない）
    }
    [void]$wb.Names.Add('nmTables', "='テーブル一覧'!" + $loSum.ListColumns.Item(1).DataBodyRange.Address())
    [void]$wb.Names.Add('nmCols', "=OFFSET($GN`$BA`$21,0,0,MAX(1,$GN`$BA`$2),1)")
    $v = $An.Range('C5:E5').Validation; $v.Delete(); $v.Add(3, 1, 1, '=nmTables'); $v.IgnoreBlank = $true
    $v = $An.Range('G5:H5').Validation; $v.Delete(); $v.Add(3, 1, 1, '=nmCols'); $v.IgnoreBlank = $true
    $An.Range('I5').Formula2 = "=IF(`$G`$5=$GN`$BA`$3,`"`",IF(`$G`$5=`"`",`"先頭の項目`",`"選択した項目はこのテーブルにないため`")&`"「`"&$GN`$BA`$3&`"」を表示中`")"
    $An.Range('I5').Font.Size = 8.5; $An.Range('I5').Font.Color = $Color.Muted

    # KPI カード
    $x0 = $An.Range('B1').Left; $PW = $An.Range('B1:J1').Width; $gap = 12
    $y = $An.Rows.Item(7).Top; $cardW = ($PW - 4 * $gap) / 5; $ch = 72
    $cards = @(
        @('件数', '$BA$10', '$BA$11', $Color.Navy, 20), @('項目数', '$BA$12', '$BA$13', $Color.Navy, 20), @('主キー候補', '$BA$14', '$BA$15', $Color.Navy, 12),
        @('インデックス提案', '$BA$16', '$BA$17', $Color.Navy, 20), @('警告（重大＋警告）', '$BA$18', '$BA$19', $Color.Burg, 20))
    for ($i = 0; $i -lt 5; $i++) { $cd = $cards[$i]; Add-Card $An ($x0 + $i * ($cardW + $gap)) $y $cardW $ch $cd[0] '' '' $cd[3] "=$GN$($cd[1])" "=$GN$($cd[2])" $cd[4] }

    # グラフ: 項目別カーディナリティ（区分で色分け・対数目盛）／上位値の構成比／値の分布
    $y += $ch + 14; $LW = [Math]::Round($PW * 0.5 - $gap / 2); $RW = $PW - $LW - $gap
    $H1 = [Math]::Max(460, $slots * 14 + 110); $H2 = ($H1 - $gap) / 2
    $c = New-Chart $An $x0 $y $LW $H1
    $clsColor = @($Color.Navy, $Color.Navy2, $Color.Navy3, $Color.Burg)
    $sn = @('pcU', 'pcH', 'pcM', 'pcL'); $sl = @('ユニーク', '高', '中', '低')
    for ($i = 0; $i -lt 4; $i++) { [void](Add-Series $c $sl[$i] "=$GN$($sn[$i])" "=$($GN)pcCat") }
    $c.ChartType = 57   # 種類の変更後に系列を取り直す（変更前の参照は無効になることがある）
    for ($i = 0; $i -lt 4; $i++) { $sc = $c.SeriesCollection($i + 1); Set-Fill $sc $clsColor[$i] }
    Format-Chart $c '' '' "=$GN`$BA`$8"
    $c.ChartGroups(1).Overlap = 100; Set-Gap $c $slots ($H1 - 110)
    $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.ReversePlotOrder = $true; $ax.Crosses = 2; $ax.TickLabels.Font.Size = 8
    $vx = $c.Axes(2); Format-Axis $vx '#,##0' $false $true $false $true; $vx.MinimumScale = [double]1; $vx.HasTitle = $true; $vx.AxisTitle.Text = 'ユニーク数（対数目盛）'; $vx.AxisTitle.Font.Size = 8; $vx.AxisTitle.Font.Bold = $false; $vx.AxisTitle.Font.Color = $Color.Muted
    Set-Legend $c

    $c = New-Chart $An ($x0 + $LW + $gap) $y $RW $H2
    $s1 = Add-Series $c '構成比' "=$($GN)tvN" "=$($GN)tvCat"; $s2 = Add-Series $c '偏り（最頻値が 50% 以上）' "=$($GN)tvD" "=$($GN)tvCat"
    $c.ChartType = 57; Format-Chart $c '' '' "=$GN`$BA`$6"
    Set-Fill $s1 $Color.Navy3; Set-Fill $s2 $Color.Burg; $c.ChartGroups(1).Overlap = 100; Set-Gap $c 12 ($H2 - 70)
    Set-Labels $s1 '0.0%;;;'; Set-Labels $s2 '0.0%;;;'
    $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.ReversePlotOrder = $true; $ax.TickLabels.Font.Size = 8
    $vx = $c.Axes(2); Format-Axis $vx '' $true; $vx.MinimumScale = [double]0
    $c.HasLegend = $false

    $c = New-Chart $An ($x0 + $LW + $gap) ($y + $H2 + $gap) $RW $H2
    $s = Add-Series $c '件数' "=$($GN)hsN" "=$($GN)hsCat"
    $c.ChartType = 51; Format-Chart $c '' '' "=$GN`$BA`$7"
    Set-Fill $s $Color.Navy; $c.ChartGroups(1).GapWidth = 30
    $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.TickLabels.Font.Size = 8
    $vx = $c.Axes(2); Format-Axis $vx '#,##0' $false $true; $vx.MinimumScale = [double]0
    $c.HasLegend = $false

    # 項目一覧（FILTER で選択テーブルの項目を抽出）
    $yl = $y + $H1 + 22; $rt = 7; while ($An.Rows.Item($rt).Top -lt $yl) { $rt++ }
    $An.Range("$($rt):1048576").Interior.ColorIndex = -4142
    $An.Rows.Item($rt).RowHeight = 24; $An.Cells.Item($rt, 2).Formula2 = '="項目一覧 ─ "&$C$5'; $f = $An.Cells.Item($rt, 2).Font; $f.Size = 12; $f.Bold = $true; $f.Color = $Color.Navy
    $listHead = '項目名', '推定型', 'NULL率', 'ユニーク数', 'カーディナリティ', '最頻値', '最頻値比率', 'インデックス', '注意'
    Set-Head $An ([string[]]$listHead) ($rt + 1) 2
    $pick = ($listHead | ForEach-Object { [Array]::IndexOf($DistHead, $_) + 1 }) -join ','
    $rd = $rt + 2; $rl = $rd + [Math]::Max(1, $maxCols) - 1
    $An.Cells.Item($rd, 2).Formula2 = "=IFERROR(CHOOSECOLS(FILTER(T_DIST,T_DIST[テーブル]=`$C`$5),$pick),`"`")"
    $body = $An.Range("B$($rd):J$($rl)"); $body.RowHeight = 18; $body.VerticalAlignment = -4108
    $An.Range("D$($rd):D$($rl)").NumberFormat = '0.0%'; $An.Range("E$($rd):E$($rl)").NumberFormat = '#,##0'
    $An.Range("C$($rd):C$($rl)").HorizontalAlignment = -4108; $An.Range("F$($rd):F$($rl)").HorizontalAlignment = -4108; $An.Range("G$($rd):G$($rl)").HorizontalAlignment = -4131
    $An.Range("I$($rd):I$($rl)").Font.Color = $Color.Navy; $An.Range("J$($rd):J$($rl)").Font.Color = $Color.Burg
    $fc = $body.FormatConditions.Add(2, [Type]::Missing, '=AND(INDEX($B:$B,ROW())<>"",MOD(ROW(),2)=0)'); $fc.Interior.Color = $Color.Stripe; $fc.StopIfTrue = $false
    $fc = $body.FormatConditions.Add(2, [Type]::Missing, '=INDEX($B:$B,ROW())<>""'); $b = $fc.Borders.Item(-4107); $b.LineStyle = 1; $b.Color = $Color.Hair; $fc.StopIfTrue = $false   # -4107 = 下罫線
    $An.Range("H$($rd):H$($rl)").NumberFormat = '0.0%;-0.0%;'   # 全 NULL の項目（FILTER で 0 になる）は空欄に
    Add-DataBar $An.Range("D$($rd):D$($rl)") $Color.Gray 1; Add-DataBar $An.Range("H$($rd):H$($rl)") $Color.Burg3 1
    Add-Chips $An.Range("F$($rd):F$($rl)") $Chip.Cls
    $An.Activate(); [void]$An.Range('C5').Select()

    #── ダッシュボード ──────────────────────────────────────────────
    Write-Host '    ダッシュボード'
    $Dash = New-Canvas 'ダッシュボード' $Color.Navy
    $Dash.Columns.Item(1).ColumnWidth = 2
    foreach ($e in @(@(1, 10), @(2, 34), @(3, 18), @(4, 8), @(5, 22), @(6, 8))) { $Dash.Rows.Item($e[0]).RowHeight = $e[1] }
    $Dash.Range('B2').Value2 = 'データプロファイル・レポート'; $f = $Dash.Range('B2').Font; $f.Size = 20; $f.Bold = $true; $f.Color = $Color.Navy
    $elapsed = (Get-Date) - $Started
    $Dash.Range('B3').Value2 = "対象: $CsvDir　／　テーブル $($Tables.Count)　／　総件数 $(Format-N $total) 件　／　作成 $(Get-Date -f 'yyyy/MM/dd HH:mm')"
    $Dash.Range('B3').Font.Size = 9; $Dash.Range('B3').Font.Color = $Color.Ink2
    $X1 = $Dash.Range('B1').Left; $DW = 960; $chartW = ($DW - 16) / 2; $X2 = $X1 + $chartW + 16
    $nav = 'テーブル分析', 'テーブル一覧', 'データ分布', 'インデックス提案', '検索リスク警告', '値の分布', 'キー先頭1桁', '全データ検索', '全データ', '判定基準'
    $bw = ($DW - 9 * 6) / 10; for ($i = 0; $i -lt $nav.Count; $i++) { Add-NavButton $Dash ($X1 + $i * ($bw + 6)) $Dash.Rows.Item(5).Top $bw 21 $nav[$i] $nav[$i] }

    # KPI
    $y = $Dash.Rows.Item(7).Top + 6; $cardW = ($DW - 5 * 12) / 6; $ch = 72
    $empty = @($Tables | Where-Object { $_.Rows -eq 0 }).Count; $big = $byRows[0]
    $sevN = @{}; foreach ($s in '重大', '警告', '注意', '情報') { $sevN[$s] = @($Warns | Where-Object Sev -eq $s).Count }
    $kpis = @(
        @('テーブル数', (Format-N $NT), "0 件のテーブル: $empty", $Color.Navy),
        @('総データ件数', (Format-N $total), "最大: $(Cut $big.Name 18)（$(Format-N $big.Rows)）", $Color.Navy),
        @('総項目数', (Format-N $Profiles.Count), "低カーディナリティ: $(@($Profiles | Where-Object { $_.Class -eq '低' }).Count) 列", $Color.Navy),
        @('インデックス提案', (Format-N $Idx.Count), "優先度 高: $(@($Idx | Where-Object Pri -eq '高').Count) ／ 中: $(@($Idx | Where-Object Pri -eq '中').Count)", $Color.Navy),
        @('重大', (Format-N $sevN['重大']), "対象テーブル: $(@($Warns | Where-Object Sev -eq '重大' | Select-Object -ExpandProperty Table -Unique).Count)", $Color.Burg),
        @('警告', (Format-N $sevN['警告']), "注意: $($sevN['注意']) ／ 情報: $($sevN['情報'])", $Color.Burg2))
    for ($i = 0; $i -lt 6; $i++) { $k = $kpis[$i]; Add-Card $Dash ($X1 + $i * ($cardW + 12)) $y $cardW $ch $k[0] $k[1] $k[2] $k[3] }
    $y += $ch + 16

    # 1 段目: 件数の多いテーブル／少ないテーブル（件数順の横棒）
    $hh = [Math]::Max(240, [Math]::Max($topN, $botN) * 18 + 90)
    foreach ($e in @(@($X1, 1, $topN, '件数の多いテーブル', "上位 $topN ／ 全 $NT テーブル", $Color.Navy), @($X2, 4, $botN, '件数の少ないテーブル', "下位 $botN（少ない順）", $Color.Navy3))) {
        if (!$e[2]) { Add-Empty $Dash $e[0] $y $chartW $hh $e[3] 'テーブルが 1 つのため省略'; continue }
        $c = New-Chart $Dash $e[0] $y $chartW $hh
        $s = Add-Series $c '件数' (GRange 2 ($e[1] + 1) $e[2]) (GRange 2 $e[1] $e[2]); $c.ChartType = 57
        Format-Chart $c $e[3] $e[4]; $c.HasLegend = $false
        Set-Fill $s $e[5]; Set-Gap $c $e[2] ($hh - 70); Set-Labels $s '#,##0'
        $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.ReversePlotOrder = $true
        $mx = $(if ($e[1] -eq 4) { ($byRows | Select-Object -Last $botN | ForEach-Object { $_.Rows } | Measure-Object -Maximum).Maximum } else { $maxRows })
        $vx = $c.Axes(2); Format-Axis $vx '' $true; $vx.MinimumScale = [double]0; $vx.MaximumScale = [double][Math]::Max(1.0, $mx * 1.22)
    }
    $y += $hh + 16

    # 2 段目: 規模別テーブル数／カーディナリティ区分別の項目数（順序のある区分 → ネイビーの濃淡、低はバーガンディ）
    $hh = 250
    $c = New-Chart $Dash $X1 $y $chartW $hh
    $s = Add-Series $c 'テーブル数' (GRange 2 8 7) (GRange 2 7 7); $c.ChartType = 51
    Format-Chart $c '規模別テーブル数' '件数の桁ごとのテーブル数'; $c.HasLegend = $false
    Set-Fill $s $Color.Navy; $s.Points(1).Format.Fill.ForeColor.RGB = $Color.Gray; $c.ChartGroups(1).GapWidth = 60; Set-Labels $s '#,##0;;'
    $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $vx = $c.Axes(2); Format-Axis $vx '' $true; $vx.MinimumScale = [double]0
    $c = New-Chart $Dash $X2 $y $chartW $hh
    $s = Add-Series $c '項目数' (GRange 2 11 6) (GRange 2 10 6); $c.ChartType = 51
    Format-Chart $c 'カーディナリティ区分別の項目数' "全 $(Format-N $Profiles.Count) 項目（ユニーク数で区分）"; $c.HasLegend = $false
    Set-Fill $s $Color.Navy; $pc = @($Color.Navy, $Color.Navy2, $Color.Navy3, $Color.Burg, $Color.Gray, $Color.Rule); for ($i = 1; $i -le 6; $i++) { $s.Points($i).Format.Fill.ForeColor.RGB = $pc[$i - 1] }
    $c.ChartGroups(1).GapWidth = 60; Set-Labels $s '#,##0;;'
    $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $vx = $c.Axes(2); Format-Axis $vx '' $true; $vx.MinimumScale = [double]0
    $y += $hh + 16

    # 3 段目: カーディナリティマップ（散布図・両対数）／検索リスクの内訳（積み上げ横棒）
    $hh = 340
    $c = New-Chart $Dash $X1 $y $chartW $hh
    $sty = [Collections.Generic.List[object]]::new()   # 系列ごとの書式: 点（色・透明度）／線（色）
    if ($ptHi.Count) { [void](Add-Series $c "ユニーク数 $($TH.LowCard) 以上" (GRange 2 14 $ptHi.Count) (GRange 2 13 $ptHi.Count)); $sty.Add(@('点', $Color.Navy, 0.35)) }
    if ($ptLo.Count) { [void](Add-Series $c "低カーディナリティ（$($TH.LowCard) 未満）" (GRange 2 16 $ptLo.Count) (GRange 2 15 $ptLo.Count)); $sty.Add(@('点', $Color.Burg, 0.2)) }
    [void](Add-Series $c '一意（ユニーク数 = 件数）' (GRange 2 18 2) (GRange 2 17 2)); $sty.Add(@('線', $Color.Gray, 0))
    [void](Add-Series $c "低カーディナリティの境界（$($TH.LowCard)）" (GRange 2 20 2) (GRange 2 19 2)); $sty.Add(@('線', $Color.Burg3, 0))
    $c.ChartType = -4169
    for ($i = 0; $i -lt $sty.Count; $i++) {
        $s = $c.SeriesCollection($i + 1); $e = $sty[$i]
        if ($e[0] -eq '点') { $s.MarkerStyle = 8; $s.MarkerSize = 6; $s.MarkerBackgroundColor = $e[1]; $s.MarkerForegroundColor = $Color.White }   # 透明度を設定すると色が自動に戻るため使わない
        else { $s.ChartType = 75; $s = $c.SeriesCollection($i + 1); $s.Format.Line.Visible = -1; $s.Format.Line.ForeColor.RGB = $e[1]; $s.Format.Line.Weight = 1; $s.Format.Line.DashStyle = 4 }
    }
    Format-Chart $c 'カーディナリティマップ' '1 点 = 1 項目。破線の対角線上は一意、横の破線より下は低カーディナリティ'
    foreach ($i in 1, 2) {
        $ax = $c.Axes($i)
        Format-Axis $ax '#,##0' $false $true $true $true
        $ax.MinimumScale = [double]1
        $ax.MaximumScale = [double]$axMax
        $ax.HasTitle = $true; $ax.AxisTitle.Text = $(if ($i -eq 1) { '非 NULL 件数（対数）' } else { 'ユニーク数（対数）' }); $ax.AxisTitle.Font.Size = 8; $ax.AxisTitle.Font.Bold = $false; $ax.AxisTitle.Font.Color = $Color.Muted
    }
    Set-Legend $c
    if ($wCats.Count) {
        $c = New-Chart $Dash $X2 $y $chartW $hh
        $sevC = @{ '重大' = $Color.Burg; '警告' = $Color.Burg2; '注意' = $Color.Burg3; '情報' = $Color.Gray }; $i = 0
        $used = @(foreach ($sv in '重大', '警告', '注意', '情報') { if ($sevN[$sv]) { [void](Add-Series $c $sv (GRange 2 (23 + $i) $wCats.Count) (GRange 2 22 $wCats.Count)); $sv }; $i++ })   # 0 件の重要度は系列にしない
        $c.ChartType = 58
        for ($i = 0; $i -lt $used.Count; $i++) { $sc = $c.SeriesCollection($i + 1); Set-Fill $sc $sevC[$used[$i]] $true }
        Format-Chart $c '検索リスクの内訳' "分類別・重要度別の件数（全 $(Format-N $Warns.Count) 件）"
        Set-Gap $c $wCats.Count ($hh - 100)
        $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.ReversePlotOrder = $true; $ax.Crosses = 2; $ax.TickLabels.Font.Size = 8
        $vx = $c.Axes(2); Format-Axis $vx '#,##0' $false $true; $vx.MinimumScale = [double]0; if ($wCats[0].Count -le 8) { $vx.MajorUnit = [double]1 }
        Set-Legend $c
    } else { Add-Empty $Dash $X2 $y $chartW $hh '検索リスクの内訳' '警告はありません' }
    $y += $hh + 16

    # 4 段目: インデックス提案の内訳／対応を優先すべき警告
    $hh = [Math]::Max(250, $iKinds.Count * 26 + 110)
    if ($iKinds.Count) {
        $c = New-Chart $Dash $X1 $y $chartW $hh
        $priC = @($Color.Navy, $Color.Navy3, $Color.Navy4); $i = 0
        $used = @(foreach ($pr in '高', '中', '低') { if (@($Idx | Where-Object Pri -eq $pr).Count) { [void](Add-Series $c "優先度 $pr" (GRange 2 (29 + $i) $iKinds.Count) (GRange 2 28 $iKinds.Count)); $i }; $i++ })
        $c.ChartType = 58
        for ($i = 0; $i -lt $used.Count; $i++) { $sc = $c.SeriesCollection($i + 1); Set-Fill $sc $priC[$used[$i]] $true }
        Format-Chart $c 'インデックス提案の内訳' "種別・優先度別の件数（全 $(Format-N $Idx.Count) 件）"
        Set-Gap $c $iKinds.Count ($hh - 100)
        $ax = $c.Axes(1); Format-Axis $ax '' $false $false $true; $ax.ReversePlotOrder = $true; $ax.Crosses = 2
        $vx = $c.Axes(2); Format-Axis $vx '#,##0' $false $true; $vx.MinimumScale = [double]0; if ($iKinds[0].Count -le 8) { $vx.MajorUnit = [double]1 }
        Set-Legend $c
    } else { Add-Empty $Dash $X1 $y $chartW $hh 'インデックス提案の内訳' '提案はありません' }
    # 対応を優先すべき警告（重大・警告を大きいテーブル順に）
    $bg = $Dash.Shapes.AddShape(1, $X2, $y, $chartW, $hh); $bg.Fill.ForeColor.RGB = $Color.White; $bg.Line.Visible = 0; $bg.Shadow.Visible = 0; $bg.Placement = 3
    [void](Add-Text $Dash ($X2 + 10) ($y + 7) ($chartW - 20) 16 '対応を優先すべき警告' 11 $Color.Ink $true)
    [void](Add-Text $Dash ($X2 + 10) ($y + 25) ($chartW - 20) 14 '重大・警告のうち、件数の多いテーブルから最大 12 件（全件は「検索リスク警告」シート）' 8.5 $Color.Muted)
    $rowsOf = @{}; foreach ($tb in $Tables) { $rowsOf[$tb.Name] = $tb.Rows }
    $hot = @($Warns | Where-Object { $_.Sev -in '重大', '警告' } | Sort-Object { $SevOrder[$_.Sev] }, @{ Expression = { $rowsOf[$_.Table] }; Descending = $true } | Select-Object -First 12)
    if ($hot.Count) {
        $lines = @($hot | ForEach-Object { Limit-Width "$($_.Sev)　$($_.Table).$($_.Col)　$($_.Cat)：$($_.Detail)" 92 })
        $tb = Add-Text $Dash ($X2 + 10) ($y + 46) ($chartW - 20) ($hh - 54) ($lines -join "`r") 8.5 $Color.Ink
        $tb.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 3
        $pos = 1; foreach ($l in $lines) { $sr = $tb.TextFrame2.TextRange.Characters($pos, 2); $sr.Font.Bold = -1; $sr.Font.Fill.ForeColor.RGB = $(if ($l.StartsWith('重大')) { $Color.Burg } else { $Color.Burg2 }); $pos += $l.Length + 1 }
    } else { [void](Add-Text $Dash ($X2 + 10) ($y + $hh / 2) ($chartW - 20) 16 '重大・警告はありません' 9 $Color.Muted $false '' 2) }
    $y += $hh + 16
    [void](Add-Text $Dash $X1 $y $DW 14 "処理時間 $([int]$elapsed.TotalMinutes) 分 $($elapsed.Seconds) 秒　／　インデックス提案は CSV のデータからの推定です。既存のインデックス・実際の SQL を踏まえ、EXPLAIN で効果を確認してください。" 8 $Color.Muted)

    #── 仕上げ ──────────────────────────────────────────────────────
    $GD.Visible = 0   # グラフデータは非表示
    $An.Move($wb.Worksheets.Item(1)); $Dash.Move($wb.Worksheets.Item(1))
    [void]$init.Delete()   # 初期シートを削除
    $xl.Calculation = -4105; $xl.CalculateFull()
    Invoke-Retry { $Dash.Activate(); [void]$Dash.Range('A1').Select(); $xl.ActiveWindow.ScrollRow = 1 }
    Invoke-Retry { $wb.SaveAs($OutFile, 51) }
    Invoke-Retry { $wb.Close($false) }
} finally { Invoke-Retry { $xl.Quit() }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
Write-Host ''
Write-Host "  出力完了: $OutFile"
Write-Host ("  テーブル {0} ／ 総件数 {1:N0} ／ インデックス提案 {2} ／ 警告 重大 {3}・警告 {4}・注意 {5}" -f $Tables.Count, $total, $Idx.Count, $sevN['重大'], $sevN['警告'], $sevN['注意'])

#endregion
