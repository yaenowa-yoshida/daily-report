# Nippo-Batch.ps1 — 15分ごとに作業内容を記録する常駐アプリの本体
#
# 使い方・運用フローは同じフォルダの ユーザーマニュアル.xlsx、構成と保守メモは README.md を参照。
# このヘッダにはコードを読むために必要な構造だけを書く。
#
# 構成:
#   1. 定数とパス       … 日報フォルダ配下のファイル位置、UI共通の定数
#   2. 共通ユーティリティ … 15分枠の計算、CSV行の組み立て、メッセージ表示
#   3. 設定ファイル      … 案件.txt / 作業名.txt の読み書き
#   4. 日報CSV          … 日報ログ_YYYY-MM-DD.csv の読み書き（1行 = 15分枠）
#   5. 記録位置         … 「次に記録すべき枠」($pendingStart) の管理
#   6. UI部品ファクトリ  … フォント・スタイルを1か所に集約したコントロール生成
#   7. 機能テーブル      … 機能ID → 表記と処理（入力ウィンドウ／日報メニュー／トレイ共通）
#   8. 入力ウィンドウ    … 15分区切りごとに出る主入力画面
#   9. 休憩中ウィンドウ  … 「休憩完了」を押した時点で休憩枠を確定する
#  10. 一覧編集ウィンドウ … 1日分を表で編集。業務終了時の確認画面も兼ねる
#  11. MTG予約ウィンドウ … 先の時間帯を事前に記録する
#  12. 設定メンテナンス  … 案件・作業内容の選択肢（案件.txt / 作業名.txt）を画面から編集する
#  13. 日報メニュー      … 常時表示の小窓（編集／MTG予約／休憩／選択肢の編集／業務終了）
#  14. セルフテスト      … -TestInit。ウィンドウを出さずに主要処理を検証する
#  15. 画面キャプチャ    … -Capture。各ウィンドウのPNGを 画像\ に書き出す（資料用）
#  16. メイン           … 二重起動防止 → 業務開始ゲート → トレイ常駐 → メインループ
#
# 記録の考え方: 「これからやる作業」を先に記録する。$pendingStart（未記録の先頭の枠）から
#   「いまの枠＋選択した時間」までを、同じ内容の15分枠として一気に書く。
#   MTG予約で先に埋まっている枠は二重に書かず、入力ウィンドウも出さずに飛ばす。
#
# CSV: 日報ログ\日報ログ_YYYY-MM-DD.csv（UTF-8 BOM、列: 日付,時刻,案件,作業内容。1行 = 15分枠）
#
# 配置: このスクリプトと 日報入力.bat・日報ログ\・日報設定\ は同じ「日報記録」フォルダに置く。
#   起動は 日報入力.bat 経由（利用者はデスクトップのショートカットから毎朝手動で起動する）。
#   スタートアップ登録とexe化は、どちらも組織のポリシーで使えないため採用しない。
#   代わりに、起動時にデスクトップのショートカットが無ければ作り直す（Initialize-DesktopShortcut）。

param([switch]$TestInit, [switch]$Capture)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# 1. 定数とパス
# ============================================================

# ログ・設定はこのスクリプトと同じフォルダ（日報記録）に置く
$reportDir = $PSScriptRoot
$logDir    = Join-Path $reportDir '日報ログ'
$configDir = Join-Path $reportDir '日報設定'
$projectListPath = Join-Path $configDir '案件.txt'
$taskListPath    = Join-Path $configDir '作業名.txt'

$SLOT_MINUTES = 15                                          # 記録の最小単位（1行 = この分数）
$DURATIONS = @('15', '30', '45', '60', '75', '90', '120')   # 「記録する時間」で選べる分数
$RESERVE_DURATIONS = @('30', '60')                          # 「MTG予約」で選べる分数
# 案件.txt が無い（または空の）ときだけ使う雛形。実際の案件名は 案件.txt 側で管理する
$DEFAULT_PROJECTS = @('案件A', '案件B', '案件C', '社内', 'その他')
$BREAK_PROJECT = 'その他'   # 休憩を記録するときの案件
$BREAK_TASK    = '休憩'     # 休憩を記録するときの作業内容
$DATE_FMT = 'yyyy/MM/dd'
$TIME_FMT = 'H:mm'
$SLOT_FMT = 'yyyy/MM/dd H:mm'   # 書き出し形式。枠を一意に指すキーにも使う
# 読み込み時に許容する形式。ExcelでCSVを開いて保存すると 2026/8/7 や 9:15:00 に変わることがあるため、
# 桁揃えなし・秒つきも受け付ける（書き戻すときは $SLOT_FMT に揃うので形式は自動で正される）
$SLOT_PARSE_FMTS = [string[]]@(
    'yyyy/MM/dd H:mm', 'yyyy/M/d H:mm', 'yyyy/MM/dd H:mm:ss', 'yyyy/M/d H:mm:ss',
    'yyyy-MM-dd H:mm', 'yyyy-M-d H:mm', 'yyyy-MM-dd H:mm:ss', 'yyyy-M-d H:mm:ss'
)

$Utf8Bom   = New-Object Text.UTF8Encoding $true    # 新規書き出し用（Excelで開けるようBOM付き）
$Utf8NoBom = New-Object Text.UTF8Encoding $false   # 既存ファイルへの追記用（BOMを重ねない）
$UiFont    = New-Object Drawing.Font('Meiryo UI', 10)

# ============================================================
# 2. 共通ユーティリティ
# ============================================================

function Get-LogPath([datetime]$d) { return (Join-Path $logDir ('日報ログ_{0:yyyy-MM-dd}.csv' -f $d)) }

# その時刻が属する15分枠の先頭（9:07 → 9:00）
function Get-FloorSlot([datetime]$t) {
    return $t.Date.AddMinutes([math]::Floor(($t.Hour * 60 + $t.Minute) / $SLOT_MINUTES) * $SLOT_MINUTES)
}

# 枠を一意に指すキー。記録済み判定のハッシュテーブルで使う
function Get-SlotKey([datetime]$t) { return $t.ToString($SLOT_FMT) }

# CSVの「日付」「時刻」列（例: 2026/08/06 と 9:15）を [datetime] に戻す。読めない値は例外
function ConvertTo-SlotTime([string]$dateStr, [string]$timeStr) {
    return [datetime]::ParseExact("$dateStr $timeStr", $SLOT_PARSE_FMTS,
        [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None)
}

# 同上。手編集で壊れた行を読み飛ばしたい場所で使う（読めなければ $null）
function ConvertTo-SlotTimeOrNull([string]$dateStr, [string]$timeStr) {
    try { return (ConvertTo-SlotTime $dateStr $timeStr) } catch { return $null }
}

# 各値をダブルクォートで囲んでCSVの1行にする
function ConvertTo-CsvLine([string[]]$fields) {
    $quoted = foreach ($f in $fields) { '"' + $f.Replace('"', '""') + '"' }
    return ($quoted -join ',')
}

function Read-LogRows([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    return @(Get-Content -LiteralPath $path -Encoding UTF8 | ConvertFrom-Csv)
}

# OKボタンだけの案内メッセージ
function Show-Info([string]$text, [string]$title = '日報入力') {
    [void][Windows.Forms.MessageBox]::Show($text, $title)
}

# CSV書き込み失敗時の案内（原因はExcelでの排他がほとんど）。$hint に画面ごとの次の一手を書く
function Show-CsvError($errorRecord, [string]$hint = '', [string]$title = '日報入力') {
    $msg = 'CSVに書き込めませんでした。ExcelなどでCSVを開いていたら閉じてください。'
    if ($hint) { $msg += $hint }
    Show-Info ($msg + "`n`n" + $errorRecord.Exception.Message) $title
}

# ドロップダウンで選ばれている分数（数字以外なら既定値）
function Get-ComboMinutes($combo, [int]$default) {
    if ($combo.Text -match '^\d+$') { return [int]$combo.Text }
    return $default
}

# ============================================================
# 3. 設定ファイル（案件.txt / 作業名.txt）
# ============================================================

# 設定ファイル（1行1項目、#で始まる行はコメント）を読む。重複は除去
function Read-ListFile([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $seen  = New-Object Collections.Generic.HashSet[string]
    $items = New-Object Collections.Generic.List[string]
    foreach ($line in @(Get-Content -LiteralPath $path -Encoding UTF8)) {
        $s = ([string]$line).Trim()
        if ($s -and -not $s.StartsWith('#') -and $seen.Add($s)) { $items.Add($s) }
    }
    return @($items.ToArray())
}

function Get-ProjectList {
    $items = Read-ListFile $projectListPath
    if ($items.Count -eq 0) { return $DEFAULT_PROJECTS }
    return $items
}

# 作業名.txt を解析する。[案件名] の見出しでセクション分けされ、
# 見出しより前の行はどの案件でも出す「共通」扱い（キー ''）
function Read-TaskMap {
    $map = [ordered]@{ '' = (New-Object Collections.Generic.List[string]) }
    $current = ''
    if (Test-Path -LiteralPath $taskListPath) {
        foreach ($line in @(Get-Content -LiteralPath $taskListPath -Encoding UTF8)) {
            $s = ([string]$line).Trim()
            if (-not $s -or $s.StartsWith('#')) { continue }
            if ($s -match '^\[(.+)\]$') {
                $current = $Matches[1].Trim()
                if (-not $map.Contains($current)) { $map[$current] = New-Object Collections.Generic.List[string] }
                continue
            }
            if (-not $map[$current].Contains($s)) { $map[$current].Add($s) }
        }
    }
    return $map
}

# 指定案件の選択肢（その案件のセクション＋共通分）
function Get-TasksForProject([string]$project) {
    $map = Read-TaskMap
    $seen  = New-Object Collections.Generic.HashSet[string]
    $items = New-Object Collections.Generic.List[string]
    if ($project -and $map.Contains($project)) {
        foreach ($t in $map[$project]) { if ($seen.Add($t)) { $items.Add($t) } }
    }
    foreach ($t in $map['']) { if ($seen.Add($t)) { $items.Add($t) } }
    return @($items.ToArray())
}

# 全案件の作業名（重複除去）
function Get-TaskList {
    $map = Read-TaskMap
    $seen  = New-Object Collections.Generic.HashSet[string]
    $items = New-Object Collections.Generic.List[string]
    foreach ($key in @($map.Keys)) {
        foreach ($t in $map[$key]) { if ($seen.Add($t)) { $items.Add($t) } }
    }
    return @($items.ToArray())
}

# 作業名.txt の [案件] セクション末尾に作業名を追記（セクションが無ければ末尾に新設）
function Add-TaskToFile([string]$project, [string]$task) {
    $lines = New-Object Collections.Generic.List[string]
    if (Test-Path -LiteralPath $taskListPath) {
        foreach ($l in @(Get-Content -LiteralPath $taskListPath -Encoding UTF8)) { $lines.Add([string]$l) }
    }
    $headerIdx = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Trim() -match '^\[(.+)\]$') {
            if ($Matches[1].Trim() -eq $project) { $headerIdx = $i; break }
        }
    }
    if ($headerIdx -lt 0) {
        if ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim()) { $lines.Add('') }
        $lines.Add("[$project]")
        $lines.Add($task)
    } else {
        $insertAt = $lines.Count
        for ($i = $headerIdx + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i].Trim() -match '^\[.+\]$') { $insertAt = $i; break }
        }
        # 次の見出し直前の空行は飛ばしてセクション本体の末尾に入れる
        while ($insertAt -gt $headerIdx + 1 -and -not $lines[$insertAt - 1].Trim()) { $insertAt-- }
        $lines.Insert($insertAt, $task)
    }
    [IO.File]::WriteAllLines($taskListPath, $lines.ToArray(), $Utf8Bom)
}

# 記録に使った作業名が選択肢に無ければ登録する（次回から選べるように）。
# 失敗しても記録自体は成立しているので、握りつぶして続行する
function Register-TaskIfNew([string]$project, [string]$task) {
    if (-not $project -or -not $task) { return }
    if ((Get-TasksForProject $project) -contains $task) { return }
    try { Add-TaskToFile $project $task } catch {}
}

# --- 設定ファイルの書き戻し（設定メンテナンスウィンドウ用） ---
# 読み取り側（Read-ListFile / Read-TaskMap）は選択肢を作るだけなのでコメントを捨ててよいが、
# 書き戻し側はファイルを丸ごと作り直すため、コメント行も保持して書き戻す。
#   Header   … ファイル冒頭の説明コメント（そのまま先頭に戻す）
#   Comments … 本文中のコメント行（無効化した項目のメモなど）。該当セクションの末尾に寄せて残す

# 案件.txt を「冒頭コメント／項目／本文中のコメント」に分けて読む
function Read-ProjectFileModel {
    $model = @{
        Header   = New-Object Collections.Generic.List[string]
        Items    = New-Object Collections.Generic.List[string]
        Comments = New-Object Collections.Generic.List[string]
    }
    if (-not (Test-Path -LiteralPath $projectListPath)) { return $model }
    $inHeader = $true
    foreach ($line in @(Get-Content -LiteralPath $projectListPath -Encoding UTF8)) {
        $s = ([string]$line).Trim()
        if ($inHeader) {
            if (-not $s) { continue }
            if ($s.StartsWith('#')) { $model.Header.Add($s); continue }
            $inHeader = $false
        }
        if (-not $s) { continue }
        if ($s.StartsWith('#')) { $model.Comments.Add($s); continue }
        if (-not $model.Items.Contains($s)) { $model.Items.Add($s) }
    }
    return $model
}

function Write-ProjectFileModel($model) {
    $out = New-Object Collections.Generic.List[string]
    foreach ($h in $model.Header) { $out.Add($h) }
    foreach ($p in $model.Items) { $out.Add($p) }
    foreach ($c in $model.Comments) { $out.Add($c) }
    [IO.File]::WriteAllLines($projectListPath, $out.ToArray(), $Utf8Bom)
}

# 作業名.txt を「冒頭コメント／セクション（案件）ごとの項目とコメント」に分けて読む。
# キー '' は見出しより前の共通セクション（常に先頭に存在する）
function Read-TaskFileModel {
    $model = @{
        Header   = New-Object Collections.Generic.List[string]
        Keys     = New-Object Collections.Generic.List[string]
        Items    = @{}
        Comments = @{}
    }
    $addKey = {
        param($k)
        if (-not $model.Keys.Contains($k)) {
            $model.Keys.Add($k)
            $model.Items[$k]    = New-Object Collections.Generic.List[string]
            $model.Comments[$k] = New-Object Collections.Generic.List[string]
        }
    }
    & $addKey ''
    if (-not (Test-Path -LiteralPath $taskListPath)) { return $model }
    $cur = ''
    $inHeader = $true
    foreach ($line in @(Get-Content -LiteralPath $taskListPath -Encoding UTF8)) {
        $s = ([string]$line).Trim()
        if ($inHeader) {
            if (-not $s) { continue }
            if ($s.StartsWith('#')) { $model.Header.Add($s); continue }
            $inHeader = $false
        }
        if (-not $s) { continue }
        if ($s -match '^\[(.+)\]$') { $cur = $Matches[1].Trim(); & $addKey $cur; continue }
        if ($s.StartsWith('#')) { $model.Comments[$cur].Add($s); continue }
        if (-not $model.Items[$cur].Contains($s)) { $model.Items[$cur].Add($s) }
    }
    return $model
}

function Write-TaskFileModel($model) {
    $out = New-Object Collections.Generic.List[string]
    foreach ($h in $model.Header) { $out.Add($h) }
    foreach ($key in $model.Keys) {
        $items    = $model.Items[$key]
        $comments = $model.Comments[$key]
        if ($key -eq '') {
            if ($items.Count -eq 0 -and $comments.Count -eq 0) { continue }
            if ($out.Count -gt 0) { $out.Add('') }
        } else {
            if ($out.Count -gt 0) { $out.Add('') }
            $out.Add("[$key]")
        }
        foreach ($t in $items) { $out.Add($t) }
        foreach ($c in $comments) { $out.Add($c) }
    }
    [IO.File]::WriteAllLines($taskListPath, $out.ToArray(), $Utf8Bom)
}

# 設定ファイルが無ければ雛形を作る
function Initialize-ConfigFiles {
    if (-not (Test-Path -LiteralPath $configDir)) { New-Item -ItemType Directory -Path $configDir | Out-Null }
    if (-not (Test-Path -LiteralPath $projectListPath)) {
        [IO.File]::WriteAllText($projectListPath,
            ("# 案件（ブランド）一覧。1行1項目。#で始まる行は無視されます`r`n" + ($DEFAULT_PROJECTS -join "`r`n") + "`r`n"), $Utf8Bom)
    }
    if (-not (Test-Path -LiteralPath $taskListPath)) {
        [IO.File]::WriteAllText($taskListPath,
            ("# 作業名一覧。[案件名] の見出しの下に、その案件の作業名を1行1項目で書きます`r`n" +
             "# 見出しより前に書いた行はどの案件でも選択肢に出ます（共通）。#で始まる行は無視されます`r`n" +
             "# 新しい作業名をウィンドウで直接入力して記録すると、選択中の案件の見出しの下に自動追加されます`r`n"), $Utf8Bom)
    }
}
Initialize-ConfigFiles

# デスクトップに起動用ショートカットが無ければ作る。
# スタートアップ登録が組織のポリシーで使えないため、デスクトップのアイコンが唯一の起動口になる。
# 消えると起動手段を見失うので、毎回の起動で作り直す
function Initialize-DesktopShortcut {
    $batPath = Join-Path $reportDir '日報入力.bat'
    if (-not (Test-Path -LiteralPath $batPath)) { return }
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) '日報入力.lnk'
    if (Test-Path -LiteralPath $lnkPath) { return }
    $ws = New-Object -ComObject WScript.Shell
    $lnk = $ws.CreateShortcut($lnkPath)
    $lnk.TargetPath       = $batPath
    $lnk.WorkingDirectory = $reportDir
    $lnk.WindowStyle      = 7   # 最小化（batの黒い窓が前面に出ないように）
    $lnk.Description      = '日報入力（15分ごとの作業記録）'
    $lnk.Save()
    # 勝手にファイルが増えたように見えないよう、作ったことを必ず伝える（OKを押すまで閉じない）
    Show-Info ("デスクトップに「日報入力」のショートカットを作成しました。`n" +
               "次回からは、このアイコンをダブルクリックして起動してください。`n`n" +
               "作成先: $lnkPath")
}

# ============================================================
# 4. 日報CSV（日報ログ_YYYY-MM-DD.csv）
# ============================================================

# その日のCSVで既に埋まっている枠のキー集合。表記ゆれを吸収するため一度[datetime]に直してから鍵にする
function Get-RecordedSlotKeys([string]$path) {
    $recorded = @{}
    foreach ($r in @(Read-LogRows $path)) {
        $rt = ConvertTo-SlotTimeOrNull ([string]$r.'日付') ([string]$r.'時刻')
        if ($rt) { $recorded[(Get-SlotKey $rt)] = $true }
    }
    return $recorded
}

# [$from, $toExclusive) の15分枠を1行ずつCSVに追記し、書いた枠数を返す。
# MTG予約などで先に記録済みの枠は二重に書かずスキップする
function Write-Slots([datetime]$from, [datetime]$toExclusive, [string]$project, [string]$task) {
    $path = Get-LogPath $from
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    if (-not (Test-Path -LiteralPath $path)) {
        [IO.File]::WriteAllText($path, "日付,時刻,案件,作業内容`r`n", $Utf8Bom)
    }
    $existing = Get-RecordedSlotKeys $path
    $sb = New-Object Text.StringBuilder
    $slot = $from
    $n = 0
    while ($slot -lt $toExclusive) {
        if (-not $existing[(Get-SlotKey $slot)]) {
            [void]$sb.AppendLine((ConvertTo-CsvLine @(
                $slot.ToString($DATE_FMT), $slot.ToString($TIME_FMT), $project, $task)))
            $n++
        }
        $slot = $slot.AddMinutes($SLOT_MINUTES)
    }
    if ($n -gt 0) {
        [IO.File]::AppendAllText($path, $sb.ToString(), $Utf8NoBom)
    }
    return $n
}

# 1日分のCSVを丸ごと書き直す（時刻順に整列、UTF-8 BOM）
function Write-DayFile([string]$path, [object[]]$rows) {
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    $sorted = @($rows | Sort-Object {
        $t = ConvertTo-SlotTimeOrNull ([string]$_.'日付') ([string]$_.'時刻')
        if ($t) { $t } else { [datetime]::MaxValue }   # 読めない行は捨てずに末尾へ回す
    })
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine('日付,時刻,案件,作業内容')
    foreach ($r in $sorted) {
        # 読めた行はここで書式を正規化する（Excelで開いて保存された表記ゆれが元に戻る）
        $t = ConvertTo-SlotTimeOrNull ([string]$r.'日付') ([string]$r.'時刻')
        $dateStr = if ($t) { $t.ToString($DATE_FMT) } else { [string]$r.'日付' }
        $timeStr = if ($t) { $t.ToString($TIME_FMT) } else { [string]$r.'時刻' }
        [void]$sb.AppendLine((ConvertTo-CsvLine @(
            $dateStr, $timeStr, [string]$r.'案件', [string]$r.'作業内容')))
    }
    [IO.File]::WriteAllText($path, $sb.ToString(), $Utf8Bom)
}

# 日報ログが存在する日付（yyyy-MM-dd）を新しい順に返す
function Get-LogDates {
    $dates = New-Object Collections.Generic.List[string]
    if (Test-Path -LiteralPath $logDir) {
        foreach ($f in @(Get-ChildItem -LiteralPath $logDir -Filter '日報ログ_*.csv' | Sort-Object Name -Descending)) {
            if ($f.BaseName -match '_(\d{4}-\d{2}-\d{2})$') { $dates.Add($Matches[1]) }
        }
    }
    return @($dates.ToArray())
}

# ============================================================
# 5. 記録位置（$pendingStart = まだ記録していない先頭の枠）
# ============================================================

# 直近の日次ファイルの最終行を「前回の入力」として読む（ウィンドウの初期値に使う）
function Initialize-LastEntry {
    $script:lastProject = ''
    $script:lastTask    = ''
    if (-not (Test-Path -LiteralPath $logDir)) { return }
    $recentFile = Get-ChildItem -LiteralPath $logDir -Filter '日報ログ_*.csv' |
        Sort-Object Name | Select-Object -Last 1
    if (-not $recentFile) { return }
    $rows = Read-LogRows $recentFile.FullName
    if ($rows.Count -gt 0) {
        $script:lastProject = [string]$rows[-1].'案件'
        $script:lastTask    = [string]$rows[-1].'作業内容'
    }
}
Initialize-LastEntry

# 記録開始位置: 当日すでに記録があればその続き、なければ今いる15分枠の頭から。
# 「業務開始」押下のタイミングでCSVを読み直して確定する。
# MTG予約で未来の枠が先に書かれている場合があるので「現在以前で最も遅い記録」を基準にする
# （行の並び順にも依存しない）。未来の予約分はメインループのスキップ処理が飛ばす
function Get-InitialPendingStart {
    $n = Get-Date
    $max = [datetime]::MinValue
    foreach ($r in @(Read-LogRows (Get-LogPath $n))) {
        $rt = ConvertTo-SlotTimeOrNull ([string]$r.'日付') ([string]$r.'時刻')
        if (-not $rt -or $rt.Date -ne $n.Date) { continue }
        if ($rt -le $n -and $rt -gt $max) { $max = $rt }
    }
    if ($max -gt [datetime]::MinValue) { return $max.AddMinutes($SLOT_MINUTES) }
    return (Get-FloorSlot $n)
}
$pendingStart = Get-InitialPendingStart

# MTG予約などで先に記録済みの枠を飛ばして記録開始位置を進める（その時間帯はウィンドウを出さない）
function Update-PendingStart {
    $recorded = Get-RecordedSlotKeys (Get-LogPath (Get-Date))
    while ($recorded[(Get-SlotKey $script:pendingStart)]) {
        $script:pendingStart = $script:pendingStart.AddMinutes($SLOT_MINUTES)
    }
}

# ============================================================
# 6. UI部品ファクトリ（フォント・スタイル設定を1か所に集約）
# ============================================================

# 既定は最前面の固定ダイアログ。-Resident は日報メニュー用（最小化でき、最前面固定にしない）
function New-UiForm([string]$title, [int]$w, [int]$h, [switch]$Resident) {
    $f = New-Object Windows.Forms.Form
    $f.Text        = $title
    $f.ClientSize  = New-Object Drawing.Size($w, $h)
    $f.MaximizeBox = $false
    $f.Font        = $UiFont
    if ($Resident) {
        $f.FormBorderStyle = 'FixedSingle'
        $f.StartPosition   = 'WindowsDefaultLocation'
    } else {
        $f.FormBorderStyle = 'FixedDialog'
        $f.MinimizeBox     = $false
        $f.StartPosition   = 'CenterScreen'
        $f.TopMost         = $true
    }
    return $f
}

# 幅・高さを省略するとAutoSize
function New-UiLabel([int]$x, [int]$y, [string]$text, [int]$w = 0, [int]$h = 0) {
    $l = New-Object Windows.Forms.Label
    $l.Location = New-Object Drawing.Point($x, $y)
    if ($w -gt 0) { $l.Size = New-Object Drawing.Size($w, $h) } else { $l.AutoSize = $true }
    $l.Text = $text
    return $l
}

function New-UiButton([int]$x, [int]$y, [int]$w, [int]$h, [string]$text) {
    $b = New-Object Windows.Forms.Button
    $b.Location = New-Object Drawing.Point($x, $y)
    $b.Size     = New-Object Drawing.Size($w, $h)
    $b.Text     = $text
    return $b
}

# 既定は選択式（DropDownList）。-Editable は直接入力もできる作業名用（入力補完付き）
function New-UiCombo([int]$x, [int]$y, [int]$w, [switch]$Editable) {
    $c = New-Object Windows.Forms.ComboBox
    $c.Location = New-Object Drawing.Point($x, $y)
    $c.Size     = New-Object Drawing.Size($w, 28)
    if ($Editable) {
        $c.AutoCompleteMode   = 'SuggestAppend'
        $c.AutoCompleteSource = 'ListItems'
    } else {
        $c.DropDownStyle = 'DropDownList'
    }
    return $c
}

# 分数のドロップダウン（記録する時間・MTG予約の時間。選択肢はそれぞれ別）
function New-UiDurationCombo([int]$x, [int]$y, [string]$default, [string[]]$items = $DURATIONS) {
    $c = New-UiCombo $x $y 75
    $c.Items.AddRange($items)
    $c.SelectedItem = $default
    return $c
}

# 案件のドロップダウンに選択肢を入れ直す。$select が選択肢にあればそれを選ぶ
function Update-ProjectComboItems($combo, [string[]]$projects, [string]$select = '') {
    $combo.Items.Clear()
    $combo.Items.AddRange($projects)
    if ($select -and ($projects -contains $select)) { $combo.SelectedItem = $select } else { $combo.SelectedIndex = -1 }
}

# 案件に応じて作業名コンボの選択肢を入れ替える（既定では入力中のテキストを維持する）
function Update-TaskComboItems($combo, [string]$project, [switch]$DiscardText) {
    $tasks = @(Get-TasksForProject $project)
    $cur = [string]$combo.Text
    $combo.Items.Clear()
    if ($tasks.Count -gt 0) { $combo.Items.AddRange($tasks) }
    if (-not $DiscardText) { $combo.Text = $cur }
}

# 1行だけ入力させる小さなダイアログ（設定メンテナンスの「追加」「名前変更」用）。
# キャンセル・空入力なら $null を返す。使い捨てなので毎回作って破棄する
function Show-TextPrompt([string]$title, [string]$message, [string]$default = '') {
    $f  = New-UiForm $title 430 175
    $l  = New-UiLabel 15 12 $message 400 40
    $tb = New-Object Windows.Forms.TextBox
    $tb.Location = New-Object Drawing.Point(15, 60)
    $tb.Size     = New-Object Drawing.Size(400, 28)
    $tb.Text     = $default
    $ok = New-UiButton 15 105 130 34 'OK'
    $ok.DialogResult = [Windows.Forms.DialogResult]::OK
    $ng = New-UiButton 285 105 130 34 'キャンセル'
    $ng.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $f.Controls.AddRange(@($l, $tb, $ok, $ng))
    $f.AcceptButton = $ok
    $f.CancelButton = $ng
    $tb.Select()
    $r = $f.ShowDialog()
    $v = ([string]$tb.Text).Trim()
    $f.Dispose()
    if ($r -ne [Windows.Forms.DialogResult]::OK -or -not $v) { return $null }
    return $v
}

# 選択肢から1つ選ばせる小さなダイアログ（作業名の「別の案件へ移動」用）。キャンセルなら $null
function Show-ChoicePrompt([string]$title, [string]$message, [string[]]$items) {
    if ($items.Count -eq 0) { return $null }
    $f = New-UiForm $title 430 175
    $l = New-UiLabel 15 12 $message 400 40
    $c = New-UiCombo 15 60 400
    $c.Items.AddRange($items)
    $c.SelectedIndex = 0
    $ok = New-UiButton 15 105 130 34 'OK'
    $ok.DialogResult = [Windows.Forms.DialogResult]::OK
    $ng = New-UiButton 285 105 130 34 'キャンセル'
    $ng.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $f.Controls.AddRange(@($l, $c, $ok, $ng))
    $f.AcceptButton = $ok
    $f.CancelButton = $ng
    $r = $f.ShowDialog()
    $v = [string]$c.Text
    $f.Dispose()
    if ($r -ne [Windows.Forms.DialogResult]::OK) { return $null }
    return $v
}

# はい／いいえの確認。破棄をともなう操作の前に使う
function Test-Confirm([string]$text, [string]$title) {
    return ([Windows.Forms.MessageBox]::Show($text, $title,
        [Windows.Forms.MessageBoxButtons]::YesNo) -eq [Windows.Forms.DialogResult]::Yes)
}

# ============================================================
# 7. 機能テーブル（機能IDごとに表記と処理を1か所で定義）
# ============================================================
# 入力ウィンドウ・日報メニュー・タスクトレイのどの入口からも Invoke-Feature 経由で同じ処理を呼び、
# ボタンやメニューの表記も Get-FeatureLabel でここから取る（入口ごとの文言・挙動のズレを防ぐ）。
# Action は実行時に解決されるので、ここより後ろで定義する関数を参照してよい
$FEATURES = [ordered]@{
    Edit    = @{ Label = '一覧編集（今日の記録）';             Action = { param($date) if ($date) { [void](Show-GridDialog -Date $date) } else { [void](Show-GridDialog) } } }
    Reserve = @{ Label = 'MTG予約（当日の事前入力）';          Action = { Show-ReserveDialog } }
    Break   = @{ Label = '休憩（戻ったら「休憩完了」で記録）'; Action = { Invoke-BreakFlow } }
    Maint   = @{ Label = '選択肢の編集（案件・作業内容）';     Action = { Show-MaintDialog } }
    Close   = @{ Label = '業務終了（今日の記録を締める）';     Action = { Invoke-ExitFlow } }
}
function Invoke-Feature([string]$id, $arg = $null) { & $script:FEATURES[$id].Action $arg }
function Get-FeatureLabel([string]$id) { return [string]$script:FEATURES[$id].Label }

# ============================================================
# 8. 入力ウィンドウ（15分区切りごとの主入力画面）
# ============================================================
$form = New-UiForm '日報入力（15分ごと）' 500 290

$lblRange = New-UiLabel 15 12 '' 470 26
$lblRange.Font = New-Object Drawing.Font('Meiryo UI', 11, [Drawing.FontStyle]::Bold)

$lblP = New-UiLabel 15 48 '案件（ブランド）:'

# 案件は設定ファイル（案件.txt）からの選択式。選択が変わると作業名の選択肢も切り替わる
$cmbProject = New-UiCombo 15 70 220
$cmbProject.Add_SelectedIndexChanged({ Update-TaskCombo })

$lblT = New-UiLabel 15 105 '作業内容:'

# 作業名は設定ファイル（作業名.txt）から選択（新規は直接入力もOK→ファイルに自動追加）
$cmbTask = New-UiCombo 15 127 470 -Editable

$lblDur    = New-UiLabel 15 168 '記録する時間:'
$cmbDur    = New-UiDurationCombo 125 164 '15'
$lblDurMin = New-UiLabel 207 168 '分（60分MTGなどは長めに選択）'

$keepChk = New-Object Windows.Forms.CheckBox
$keepChk.Location = New-Object Drawing.Point(15, 200)
$keepChk.Size     = New-Object Drawing.Size(300, 24)
$keepChk.Text     = '前回の入力を保持する'
$keepChk.Checked  = $true

$btnOK    = New-UiButton 15 240 120 34 '記録 (OK)'
$btnSkip  = New-UiButton 143 240 110 34 'スキップ'
$btnBreak = New-UiButton 261 240 100 34 '休憩'
$btnExit  = New-UiButton 370 240 115 34 '業務終了'

$form.Controls.AddRange(@($lblRange, $lblP, $cmbProject, $lblT, $cmbTask, $lblDur, $cmbDur, $lblDurMin, $keepChk, $btnOK, $btnSkip, $btnBreak, $btnExit))
$form.AcceptButton = $btnOK

# 選択中の案件に紐づく作業名だけをドロップダウンに出す
function Update-TaskCombo { Update-TaskComboItems $script:cmbTask ([string]$script:cmbProject.Text) }

# 「記録する時間」で選ばれている分数
function Get-SelectedDuration { return (Get-ComboMinutes $script:cmbDur 15) }

# 今回書き込む範囲の終わり（いまの枠＋選択した時間）
function Get-RecordEnd { return (Get-FloorSlot (Get-Date)).AddMinutes((Get-SelectedDuration)) }

# 対象範囲は「保留中の枠 ～ いまの枠＋選択した時間」まで（これからやる作業を先に記録する方式）
function Get-RangeText {
    $to = Get-RecordEnd
    $n = [int](($to - $script:pendingStart).TotalMinutes / $SLOT_MINUTES)
    if ($n -le 0) { return '記録対象の枠はまだありません' }
    return ('対象: {0:H:mm} ～ {1:H:mm}（{2}枠 × {3}分）' -f $script:pendingStart, $to, $n, $SLOT_MINUTES)
}

# 表示中は対象範囲の表示を毎秒更新する（区切りをまたいでも案内がずれない）
$uiTimer = New-Object Windows.Forms.Timer
$uiTimer.Interval = 1000
$uiTimer.Add_Tick({ $script:lblRange.Text = Get-RangeText })

$btnOK.Add_Click({
    $p = $script:cmbProject.Text.Trim()
    $t = $script:cmbTask.Text.Trim()
    if (-not $p -or -not $t) {
        Show-Info '案件と作業内容を入力してください。'
        return
    }
    $to = Get-RecordEnd
    try {
        [void](Write-Slots $script:pendingStart $to $p $t)
    } catch {
        Show-CsvError $_ 'もう一度「記録」を押してください。'
        return
    }
    $script:pendingStart = $to
    $script:lastProject  = $p
    $script:lastTask     = $t
    Register-TaskIfNew $p $t
    $script:form.DialogResult = [Windows.Forms.DialogResult]::OK
})

$btnSkip.Add_Click({
    # 選択した時間ぶん記録せず飛ばす（表には未入力の穴が残る）
    $script:pendingStart = Get-RecordEnd
    $script:form.DialogResult = [Windows.Forms.DialogResult]::Ignore
})

$btnBreak.Add_Click({
    # 終了時間は決めずに休憩に入る（メインループ側で休憩フローに移る）
    $script:form.DialogResult = [Windows.Forms.DialogResult]::Retry
})

$btnExit.Add_Click({ Invoke-Feature 'Close' })

# 入力ウィンドウを出して結果を返す。表示のたびに設定ファイルを読み直す（日中の編集が反映される）
function Show-EntryDialog {
    $keep = $script:keepChk.Checked
    Update-ProjectComboItems $script:cmbProject @(Get-ProjectList) $(if ($keep) { $script:lastProject } else { '' })
    Update-TaskCombo
    $script:cmbTask.Text = $(if ($keep) { $script:lastTask } else { '' })
    $script:cmbDur.SelectedItem = '15'   # 長時間指定の引きずり防止のため毎回15分に戻す
    $script:lblRange.Text = Get-RangeText
    [System.Media.SystemSounds]::Asterisk.Play()
    $script:uiTimer.Start()
    $r = $script:form.ShowDialog()
    $script:uiTimer.Stop()
    return $r
}

# ============================================================
# 9. 休憩中ウィンドウ（「休憩完了」を押した時点で休憩枠を確定する）
# ============================================================
$breakForm = New-UiForm '休憩中' 360 140

$breakLbl = New-UiLabel 15 15 '' 330 40

$btnBreakDone = New-UiButton 15 70 330 48 '休憩完了（業務に戻る）'
$btnBreakDone.Font         = New-Object Drawing.Font('Meiryo UI', 12, [Drawing.FontStyle]::Bold)
$btnBreakDone.DialogResult = [Windows.Forms.DialogResult]::OK

$breakForm.Controls.AddRange(@($breakLbl, $btnBreakDone))
$breakForm.AcceptButton = $btnBreakDone

$breakTimer = New-Object Windows.Forms.Timer
$breakTimer.Interval = 1000
$breakTimer.Add_Tick({
    $min = [int][math]::Floor(((Get-Date) - $script:breakClockStart).TotalMinutes)
    $script:breakLbl.Text = ('休憩中（{0:H:mm} 開始・約{1}分経過）' -f $script:breakClockStart, $min)
})

# 休憩フロー（機能ID 'Break'）: 「休憩完了」まで待ち、そこまでの15分枠を休憩として記録する。
# ×で閉じた場合も完了扱い。業務終了フローから強制クローズ（Abort）された場合は記録しない
function Invoke-BreakFlow {
    $from = $script:pendingStart
    $script:breakClockStart = Get-Date
    $script:breakLbl.Text = ('休憩中（{0:H:mm} 開始）' -f $script:breakClockStart)
    $script:breakTimer.Start()
    $r = $script:breakForm.ShowDialog()
    $script:breakTimer.Stop()
    if ($r -eq [Windows.Forms.DialogResult]::Abort) { return }
    $to = Get-FloorSlot (Get-Date)
    if ($to -le $from) { return }   # 15分未満の休憩は枠が無いので記録なし（すぐ入力ウィンドウに戻る）
    try {
        [void](Write-Slots $from $to $BREAK_PROJECT $BREAK_TASK)
        $script:pendingStart = $to
    } catch {
        Show-CsvError $_ '休憩の時間帯は「一覧編集」からも入力できます。'
    }
}

# ============================================================
# 10. 一覧編集ウィンドウ（1日分を表で編集／業務終了時の確認画面）
# ============================================================
$gridForm = New-UiForm '一覧編集（今日の記録）' 640 535

$grid = New-Object Windows.Forms.DataGridView
$grid.Location = New-Object Drawing.Point(12, 12)
$grid.Size     = New-Object Drawing.Size(616, 380)
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible     = $false
$grid.SelectionMode         = 'CellSelect'
$grid.MultiSelect           = $true
$grid.Add_DataError({ param($s, $e) $e.ThrowException = $false })   # 選択肢に無い案件でも落とさない

$COL_TIME = 0   # 時刻（読み取り専用）
$COL_PROJ = 1   # 案件（ドロップダウン）
$COL_TASK = 2   # 作業内容

$colTime = New-Object Windows.Forms.DataGridViewTextBoxColumn
$colTime.HeaderText = '時刻'
$colTime.Width      = 70
$colTime.ReadOnly   = $true

$colProj = New-Object Windows.Forms.DataGridViewComboBoxColumn
$colProj.HeaderText = '案件'
$colProj.Width      = 150
$colProj.FlatStyle  = 'Flat'

$colTask = New-Object Windows.Forms.DataGridViewTextBoxColumn
$colTask.HeaderText = '作業内容'
$colTask.Width      = 370

[void]$grid.Columns.Add($colTime)
[void]$grid.Columns.Add($colProj)
[void]$grid.Columns.Add($colTask)

$gLblFill = New-UiLabel 12 406 '一括入力:'

$gCmbProject = New-UiCombo 95 402 150
$gCmbProject.Add_SelectedIndexChanged({ Update-GridTaskCombo })

$gCmbTask = New-UiCombo 255 402 373 -Editable

$gBtnApply  = New-UiButton 95 438 170 32 '選択した行に入力'
$gBtnClear  = New-UiButton 275 438 170 32 '選択した行をクリア'
$gBtnCopy   = New-UiButton 455 438 173 32 'この行を入力欄へコピー'
$gBtnSave   = New-UiButton 12 486 210 36 '保存して閉じる'
$gBtnCancel = New-UiButton 488 486 140 36 'キャンセル'
$gBtnCancel.DialogResult = [Windows.Forms.DialogResult]::Cancel

$gridForm.Controls.AddRange(@($grid, $gLblFill, $gCmbProject, $gCmbTask, $gBtnApply, $gBtnClear, $gBtnCopy, $gBtnSave, $gBtnCancel))

function Update-GridTaskCombo { Update-TaskComboItems $script:gCmbTask ([string]$script:gCmbProject.Text) -DiscardText }

# 選択中セルの行番号一覧
function Get-SelectedGridRows {
    $idx = @{}
    foreach ($c in $script:grid.SelectedCells) { $idx[$c.RowIndex] = $true }
    return @($idx.Keys | Sort-Object)
}

# 表の1行から案件・作業内容を取り出す
function Get-GridRowValues($row) {
    return @{
        Project = ([string]$row.Cells[$COL_PROJ].Value).Trim()
        Task    = ([string]$row.Cells[$COL_TASK].Value).Trim()
        Time    = [string]$row.Cells[$COL_TIME].Value
    }
}

$gBtnApply.Add_Click({
    $p = [string]$script:gCmbProject.Text
    $w = ([string]$script:gCmbTask.Text).Trim()
    if (-not $p -or -not $w) {
        Show-Info '一括入力する案件と作業内容を選んでください。' '一覧編集'
        return
    }
    $rows = @(Get-SelectedGridRows)
    if ($rows.Count -eq 0) {
        Show-Info '表で入力したい行（セル）を選択してから押してください。' '一覧編集'
        return
    }
    foreach ($i in $rows) {
        $script:grid.Rows[$i].Cells[$COL_PROJ].Value = $p
        $script:grid.Rows[$i].Cells[$COL_TASK].Value = $w
    }
})

$gBtnClear.Add_Click({
    foreach ($i in @(Get-SelectedGridRows)) {
        $script:grid.Rows[$i].Cells[$COL_PROJ].Value = ''
        $script:grid.Rows[$i].Cells[$COL_TASK].Value = ''
    }
})

$gBtnCopy.Add_Click({
    # いまクリックしている行の案件・作業内容を一括入力欄へ写す。
    # あとは複製先の行を選んで「選択した行に入力」を押すだけで、ドロップダウン操作なしで複製できる
    $cell = $script:grid.CurrentCell
    if (-not $cell) {
        Show-Info 'コピーしたい行（セル）をクリックしてから押してください。' '一覧編集'
        return
    }
    $v = Get-GridRowValues $script:grid.Rows[$cell.RowIndex]
    if (-not $v.Project -and -not $v.Task) {
        Show-Info 'その行はまだ未入力です。内容のある行を選んでください。' '一覧編集'
        return
    }
    if ($v.Project) {
        # CSV由来で案件.txtに無い値でも選べるよう、無ければ選択肢に足してから選ぶ
        if (-not $script:gCmbProject.Items.Contains($v.Project)) { [void]$script:gCmbProject.Items.Add($v.Project) }
        $script:gCmbProject.SelectedItem = $v.Project   # 選択変更で作業名の選択肢も切り替わる
    }
    $script:gCmbTask.Text = $v.Task
})

$gBtnSave.Add_Click({
    [void]$script:grid.EndEdit()
    $dateStr = $script:gridDate.ToString($DATE_FMT)
    # 表の内容を行データに変換（案件・作業内容の片方だけ入力されている行はエラー）
    $newRows = @()
    foreach ($row in $script:grid.Rows) {
        $v = Get-GridRowValues $row
        if (-not $v.Project -and -not $v.Task) { continue }
        if (-not $v.Project -or -not $v.Task) {
            Show-Info "$($v.Time) の行は案件と作業内容の両方を入力してください。" '一覧編集'
            return
        }
        $newRows += [pscustomobject]@{ '日付' = $dateStr; '時刻' = $v.Time; '案件' = $v.Project; '作業内容' = $v.Task }
    }
    # 表の範囲外の行（先行記録した未来分など）はそのまま残す。
    # 別の日の行・読めない行も触らずに残す（表に出ていないものを保存で消さない）
    $keep = @()
    foreach ($r in @(Read-LogRows (Get-LogPath $script:gridDate))) {
        $rt = ConvertTo-SlotTimeOrNull ([string]$r.'日付') ([string]$r.'時刻')
        if (-not $rt -or $rt.Date -ne $script:gridDate) { $keep += $r; continue }
        if ($rt -lt $script:gridRangeStart -or $rt -ge $script:gridRangeEnd) { $keep += $r }
    }
    try {
        Write-DayFile (Get-LogPath $script:gridDate) (@($keep) + $newRows)
    } catch {
        Show-CsvError $_ '' '一覧編集'
        return
    }
    # 新しい作業名は作業名.txtの該当案件セクションへ（同じ組み合わせは1回だけ）
    $done = @{}
    foreach ($nr in $newRows) {
        $k = "$($nr.'案件')|$($nr.'作業内容')"
        if ($done[$k]) { continue }
        $done[$k] = $true
        Register-TaskIfNew $nr.'案件' $nr.'作業内容'
    }
    # 記録済みだった未来の枠（MTG予約など）をクリアした場合は、その枠から入力ウィンドウを
    # 再開できるよう記録開始位置を巻き戻す（未入力のまま残していた枠は対象外。当日の編集のみ）
    if ($script:gridDate -eq (Get-Date).Date) {
        $nowSlot = Get-FloorSlot (Get-Date)
        foreach ($row in $script:grid.Rows) {
            $v = Get-GridRowValues $row
            $rt = ConvertTo-SlotTime $dateStr $v.Time
            if ($rt -lt $nowSlot) { continue }
            if (-not $v.Project -and -not $v.Task -and $script:gridPrevData[$v.Time]) {
                if ($rt -lt $script:pendingStart) { $script:pendingStart = $rt }
                break
            }
        }
    }
    $script:gridForm.DialogResult = [Windows.Forms.DialogResult]::OK
})

# 表に出す時間帯を決める。
#   開始: 7:00（それより早い記録があればそこ）
#   終了: 当日は $pendingStart、過去日は19:00。記録済みの枠はすべて含める（予約の修正・取り消し用）。
#         業務終了モードでは作業中のいまの枠の終わりまで広げ、終業間際の最後の枠も入力できるようにする
function Get-GridRange([datetime]$date, $byTime, [bool]$isToday, [bool]$closing) {
    $start = $date.Date.AddHours(7)
    foreach ($rt in @($byTime.Keys)) {
        if ($rt -lt $start) { $start = $rt }
    }
    if ($isToday) { $end = $script:pendingStart } else { $end = $date.Date.AddHours(19) }
    foreach ($rt in @($byTime.Keys)) {
        if ($rt.AddMinutes($SLOT_MINUTES) -gt $end) { $end = $rt.AddMinutes($SLOT_MINUTES) }
    }
    if ($closing) {
        $slotEnd = (Get-FloorSlot (Get-Date)).AddMinutes($SLOT_MINUTES)
        if ($slotEnd -gt $end) { $end = $slotEnd }
    }
    return @{ Start = $start; End = $end }
}

function Show-GridDialog([datetime]$Date = (Get-Date), [switch]$Closing) {
    $isToday = ($Date.Date -eq (Get-Date).Date)
    $dateStr = $Date.ToString($DATE_FMT)
    $script:gridDate = $Date.Date
    # その日の行だけを時刻をキーに拾う（読めない行はここでは扱わず、保存時にそのまま残す）
    $rows = @()
    $byTime = @{}
    foreach ($r in @(Read-LogRows (Get-LogPath $Date))) {
        $rt = ConvertTo-SlotTimeOrNull ([string]$r.'日付') ([string]$r.'時刻')
        if (-not $rt -or $rt.Date -ne $Date.Date) { continue }
        $rows += $r
        $byTime[$rt] = $r
    }

    $range = Get-GridRange $Date $byTime $isToday $Closing.IsPresent
    if ($range.End -le $range.Start) {
        if (-not $Closing) { Show-Info 'まだ編集できる時間帯がありません。' '一覧編集' }
        return [Windows.Forms.DialogResult]::None
    }
    $script:gridRangeStart = $range.Start
    $script:gridRangeEnd   = $range.End

    # 案件列の選択肢（CSV内の値も含めて表示エラーを防ぐ）
    $projects = @(Get-ProjectList)
    $script:colProj.Items.Clear()
    [void]$script:colProj.Items.Add('')
    foreach ($p in $projects) { [void]$script:colProj.Items.Add($p) }
    foreach ($r in $rows) {
        $v = [string]$r.'案件'
        if ($v -and -not $script:colProj.Items.Contains($v)) { [void]$script:colProj.Items.Add($v) }
    }
    # 一括入力パネル
    Update-ProjectComboItems $script:gCmbProject $projects
    $script:gCmbTask.Items.Clear()
    $script:gCmbTask.Text = ''
    # 行を構築。記録済みだった時刻を控えておき、保存時に「予約の取り消し」を検出できるようにする
    $script:gridPrevData = @{}
    $script:grid.Rows.Clear()
    for ($t = $range.Start; $t -lt $range.End; $t = $t.AddMinutes($SLOT_MINUTES)) {
        $p = ''; $w = ''
        if ($byTime.ContainsKey($t)) {
            $p = [string]$byTime[$t].'案件'
            $w = [string]$byTime[$t].'作業内容'
            $script:gridPrevData[$t.ToString($TIME_FMT)] = $true
        }
        [void]$script:grid.Rows.Add(@($t.ToString($TIME_FMT), $p, $w))
    }
    if ($Closing) {
        $script:gridForm.Text = '業務終了（今日の記録の確認）'
        $script:gBtnSave.Text = '締め（保存して業務終了）'
    } elseif ($isToday) {
        $script:gridForm.Text = '一覧編集（今日の記録）'
        $script:gBtnSave.Text = '保存して閉じる'
    } else {
        $script:gridForm.Text = ('一覧編集（{0}）' -f $dateStr)
        $script:gBtnSave.Text = '保存して閉じる'
    }
    $script:gridForm.DialogResult = [Windows.Forms.DialogResult]::None
    return $script:gridForm.ShowDialog()
}

# 業務終了フロー: 一日の記録を表で見せ、「締め（保存して業務終了）」で確定する。
# 戻り値: $true = 締めた（アプリを終了してよい）、$false = キャンセル（通常運転を続ける）
function Invoke-WorkdayClose {
    $r = Show-GridDialog -Closing
    if ($r -eq [Windows.Forms.DialogResult]::None) {
        # 表示できる記録がまだ無い日は終了確認だけ出す
        $c = [Windows.Forms.MessageBox]::Show('今日の記録はまだありません。日報入力を終了しますか？',
            '日報入力', [Windows.Forms.MessageBoxButtons]::YesNo)
        return ($c -eq [Windows.Forms.DialogResult]::Yes)
    }
    return ($r -eq [Windows.Forms.DialogResult]::OK)
}

# 業務終了の共通入口（機能ID 'Close'）。締めが確定したらメインループを終了させる
function Invoke-ExitFlow {
    if ($script:gridForm.Visible) {
        Show-Info '一覧編集を閉じてから業務終了してください。'
        return
    }
    if (Invoke-WorkdayClose) {
        $script:trayExit = $true
        # 入力ウィンドウ・休憩中ウィンドウが開いたままなら閉じてメインループを抜けさせる
        # （休憩中の枠は締めの表で確認済みのため、Abortで閉じて後から書き足さない）
        if ($script:form.Visible) { $script:form.DialogResult = [Windows.Forms.DialogResult]::Abort }
        if ($script:breakForm.Visible) { $script:breakForm.DialogResult = [Windows.Forms.DialogResult]::Abort }
    }
}

# ============================================================
# 11. MTG予約ウィンドウ（先の時間帯を事前に記録する）
# ============================================================
$rsvForm = New-UiForm 'MTG予約（事前入力）' 500 270

$rsvLbl = New-UiLabel 15 12 '先の時間帯の作業（MTGなど）をあらかじめ記録します。予約した時間帯は15分ごとの入力ウィンドウが出ません。' 470 40

$lblRsvStart = New-UiLabel 15 62 '開始:'
$cmbRsvStart = New-UiCombo 65 58 85

$lblRsvDur    = New-UiLabel 175 62 '時間:'
$cmbRsvDur    = New-UiDurationCombo 225 58 '60' $RESERVE_DURATIONS
$lblRsvDurMin = New-UiLabel 305 62 '分'

$lblRsvP = New-UiLabel 15 96 '案件（ブランド）:'

$cmbRsvProject = New-UiCombo 15 118 220
$cmbRsvProject.Add_SelectedIndexChanged({ Update-RsvTaskCombo })

$lblRsvT = New-UiLabel 15 153 '作業内容:'

$cmbRsvTask = New-UiCombo 15 175 470 -Editable

$btnRsvOK     = New-UiButton 15 220 150 34 '予約する'
$btnRsvCancel = New-UiButton 335 220 150 34 'キャンセル'
$btnRsvCancel.DialogResult = [Windows.Forms.DialogResult]::Cancel

$rsvForm.Controls.AddRange(@($rsvLbl, $lblRsvStart, $cmbRsvStart, $lblRsvDur, $cmbRsvDur, $lblRsvDurMin, $lblRsvP, $cmbRsvProject, $lblRsvT, $cmbRsvTask, $btnRsvOK, $btnRsvCancel))
$rsvForm.AcceptButton = $btnRsvOK

function Update-RsvTaskCombo { Update-TaskComboItems $script:cmbRsvTask ([string]$script:cmbRsvProject.Text) }

$btnRsvOK.Add_Click({
    $p = $script:cmbRsvProject.Text.Trim()
    $t = $script:cmbRsvTask.Text.Trim()
    if (-not $script:cmbRsvStart.Text -or -not $p -or -not $t) {
        Show-Info '開始時刻・案件・作業内容を入力してください。' 'MTG予約'
        return
    }
    $start = ConvertTo-SlotTime ((Get-Date).ToString($DATE_FMT)) ([string]$script:cmbRsvStart.Text)
    $end = $start.AddMinutes((Get-ComboMinutes $script:cmbRsvDur 60))
    try {
        $n = Write-Slots $start $end $p $t
    } catch {
        Show-CsvError $_ 'もう一度お試しください。' 'MTG予約'
        return
    }
    if ($n -eq 0) {
        Show-Info 'その時間帯はすべて記録済みです。修正する場合は「一覧編集」を使ってください。' 'MTG予約'
        return
    }
    Register-TaskIfNew $p $t
    Show-Info ('{0:H:mm}～{1:H:mm} に「{2}」を予約しました。この時間帯の入力ウィンドウは出ません。' -f $start, $end, $t) 'MTG予約'
    $script:rsvForm.DialogResult = [Windows.Forms.DialogResult]::OK
})

function Show-ReserveDialog {
    # どの入口から呼ばれても同じガード（表示中の再入・一覧編集との同時表示を防ぐ）
    if ($script:rsvForm.Visible) { return }
    if ($script:gridForm.Visible) {
        Show-Info '一覧編集を閉じてからMTG予約してください。'
        return
    }
    Update-ProjectComboItems $script:cmbRsvProject @(Get-ProjectList) $script:lastProject
    $script:cmbRsvTask.Text = ''
    Update-RsvTaskCombo
    # 開始時刻の選択肢はいまの枠から当日末まで。既定はMTGで多い「次の枠」
    $script:cmbRsvStart.Items.Clear()
    $s = Get-FloorSlot (Get-Date)
    while ($s.Date -eq (Get-Date).Date) {
        [void]$script:cmbRsvStart.Items.Add($s.ToString($TIME_FMT))
        $s = $s.AddMinutes($SLOT_MINUTES)
    }
    $script:cmbRsvStart.SelectedIndex = [Math]::Min(1, $script:cmbRsvStart.Items.Count - 1)
    $script:cmbRsvDur.SelectedItem = '60'
    $script:rsvForm.DialogResult = [Windows.Forms.DialogResult]::None
    [void]$script:rsvForm.ShowDialog()
}

# ============================================================
# 12. 設定メンテナンスウィンドウ（案件・作業内容の選択肢を画面から編集する）
# ============================================================
# 左が案件（案件.txt）、右が選択中の案件の作業内容（作業名.txt の [案件] セクション）。
# 左の一覧の先頭にある「（共通）」は見出しより前に書く行＝どの案件でも選択肢に出る作業内容。
# 編集はいったんメモリ上のモデル（Read-ProjectFileModel / Read-TaskFileModel）に対して行い、
# 「保存して閉じる」で2つのファイルをまとめて書き戻す。
# **過去の日報CSVには手を触れない**（記録は当時の表記のまま残す）ので、名前変更しても
# 変わるのは次回からの選択肢だけ。集計の名寄せが要るときは一覧編集で直す。

$MAINT_COMMON_LABEL = '（共通：すべての案件で表示）'
$MAINT_ORPHAN_MARK  = '（案件一覧に無い）'

$maintForm = New-UiForm '選択肢の編集（案件・作業内容）' 720 512

$mLblP = New-UiLabel 12 10 '案件（ブランド）'
$mLblT = New-UiLabel 300 10 '作業内容'

$mLstProject = New-Object Windows.Forms.ListBox
$mLstProject.Location = New-Object Drawing.Point(12, 34)
$mLstProject.Size     = New-Object Drawing.Size(270, 330)
$mLstProject.Font     = $UiFont

$mLstTask = New-Object Windows.Forms.ListBox
$mLstTask.Location = New-Object Drawing.Point(300, 34)
$mLstTask.Size     = New-Object Drawing.Size(408, 330)
$mLstTask.Font     = $UiFont

$mBtnPAdd  = New-UiButton  12 372  86 30 '追加'
$mBtnPRen  = New-UiButton 104 372  86 30 '名前変更'
$mBtnPDel  = New-UiButton 196 372  86 30 '削除'
$mBtnPUp   = New-UiButton  12 406  86 30 '↑ 上へ'
$mBtnPDown = New-UiButton 104 406  86 30 '↓ 下へ'

$mBtnTAdd  = New-UiButton 300 372  90 30 '追加'
$mBtnTRen  = New-UiButton 396 372  90 30 '名前変更'
$mBtnTDel  = New-UiButton 492 372  90 30 '削除'
$mBtnTMove = New-UiButton 588 372 120 30 '別の案件へ移動'
$mBtnTUp   = New-UiButton 300 406  90 30 '↑ 上へ'
$mBtnTDown = New-UiButton 396 406  90 30 '↓ 下へ'

$mLblNote = New-UiLabel 492 408 '※ 過去の記録は書き換えません' 216 26

$mBtnSave   = New-UiButton  12 448 230 36 '保存して閉じる'
$mBtnCancel = New-UiButton 568 448 140 36 'キャンセル'

$maintForm.Controls.AddRange(@($mLblP, $mLblT, $mLstProject, $mLstTask,
    $mBtnPAdd, $mBtnPRen, $mBtnPDel, $mBtnPUp, $mBtnPDown,
    $mBtnTAdd, $mBtnTRen, $mBtnTDel, $mBtnTMove, $mBtnTUp, $mBtnTDown,
    $mLblNote, $mBtnSave, $mBtnCancel))

# 名前として使えるか。使えなければ理由（表示用の文字列）、問題なければ '' を返す。
# コメント行・見出し行と読み間違えられる書き方だけを弾く（設定ファイルの書式の制約）
function Test-MaintName([string]$name, [switch]$AsProject) {
    if (-not $name) { return '名前を入力してください。' }
    if ($name.StartsWith('#')) { return '#で始まる名前は使えません（コメント行として無視されるため）。' }
    if ($AsProject) {
        if ($name.Contains('[') -or $name.Contains(']')) { return '案件名に [ ] は使えません（作業内容の見出しに使う記号のため）。' }
    } elseif ($name -match '^\[.+\]$') {
        return '[ ] だけで囲んだ名前は使えません（案件の見出しとして読まれるため）。'
    }
    return ''
}

# 左の一覧で選ばれている案件のキー（共通は ''）。未選択なら $null
function Get-MaintKey {
    $i = $script:mLstProject.SelectedIndex
    if ($i -lt 0 -or $i -ge $script:maintKeys.Count) { return $null }
    return $script:maintKeys[$i]
}

# 作業名モデルに、その案件のセクションが無ければ作る
function Add-MaintSection([string]$key) {
    if ($script:maintTask.Keys.Contains($key)) { return }
    $script:maintTask.Keys.Add($key)
    $script:maintTask.Items[$key]    = New-Object Collections.Generic.List[string]
    $script:maintTask.Comments[$key] = New-Object Collections.Generic.List[string]
}

# 左の一覧を作り直す。案件.txt の順に並べ、作業名.txt にしか無いセクション（案件から消した
# 名残など）は末尾に印つきで出す（画面に出ないと直せなくなるため）
function Update-MaintProjectList([string]$selectKey = '') {
    $script:maintKeys = New-Object Collections.Generic.List[string]
    $script:mLstProject.Items.Clear()
    $script:maintKeys.Add('')
    [void]$script:mLstProject.Items.Add($MAINT_COMMON_LABEL)
    foreach ($p in $script:maintProj.Items) {
        $script:maintKeys.Add($p)
        [void]$script:mLstProject.Items.Add($p)
    }
    foreach ($k in $script:maintTask.Keys) {
        if (-not $k -or $script:maintProj.Items.Contains($k)) { continue }
        $script:maintKeys.Add($k)
        [void]$script:mLstProject.Items.Add("$k $MAINT_ORPHAN_MARK")
    }
    $idx = $script:maintKeys.IndexOf($selectKey)
    if ($idx -lt 0) { $idx = 0 }
    $script:mLstProject.SelectedIndex = $idx
}

# 右の一覧を、選択中の案件の作業内容で作り直す
function Update-MaintTaskList([string]$selectTask = '') {
    $key = Get-MaintKey
    $script:mLstTask.Items.Clear()
    if ($null -eq $key) { $script:mLblT.Text = '作業内容'; return }
    Add-MaintSection $key
    foreach ($t in $script:maintTask.Items[$key]) { [void]$script:mLstTask.Items.Add($t) }
    $script:mLblT.Text = $(if ($key) { "作業内容（$key）" } else { '作業内容（共通：すべての案件で表示）' })
    if ($selectTask) { $script:mLstTask.SelectedItem = $selectTask }
    elseif ($script:mLstTask.Items.Count -gt 0) { $script:mLstTask.SelectedIndex = 0 }
}

# 一覧の中の1件を上下に動かす共通処理。動かせたら動かした値、動かせなければ $null を返す
# （画面の行番号ではなく値で探す。左の一覧は「共通」の分だけ行番号がずれるため）
function Move-MaintItem([Collections.Generic.List[string]]$items, [string]$value, [int]$delta) {
    if (-not $value) { return $null }
    $i = $items.IndexOf($value)
    if ($i -lt 0) { return $null }
    $j = $i + $delta
    if ($j -lt 0 -or $j -ge $items.Count) { return $null }
    $v = $items[$i]
    $items.RemoveAt($i)
    $items.Insert($j, $v)
    $script:maintDirty = $true
    return $v
}

$mLstProject.Add_SelectedIndexChanged({ Update-MaintTaskList })

# ---- 案件の操作 ----
$mBtnPAdd.Add_Click({
    $name = Show-TextPrompt '案件の追加' '追加する案件（ブランド）の名前を入力してください。' ''
    if (-not $name) { return }
    $err = Test-MaintName $name -AsProject
    if ($err) { Show-Info $err '選択肢の編集'; return }
    if ($script:maintProj.Items.Contains($name)) { Show-Info '同じ名前の案件が既にあります。' '選択肢の編集'; return }
    $script:maintProj.Items.Add($name)
    Add-MaintSection $name
    $script:maintDirty = $true
    Update-MaintProjectList $name
})

$mBtnPRen.Add_Click({
    $key = Get-MaintKey
    if (-not $key) { Show-Info '名前を変更する案件を選んでください（「共通」は名前を変えられません）。' '選択肢の編集'; return }
    $name = Show-TextPrompt '案件の名前変更' "「$key」の新しい名前を入力してください。`n過去の記録（CSV）は書き換えません。" $key
    if (-not $name -or $name -eq $key) { return }
    $err = Test-MaintName $name -AsProject
    if ($err) { Show-Info $err '選択肢の編集'; return }
    if ($script:maintProj.Items.Contains($name) -or $script:maintTask.Keys.Contains($name)) {
        Show-Info '同じ名前の案件（または作業内容の見出し）が既にあります。' '選択肢の編集'
        return
    }
    $i = $script:maintProj.Items.IndexOf($key)
    if ($i -ge 0) { $script:maintProj.Items[$i] = $name }
    # 作業内容のセクションも同じ名前に付け替える（並び順は保つ）
    $k = $script:maintTask.Keys.IndexOf($key)
    if ($k -ge 0) {
        $script:maintTask.Keys[$k] = $name
        $script:maintTask.Items[$name]    = $script:maintTask.Items[$key]
        $script:maintTask.Comments[$name] = $script:maintTask.Comments[$key]
        $script:maintTask.Items.Remove($key)
        $script:maintTask.Comments.Remove($key)
    }
    $script:maintDirty = $true
    Update-MaintProjectList $name
})

$mBtnPDel.Add_Click({
    $key = Get-MaintKey
    if (-not $key) { Show-Info '削除する案件を選んでください（「共通」は削除できません）。' '選択肢の編集'; return }
    $n = 0
    if ($script:maintTask.Keys.Contains($key)) { $n = $script:maintTask.Items[$key].Count }
    $msg = "案件「$key」を選択肢から削除します。"
    if ($n -gt 0) { $msg += "`nこの案件の作業内容 $n 件も一緒に削除されます。" }
    $msg += "`n過去の記録（CSV）はそのまま残ります。よろしいですか？"
    if (-not (Test-Confirm $msg '選択肢の編集')) { return }
    [void]$script:maintProj.Items.Remove($key)
    if ($script:maintTask.Keys.Contains($key)) {
        [void]$script:maintTask.Keys.Remove($key)
        $script:maintTask.Items.Remove($key)
        $script:maintTask.Comments.Remove($key)
    }
    $script:maintDirty = $true
    Update-MaintProjectList
})

$mBtnPUp.Add_Click({
    $v = Move-MaintItem $script:maintProj.Items ([string](Get-MaintKey)) -1
    if ($v) { Update-MaintProjectList $v }
})

$mBtnPDown.Add_Click({
    $v = Move-MaintItem $script:maintProj.Items ([string](Get-MaintKey)) 1
    if ($v) { Update-MaintProjectList $v }
})

# ---- 作業内容の操作 ----
$mBtnTAdd.Add_Click({
    $key = Get-MaintKey
    if ($null -eq $key) { Show-Info '先に左の一覧で案件を選んでください。' '選択肢の編集'; return }
    $name = Show-TextPrompt '作業内容の追加' '追加する作業内容を入力してください。' ''
    if (-not $name) { return }
    $err = Test-MaintName $name
    if ($err) { Show-Info $err '選択肢の編集'; return }
    if ($script:maintTask.Items[$key].Contains($name)) { Show-Info '同じ作業内容が既にあります。' '選択肢の編集'; return }
    $script:maintTask.Items[$key].Add($name)
    $script:maintDirty = $true
    Update-MaintTaskList $name
})

$mBtnTRen.Add_Click({
    $key = Get-MaintKey
    $cur = [string]$script:mLstTask.SelectedItem
    if ($null -eq $key -or -not $cur) { Show-Info '名前を変更する作業内容を選んでください。' '選択肢の編集'; return }
    $name = Show-TextPrompt '作業内容の名前変更' "「$cur」の新しい名前を入力してください。`n過去の記録（CSV）は書き換えません。" $cur
    if (-not $name -or $name -eq $cur) { return }
    $err = Test-MaintName $name
    if ($err) { Show-Info $err '選択肢の編集'; return }
    if ($script:maintTask.Items[$key].Contains($name)) { Show-Info '同じ作業内容が既にあります。' '選択肢の編集'; return }
    $script:maintTask.Items[$key][$script:maintTask.Items[$key].IndexOf($cur)] = $name
    $script:maintDirty = $true
    Update-MaintTaskList $name
})

$mBtnTDel.Add_Click({
    $key = Get-MaintKey
    $cur = [string]$script:mLstTask.SelectedItem
    if ($null -eq $key -or -not $cur) { Show-Info '削除する作業内容を選んでください。' '選択肢の編集'; return }
    if (-not (Test-Confirm "作業内容「$cur」を選択肢から削除します。`n過去の記録（CSV）はそのまま残ります。よろしいですか？" '選択肢の編集')) { return }
    [void]$script:maintTask.Items[$key].Remove($cur)
    $script:maintDirty = $true
    Update-MaintTaskList
})

$mBtnTMove.Add_Click({
    $key = Get-MaintKey
    $cur = [string]$script:mLstTask.SelectedItem
    if ($null -eq $key -or -not $cur) { Show-Info '移動する作業内容を選んでください。' '選択肢の編集'; return }
    # 移動先の候補は「共通」＋自分以外の案件
    $labels = New-Object Collections.Generic.List[string]
    $keys   = New-Object Collections.Generic.List[string]
    for ($i = 0; $i -lt $script:maintKeys.Count; $i++) {
        if ($script:maintKeys[$i] -eq $key) { continue }
        $keys.Add($script:maintKeys[$i])
        $labels.Add([string]$script:mLstProject.Items[$i])
    }
    if ($keys.Count -eq 0) { Show-Info '移動先の案件がありません。' '選択肢の編集'; return }
    $sel = Show-ChoicePrompt '別の案件へ移動' "「$cur」の移動先を選んでください。" @($labels.ToArray())
    if (-not $sel) { return }
    $dest = $keys[$labels.IndexOf($sel)]
    Add-MaintSection $dest
    if ($script:maintTask.Items[$dest].Contains($cur)) {
        Show-Info '移動先に同じ作業内容が既にあります。' '選択肢の編集'
        return
    }
    [void]$script:maintTask.Items[$key].Remove($cur)
    $script:maintTask.Items[$dest].Add($cur)
    $script:maintDirty = $true
    Update-MaintTaskList
})

$mBtnTUp.Add_Click({
    $key = Get-MaintKey
    if ($null -eq $key) { return }
    $v = Move-MaintItem $script:maintTask.Items[$key] ([string]$script:mLstTask.SelectedItem) -1
    if ($v) { Update-MaintTaskList $v }
})

$mBtnTDown.Add_Click({
    $key = Get-MaintKey
    if ($null -eq $key) { return }
    $v = Move-MaintItem $script:maintTask.Items[$key] ([string]$script:mLstTask.SelectedItem) 1
    if ($v) { Update-MaintTaskList $v }
})

# ---- 保存・キャンセル ----
$mBtnSave.Add_Click({
    if ($script:maintProj.Items.Count -eq 0) {
        Show-Info '案件が1つもありません。1つ以上残してください。' '選択肢の編集'
        return
    }
    try {
        Write-ProjectFileModel $script:maintProj
        Write-TaskFileModel $script:maintTask
    } catch {
        Show-Info ("設定ファイルに書き込めませんでした。メモ帳などで開いていたら閉じてください。`n`n" +
                   $_.Exception.Message) '選択肢の編集'
        return
    }
    $script:maintDirty = $false
    $script:maintForm.DialogResult = [Windows.Forms.DialogResult]::OK
})

$mBtnCancel.Add_Click({
    if ($script:maintDirty -and -not (Test-Confirm '編集した内容は保存されません。閉じてよろしいですか？' '選択肢の編集')) { return }
    $script:maintDirty = $false
    $script:maintForm.DialogResult = [Windows.Forms.DialogResult]::Cancel
})

# ×で閉じたときもキャンセルと同じ確認を出す
$maintForm.Add_FormClosing({ param($s, $e)
    if ($e.CloseReason -ne [Windows.Forms.CloseReason]::UserClosing) { return }
    if ($script:maintDirty -and -not (Test-Confirm '編集した内容は保存されません。閉じてよろしいですか？' '選択肢の編集')) {
        $e.Cancel = $true
    }
})

# 設定メンテナンス（機能ID 'Maint'）。開くたびにファイルを読み直す
function Show-MaintDialog {
    if ($script:maintForm.Visible) { return }
    if ($script:gridForm.Visible) {
        Show-Info '一覧編集を閉じてから選択肢の編集をしてください。'
        return
    }
    $script:maintProj  = Read-ProjectFileModel
    $script:maintTask  = Read-TaskFileModel
    $script:maintDirty = $false
    # 案件.txt が空（＝既定の選択肢で動いている）状態から編集できるように、既定値を入れておく
    if ($script:maintProj.Items.Count -eq 0) {
        foreach ($p in $DEFAULT_PROJECTS) { $script:maintProj.Items.Add($p) }
    }
    Update-MaintProjectList $script:lastProject
    Update-MaintTaskList
    $script:maintForm.DialogResult = [Windows.Forms.DialogResult]::None
    [void]$script:maintForm.ShowDialog()
}

# ============================================================
# 13. 日報メニュー（常時表示の小窓）
# ============================================================
$hubForm = New-UiForm '日報メニュー' 320 234 -Resident

$lblHubDate = New-UiLabel 15 16 '日付:'

# 選択肢は日報ログがある日付のみ。ドロップダウンを開くたびに読み直す
$cmbHubDate = New-UiCombo 62 12 130
$cmbHubDate.Add_DropDown({ Update-HubDates })

$btnHubEdit  = New-UiButton 202 11 100 30 '編集'
$btnHubRsv   = New-UiButton 15 52 287 34 (Get-FeatureLabel 'Reserve')
$btnHubBreak = New-UiButton 15 96 287 34 (Get-FeatureLabel 'Break')
$btnHubMaint = New-UiButton 15 140 287 34 (Get-FeatureLabel 'Maint')
$btnHubClose = New-UiButton 15 184 287 34 (Get-FeatureLabel 'Close')

$hubForm.Controls.AddRange(@($lblHubDate, $cmbHubDate, $btnHubEdit, $btnHubRsv, $btnHubBreak, $btnHubMaint, $btnHubClose))

function Update-HubDates {
    $sel = [string]$script:cmbHubDate.Text
    $dates = @(Get-LogDates)
    $script:cmbHubDate.Items.Clear()
    if ($dates.Count -gt 0) { $script:cmbHubDate.Items.AddRange($dates) }
    if ($sel -and $dates -contains $sel) { $script:cmbHubDate.SelectedItem = $sel }
    elseif ($dates.Count -gt 0) { $script:cmbHubDate.SelectedIndex = 0 }
}

$btnHubEdit.Add_Click({
    $d = [string]$script:cmbHubDate.Text
    if (-not $d) {
        Update-HubDates
        $d = [string]$script:cmbHubDate.Text
        if (-not $d) {
            Show-Info '日報ログのある日付がまだありません。' '日報メニュー'
            return
        }
    }
    Invoke-Feature 'Edit' ([datetime]::ParseExact($d, 'yyyy-MM-dd', $null))
})

$btnHubRsv.Add_Click({ Invoke-Feature 'Reserve' })
$btnHubBreak.Add_Click({ Invoke-Feature 'Break' })
$btnHubMaint.Add_Click({ Invoke-Feature 'Maint' })
$btnHubClose.Add_Click({ Invoke-Feature 'Close' })

# ×で閉じてもアプリは終了せず隠すだけ（トレイの「日報メニューを表示」で再表示できる）
$hubForm.Add_FormClosing({ param($s, $e)
    if ($e.CloseReason -eq [Windows.Forms.CloseReason]::UserClosing) {
        $e.Cancel = $true
        $s.Hide()
    }
})

# 全ウィンドウの破棄（終了時・テスト後）
function Close-AllForms {
    foreach ($f in @($script:form, $script:breakForm, $script:gridForm, $script:rsvForm, $script:maintForm, $script:hubForm)) {
        if ($f) { $f.Dispose() }
    }
}

# ============================================================
# 14. セルフテスト（-TestInit。ウィンドウを出さずに主要処理を検証する）
# ============================================================
# 実データは壊さず、一時フォルダに作った（またはコピーした）ファイルに対して書き込みを試す。
# 各 Test-* は検証結果を返し、Invoke-SelfTest がまとめてJSONで出力する。
# 出力の *Ok がすべて true なら主要処理は健全

# 記録の追記: 引用符・カンマ入りの値が読み戻せるか。予約と重なった枠を二重に書かないか
function Test-SlotWriting {
    $from = (Get-Date).Date.AddHours(9)
    $written = Write-Slots $from $from.AddMinutes(45) '案件A' 'テスト用,"引用符"とカンマ入り'   # 3枠
    $back = @(Read-LogRows (Get-LogPath $from))
    # 9:30は既に書いた枠なので、9:30～10:00を指定しても9:45の1枠だけが増える
    $again = Write-Slots $from.AddMinutes(30) $from.AddMinutes(60) '案件A' '重複テスト'
    return @{
        Written  = $written
        ReadBack = $back.Count
        LastRow  = ('{0} {1} | {2} | {3}' -f $back[-1].'日付', $back[-1].'時刻', $back[-1].'案件', $back[-1].'作業内容')
        DedupOk  = ($again -eq 1) -and (@(Read-LogRows (Get-LogPath $from)).Count -eq 4)
    }
}

# 一覧編集の保存: 順序がばらばらの行が時刻順に並べ直されるか
function Test-GridSave {
    $today = (Get-Date).ToString($DATE_FMT)
    Write-DayFile (Get-LogPath (Get-Date)) @(
        [pscustomobject]@{ '日付' = $today; '時刻' = '9:15'; '案件' = '案件A'; '作業内容' = '一覧編集テストB' }
        [pscustomobject]@{ '日付' = $today; '時刻' = '8:00'; '案件' = '案件A'; '作業内容' = '一覧編集テストA' }
    )
    $rows = @(Read-LogRows (Get-LogPath (Get-Date)))
    return ($rows.Count -eq 2) -and ($rows[0].'時刻' -eq '8:00') -and ($rows[1].'時刻' -eq '9:15')
}

# ExcelでCSVを開いて保存すると 2026/8/7・9:15:00 のように表記が変わることがある。
# それでも記録済みと判定でき、書き直したときに元の書式へ戻るか
function Test-FormatTolerance {
    $d = (Get-Date).Date.AddDays(1)   # 他のテストと日付を分ける
    $path = Get-LogPath $d
    $loose = "日付,時刻,案件,作業内容`r`n" +
             ('"{0}/{1}/{2}","9:15:00","案件A","桁揃えなし"' -f $d.Year, $d.Month, $d.Day) + "`r`n"
    [IO.File]::WriteAllText($path, $loose, $Utf8Bom)
    $slot = $d.AddHours(9).AddMinutes(15)
    $dup = Write-Slots $slot $slot.AddMinutes(15) '案件A' '二重にならないはず'
    Write-DayFile $path @(Read-LogRows $path)
    $rows = @(Read-LogRows $path)
    return ($dup -eq 0) -and ($rows.Count -eq 1) -and
           ($rows[0].'日付' -eq $d.ToString($DATE_FMT)) -and ($rows[0].'時刻' -eq '9:15')
}

# 新しい作業名が、選択中の案件のセクションにだけ追加されるか
function Test-TaskRegistration([string]$sourceTaskList) {
    $script:taskListPath = Join-Path $script:logDir '作業名テスト.txt'
    Copy-Item -LiteralPath $sourceTaskList -Destination $script:taskListPath
    Add-TaskToFile '案件A' 'テスト新規作業ダミー'
    return ((Get-TasksForProject '案件A') -contains 'テスト新規作業ダミー') -and
           (-not ((Get-TasksForProject '案件B') -contains 'テスト新規作業ダミー'))
}

# 設定メンテナンスの保存: 画面での編集（案件の改名・作業内容の追加）がファイルに書き戻り、
# 本文中のコメント行が消えずに残るか
function Test-MaintSave([string]$sourceProjectList) {
    $script:projectListPath = Join-Path $script:logDir '案件テスト.txt'
    Copy-Item -LiteralPath $sourceProjectList -Destination $script:projectListPath
    # 作業名.txt は Test-TaskRegistration でコピー済みのものを使う
    $proj = Read-ProjectFileModel
    $task = Read-TaskFileModel
    $target = $null
    foreach ($p in $proj.Items) { if ($task.Keys.Contains($p)) { $target = $p; break } }
    if (-not $target) { return $false }
    $commentCount = $task.Comments[$target].Count
    $new = '改名テスト案件'
    $proj.Items[$proj.Items.IndexOf($target)] = $new
    $k = $task.Keys.IndexOf($target)
    $task.Keys[$k] = $new
    $task.Items[$new]    = $task.Items[$target]
    $task.Comments[$new] = $task.Comments[$target]
    $task.Items.Remove($target)
    $task.Comments.Remove($target)
    $task.Items[$new].Add('メンテナンステスト作業')
    $task.Items[''].Add('共通テスト作業')
    Write-ProjectFileModel $proj
    Write-TaskFileModel $task
    $back = Read-TaskFileModel
    return (@(Get-ProjectList) -contains $new) -and
           (-not (@(Get-ProjectList) -contains $target)) -and
           ((Get-TasksForProject $new) -contains 'メンテナンステスト作業') -and
           ((Get-TasksForProject $new) -contains '共通テスト作業') -and
           ($back.Comments[$new].Count -eq $commentCount)
}

function Invoke-SelfTest {
    # 実環境の状態（ここは読むだけ）
    $realLog             = Get-LogPath (Get-Date)
    $realTaskListPath    = $script:taskListPath
    $realProjectListPath = $script:projectListPath
    $projectCount = @(Get-ProjectList).Count
    $taskCount    = @(Get-TaskList).Count
    $byProject = (@(Get-ProjectList) | ForEach-Object { '{0}={1}' -f $_, @(Get-TasksForProject $_).Count }) -join ', '

    # ここから先は一時フォルダ上のログ・設定に差し替えて検証する
    $script:logDir = Join-Path $env:TEMP ('nippo-batch-test-' + [Guid]::NewGuid().ToString('N'))
    $tempDir = $script:logDir
    try {
        $slot   = Test-SlotWriting
        $grid   = Test-GridSave
        $format = Test-FormatTolerance
        $addOk  = Test-TaskRegistration $realTaskListPath   # 以降は作業名.txtのコピーを使う
        $maint  = Test-MaintSave $realProjectListPath       # 以降は案件.txtのコピーも使う
    } finally {
        $script:taskListPath    = $realTaskListPath
        $script:projectListPath = $realProjectListPath
    }

    [ordered]@{
        logPath         = $realLog
        projectListPath = $script:projectListPath
        taskListPath    = $script:taskListPath
        projectCount    = $projectCount
        taskCount       = $taskCount
        taskCountByProject = $byProject
        addToSectionOk  = $addOk
        maintSaveOk     = $maint
        gridSaveOk      = $grid
        dedupOk         = $slot.DedupOk
        formatToleranceOk = $format
        pendingStart    = $script:pendingStart.ToString($SLOT_FMT)
        lastProject     = $script:lastProject
        lastTask        = $script:lastTask
        testWritten     = $slot.Written
        testReadBack    = $slot.ReadBack
        testLastRow     = $slot.LastRow
    } | ConvertTo-Json
    Remove-Item -LiteralPath $tempDir -Recurse -Force
}

if ($TestInit) {
    Invoke-SelfTest
    Close-AllForms
    exit 0
}

# ============================================================
# 15. 画面キャプチャ（-Capture。マニュアル・プレゼン用のPNGを書き出す）
# ============================================================
# 実物のウィンドウをそのまま撮るので、UIを変えたら撮り直すだけで資料の画像が最新になる。
# 撮影中は画面にウィンドウが一瞬表示される（画面に出ていないと中身を撮れないため）

# 撮影用の無地の背景（丸角の外側にデスクトップが写り込まないように全面を覆う）
function New-CaptureBackdrop {
    $b = New-Object Windows.Forms.Form
    $b.FormBorderStyle = 'None'
    $b.WindowState     = 'Maximized'
    $b.BackColor       = [Drawing.Color]::FromArgb(246, 247, 249)
    $b.TopMost         = $true
    $b.ShowInTaskbar   = $false
    return $b
}

function Save-FormShot($form, [string]$path, $backdrop) {
    $wasTopMost = $form.TopMost
    $form.StartPosition = 'Manual'
    $form.Location = New-Object Drawing.Point(140, 140)
    $form.TopMost = $true
    $form.Show()
    $form.Activate()
    $form.BringToFront()
    for ($i = 0; $i -lt 8; $i++) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 80 }
    # 影の分だけ外側を広めに撮り、背景色で塗りつぶした上に重ねる
    $pad = 8
    $bmp = New-Object Drawing.Bitmap(($form.Width + $pad * 2), ($form.Height + $pad * 2))
    $g = [Drawing.Graphics]::FromImage($bmp)
    $origin = New-Object Drawing.Point(($form.Left - $pad), ($form.Top - $pad))
    $g.CopyFromScreen($origin, [Drawing.Point]::Empty, $bmp.Size)
    $g.Dispose()
    $bmp.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $form.Hide()
    $form.TopMost = $wasTopMost
    if ($backdrop) { $backdrop.Activate() }
    [Windows.Forms.Application]::DoEvents()
}

# 資料用のダミー値。実際の案件名・チケット名は社外にも出る資料に載せたくないので、
# 設定ファイルは読まずにここで固定した架空の値だけを使う
$SHOT_PROJECTS = @('案件A', '案件B', '社内', 'その他')
$SHOT_TASKS    = @('PRJ-12 商品ページ改修', 'API設計レビュー', '週次定例MTG', '資料作成')
$SHOT_PROJECT  = '案件A'
$SHOT_TASK     = 'PRJ-12 商品ページ改修'

# 案件・作業内容のコンボにダミー値を入れる（設定ファイルの読み込みイベントの後に上書きする）
function Set-ShotCombos($projectCombo, $taskCombo) {
    $projectCombo.Items.Clear()
    $projectCombo.Items.AddRange($SHOT_PROJECTS)
    $projectCombo.SelectedItem = $SHOT_PROJECT
    $taskCombo.Items.Clear()
    $taskCombo.Items.AddRange($SHOT_TASKS)
    $taskCombo.Text = $SHOT_TASK
}

function Invoke-CaptureShots {
    $outDir = Join-Path $reportDir '画像'
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $backdrop = New-CaptureBackdrop
    $backdrop.Show()
    [Windows.Forms.Application]::DoEvents()
    $slot = Get-FloorSlot (Get-Date)
    $saved = New-Object Collections.Generic.List[string]

    # 入力ウィンドウ
    Set-ShotCombos $script:cmbProject $script:cmbTask
    $script:cmbDur.SelectedItem = '15'
    $script:keepChk.Checked = $true
    $script:lblRange.Text = ('対象: {0:H:mm} ～ {1:H:mm}（1枠 × {2}分）' -f $slot, $slot.AddMinutes($SLOT_MINUTES), $SLOT_MINUTES)
    Save-FormShot $script:form (Join-Path $outDir '入力ウィンドウ.png') $backdrop
    $saved.Add('入力ウィンドウ.png')

    # 日報メニュー
    $script:cmbHubDate.Items.Clear()
    [void]$script:cmbHubDate.Items.Add((Get-Date).ToString('yyyy-MM-dd'))
    $script:cmbHubDate.SelectedIndex = 0
    Save-FormShot $script:hubForm (Join-Path $outDir '日報メニュー.png') $backdrop
    $saved.Add('日報メニュー.png')

    # 休憩中
    $script:breakClockStart = $slot
    $script:breakLbl.Text = ('休憩中（{0:H:mm} 開始・約{1}分経過）' -f $slot, 25)
    Save-FormShot $script:breakForm (Join-Path $outDir '休憩中.png') $backdrop
    $saved.Add('休憩中.png')

    # MTG予約
    Set-ShotCombos $script:cmbRsvProject $script:cmbRsvTask
    $script:cmbRsvStart.Items.Clear()
    $t0 = $slot
    while ($t0.Date -eq $slot.Date) { [void]$script:cmbRsvStart.Items.Add($t0.ToString($TIME_FMT)); $t0 = $t0.AddMinutes($SLOT_MINUTES) }
    $script:cmbRsvStart.SelectedIndex = [Math]::Min(1, $script:cmbRsvStart.Items.Count - 1)
    $script:cmbRsvDur.SelectedItem = '60'
    Save-FormShot $script:rsvForm (Join-Path $outDir 'MTG予約.png') $backdrop
    $saved.Add('MTG予約.png')

    # 一覧編集（ダミー行を流し込む。ShowDialogは使わずに見た目だけ作る）
    $script:colProj.Items.Clear()
    [void]$script:colProj.Items.Add('')
    foreach ($p in $SHOT_PROJECTS) { [void]$script:colProj.Items.Add($p) }
    $script:gCmbProject.Items.Clear()
    $script:gCmbProject.Items.AddRange($SHOT_PROJECTS)
    $script:gCmbProject.SelectedIndex = -1
    $script:gCmbTask.Items.Clear()
    $script:gCmbTask.Items.AddRange($SHOT_TASKS)
    $script:gCmbTask.Text = ''
    # 9:00から埋まっていて、途中に会議と休憩が入っている一日のイメージ
    $rows = [ordered]@{
        '9:00'  = @('案件A', 'API設計レビュー')
        '9:15'  = @('案件A', 'API設計レビュー')
        '9:30'  = @('案件A', 'PRJ-12 商品ページ改修')
        '9:45'  = @('案件A', 'PRJ-12 商品ページ改修')
        '10:00' = @('社内', '週次定例MTG')
        '10:15' = @('社内', '週次定例MTG')
        '10:30' = @('案件B', '運用保守 問い合わせ対応')
        '10:45' = @('案件B', '運用保守 問い合わせ対応')
        '11:00' = @('案件A', 'PRJ-12 商品ページ改修')
        '11:15' = @('案件A', 'PRJ-12 商品ページ改修')
        '11:30' = @('案件A', 'PRJ-12 商品ページ改修')
        '11:45' = @('案件A', 'PRJ-12 商品ページ改修')
        '12:00' = @('その他', '休憩')
        '12:15' = @('その他', '休憩')
        '12:30' = @('その他', '休憩')
        '12:45' = @('その他', '休憩')
    }
    $script:grid.Rows.Clear()
    for ($t = (Get-Date).Date.AddHours(8); $t -lt (Get-Date).Date.AddHours(13); $t = $t.AddMinutes($SLOT_MINUTES)) {
        $k = $t.ToString($TIME_FMT)
        if ($rows.Contains($k)) {
            [void]$script:grid.Rows.Add(@($k, $rows[$k][0], $rows[$k][1]))
        } else {
            [void]$script:grid.Rows.Add(@($k, '', ''))
        }
    }
    $script:gridForm.Text = '一覧編集（今日の記録）'
    $script:gBtnSave.Text = '保存して閉じる'
    Save-FormShot $script:gridForm (Join-Path $outDir '一覧編集.png') $backdrop
    $saved.Add('一覧編集.png')

    # 選択肢の編集（ダミーのモデルを組み立てて実画面に流し込む。設定ファイルは読まない）
    $script:maintProj = @{
        Header   = New-Object Collections.Generic.List[string]
        Items    = New-Object Collections.Generic.List[string]
        Comments = New-Object Collections.Generic.List[string]
    }
    foreach ($p in $SHOT_PROJECTS) { $script:maintProj.Items.Add($p) }
    $script:maintTask = @{
        Header   = New-Object Collections.Generic.List[string]
        Keys     = New-Object Collections.Generic.List[string]
        Items    = @{}
        Comments = @{}
    }
    Add-MaintSection ''
    foreach ($p in $SHOT_PROJECTS) { Add-MaintSection $p }
    foreach ($t in $SHOT_TASKS) { $script:maintTask.Items[$SHOT_PROJECT].Add($t) }
    $script:maintDirty = $false
    Update-MaintProjectList $SHOT_PROJECT
    Update-MaintTaskList
    Save-FormShot $script:maintForm (Join-Path $outDir '選択肢の編集.png') $backdrop
    $saved.Add('選択肢の編集.png')

    $backdrop.Close()
    $backdrop.Dispose()
    [ordered]@{ outDir = $outDir; files = $saved.ToArray() } | ConvertTo-Json
}

if ($Capture) {
    Invoke-CaptureShots
    Close-AllForms
    exit 0
}

# ============================================================
# 16. メイン
# ============================================================

# ---- 二重起動防止 ----
$mutex = New-Object Threading.Mutex($false, 'Local\NippoBatchMutex')
if (-not $mutex.WaitOne(0)) {
    Show-Info '日報入力は既に起動しています。タスクトレイのアイコンを右クリックして「日報メニューを表示」から操作できます。'
    exit 0
}

$notify = $null
try {
    # ショートカットが作れなくても記録はできるので、失敗しても続行する
    try { Initialize-DesktopShortcut } catch {}

    # ---- 業務開始ゲート: PC起動直後は待機し、押下してから記録を始める ----
    $gate = New-UiForm '日報入力' 360 150

    $gLbl = New-UiLabel 15 12 '「業務開始」を押すと、いまの15分枠から日報の記録を始めます。業務終了はタスクトレイのアイコン右クリックからいつでもできます。' 330 60

    $gStart = New-UiButton 15 85 200 42 '業務開始'
    $gStart.Font         = New-Object Drawing.Font('Meiryo UI', 12, [Drawing.FontStyle]::Bold)
    $gStart.DialogResult = [Windows.Forms.DialogResult]::OK

    $gQuit = New-UiButton 230 85 115 42 '今日は使わない'
    $gQuit.DialogResult = [Windows.Forms.DialogResult]::Cancel

    $gate.Controls.AddRange(@($gLbl, $gStart, $gQuit))
    $gate.AcceptButton = $gStart
    $gr = $gate.ShowDialog()
    $gate.Dispose()
    if ($gr -ne [Windows.Forms.DialogResult]::OK) { exit 0 }

    # 業務開始の時点で記録開始位置を確定（当日の続き or いまの枠の頭）
    $pendingStart = Get-InitialPendingStart

    # ---- タスクトレイ常駐アイコン: 15分区切りのウィンドウを待たずに各機能を呼べる ----
    $script:trayExit = $false
    $script:trayEdit = $false
    $notify = New-Object Windows.Forms.NotifyIcon
    try { $notify.Icon = [Drawing.Icon]::ExtractAssociatedIcon([Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }
    catch { $notify.Icon = [Drawing.SystemIcons]::Application }
    $notify.Text = '日報入力（右クリックで業務終了）'
    $trayMenu = New-Object Windows.Forms.ContextMenuStrip
    $miHub = $trayMenu.Items.Add('日報メニューを表示')
    $miHub.Add_Click({ $script:hubForm.Show(); $script:hubForm.Activate() })
    $miRsv = $trayMenu.Items.Add((Get-FeatureLabel 'Reserve'))
    $miRsv.Add_Click({ Invoke-Feature 'Reserve' })
    $miBreak = $trayMenu.Items.Add((Get-FeatureLabel 'Break'))
    $miBreak.Add_Click({ Invoke-Feature 'Break' })
    $miEdit = $trayMenu.Items.Add((Get-FeatureLabel 'Edit'))
    $miEdit.Add_Click({ $script:trayEdit = $true })   # 入力ウィンドウと重ならないようメインループ側で開く
    $miMaint = $trayMenu.Items.Add((Get-FeatureLabel 'Maint'))
    $miMaint.Add_Click({ Invoke-Feature 'Maint' })
    [void]$trayMenu.Items.Add('-')
    $miExit = $trayMenu.Items.Add((Get-FeatureLabel 'Close'))
    $miExit.Add_Click({ Invoke-Feature 'Close' })
    $notify.ContextMenuStrip = $trayMenu
    $notify.Visible = $true

    # 日報メニュー（常時表示）を開く。×で閉じても隠れるだけで、トレイから再表示できる
    Update-HubDates
    $hubForm.Show()

    # ---- メインループ ----
    # いまの枠が未記録なら即ウィンドウ表示（初回はここで最初の入力が出る）。
    # 短い間隔で回してDoEventsを呼び、トレイアイコンの操作に応答できるようにする
    $deferUntil = [datetime]::MinValue
    while (-not $script:trayExit) {
        [Windows.Forms.Application]::DoEvents()
        if ($script:trayEdit) { $script:trayEdit = $false; Invoke-Feature 'Edit' }
        if ($script:trayExit) { break }
        $t = Get-Date
        if ((Get-FloorSlot $t) -ge $pendingStart -and $t -ge $deferUntil) {
            Update-PendingStart   # MTG予約済みの時間帯はウィンドウを出さず、次の未記録枠まで進める
            if ((Get-FloorSlot $t) -ge $pendingStart) {
                $r = Show-EntryDialog
                if ($r -eq [Windows.Forms.DialogResult]::Abort) { break }
                if ($r -eq [Windows.Forms.DialogResult]::Retry) { Invoke-Feature 'Break' }   # 完了まで待って記録
                if ($r -eq [Windows.Forms.DialogResult]::Cancel) {
                    # ×で閉じた → 保留のまま、次の15分区切りで再表示
                    $deferUntil = (Get-FloorSlot (Get-Date)).AddMinutes($SLOT_MINUTES)
                }
            }
        }
        Start-Sleep -Milliseconds 50
    }
}
catch {
    Show-Info "日報入力でエラーが発生しました:`n$($_.Exception.Message)"
}
finally {
    if ($notify) { $notify.Visible = $false; $notify.Dispose() }
    Close-AllForms
    [void]$mutex.ReleaseMutex()
}
