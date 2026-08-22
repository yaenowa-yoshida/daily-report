# ユーザーマニュアル生成.ps1 — ユーザーマニュアル.xlsx を作り直す
#
# マニュアルの本文はこのスクリプトが持っている（xlsxは出力物）。
# 内容を直すときは下の「マニュアル本文」を編集して、このスクリプトを実行する:
#   powershell -NoProfile -ExecutionPolicy Bypass -File 日報記録\ユーザーマニュアル生成.ps1
# 実行にはExcelが必要（COM経由で書式付きのブックを作る）。既存のxlsxは上書きされる。

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing   # 画像の元サイズを読むのに使う
$outPath = Join-Path $PSScriptRoot 'ユーザーマニュアル.xlsx'

# ---- 配色（Excelの色は R + G*256 + B*65536） ----
function Get-Color([int]$r, [int]$g, [int]$b) { return ($r + $g * 256 + $b * 65536) }
$C_TITLE   = Get-Color 31 56 100     # 濃紺（見出し文字）
$C_HEAD_BG = Get-Color 68 114 196    # 青（表のヘッダ背景）
$C_SEC_BG  = Get-Color 217 225 242   # 薄青（セクション帯）
$C_NOTE_BG = Get-Color 255 242 204   # 薄黄（注意書き）
$C_STEP_BG = Get-Color 226 240 217   # 薄緑（手順）
$C_WHITE   = Get-Color 255 255 255
$C_LINE    = Get-Color 191 191 191   # 罫線

$COL_FIRST = 2   # B列から書く（A列は余白）
$COL_LAST  = 4   # D列まで使う
$CHARS_PER_LINE = 46   # 行の高さを見積もるための1行あたり文字数（全角換算）

# ---- 書き込みヘルパー（$ctx.Row が現在行） ----

# 文字数から必要な行数を見積もって高さを決める（結合セルは自動調整が効かないため）
function Set-RowHeight($ws, [int]$row, [string]$text, [int]$charsPerLine, [int]$minLines = 1) {
    $lines = 0
    foreach ($seg in ($text -split "`n")) {
        $n = [math]::Ceiling(([double]$seg.Length) / $charsPerLine)
        if ($n -lt 1) { $n = 1 }
        $lines += $n
    }
    if ($lines -lt $minLines) { $lines = $minLines }
    $ws.Rows.Item($row).RowHeight = 16 * $lines + 4
}

function Add-Title($ws, $ctx, [string]$text) {
    $cell = $ws.Cells.Item($ctx.Row, $COL_FIRST)
    $cell.Value2 = $text
    $cell.Font.Size = 16
    $cell.Font.Bold = $true
    $cell.Font.Color = $C_TITLE
    $ws.Rows.Item($ctx.Row).RowHeight = 30
    $ctx.Row++
}

function Add-Section($ws, $ctx, [string]$text) {
    $ctx.Row++   # 前に1行空ける
    $range = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
    $range.Merge()
    $range.Value2 = $text
    $range.Font.Bold = $true
    $range.Font.Size = 12
    $range.Font.Color = $C_TITLE
    $range.Interior.Color = $C_SEC_BG
    $range.VerticalAlignment = -4108   # 中央
    $ws.Rows.Item($ctx.Row).RowHeight = 24
    $ctx.Row++
}

function Add-Para($ws, $ctx, [string]$text) {
    $range = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
    $range.Merge()
    $range.Value2 = $text
    $range.WrapText = $true
    $range.VerticalAlignment = -4160   # 上寄せ
    Set-RowHeight $ws $ctx.Row $text ($CHARS_PER_LINE * 2)
    $ctx.Row++
}

# 背景色付きの1行（注意書き・手順）
function Add-Box($ws, $ctx, [string]$text, [int]$bg) {
    $range = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
    $range.Merge()
    $range.Value2 = $text
    $range.WrapText = $true
    $range.Interior.Color = $bg
    $range.VerticalAlignment = -4160
    $range.IndentLevel = 1
    Set-RowHeight $ws $ctx.Row $text ($CHARS_PER_LINE * 2)
    $ctx.Row++
}

function Add-Note($ws, $ctx, [string]$text) { Add-Box $ws $ctx $text $C_NOTE_BG }

# 番号付きの手順
function Add-Steps($ws, $ctx, [string[]]$items) {
    $i = 1
    foreach ($item in $items) {
        Add-Box $ws $ctx ("$i. $item") $C_STEP_BG
        $i++
    }
}

# 表。$headers は2〜3列。2列のときは3列目を結合して広く使う
function Add-Table($ws, $ctx, [string[]]$headers, [object[]]$rows) {
    $twoCol = ($headers.Count -eq 2)
    $first = $ctx.Row
    # ヘッダ
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item($ctx.Row, $COL_FIRST + $c).Value2 = $headers[$c]
    }
    if ($twoCol) {
        $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST + 1), $ws.Cells.Item($ctx.Row, $COL_LAST)).Merge()
    }
    $hdr = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
    $hdr.Font.Bold = $true
    $hdr.Font.Color = $C_WHITE
    $hdr.Interior.Color = $C_HEAD_BG
    $hdr.VerticalAlignment = -4108
    $ws.Rows.Item($ctx.Row).RowHeight = 22
    $ctx.Row++
    # 本文
    foreach ($row in $rows) {
        $longest = ''
        for ($c = 0; $c -lt $row.Count; $c++) {
            $ws.Cells.Item($ctx.Row, $COL_FIRST + $c).Value2 = $row[$c]
            if (([string]$row[$c]).Length -gt $longest.Length) { $longest = [string]$row[$c] }
        }
        if ($twoCol) {
            $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST + 1), $ws.Cells.Item($ctx.Row, $COL_LAST)).Merge()
        }
        $line = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
        $line.WrapText = $true
        $line.VerticalAlignment = -4160
        $ws.Cells.Item($ctx.Row, $COL_FIRST).Font.Bold = $true
        Set-RowHeight $ws $ctx.Row $longest $CHARS_PER_LINE
        $ctx.Row++
    }
    # 罫線
    $tbl = $ws.Range($ws.Cells.Item($first, $COL_FIRST), $ws.Cells.Item($ctx.Row - 1, $COL_LAST))
    foreach ($edge in 7, 8, 9, 10, 11, 12) {   # 左・上・下・右・内側縦・内側横
        $b = $tbl.Borders.Item($edge)
        $b.LineStyle = 1
        $b.Color = $C_LINE
    }
}

# 画面のスクリーンショットを貼る（画像\ 配下。Nippo-Batch.ps1 -Capture で撮り直せる）。
# 画像が無ければ何も貼らずに進む（Excelだけあれば生成できる状態を保つ）
function Add-Shot($ws, $ctx, [string]$fileName, [string]$caption = '', [int]$maxWidth = 460) {
    $path = Join-Path (Join-Path $PSScriptRoot '画像') $fileName
    if (-not (Test-Path -LiteralPath $path)) { return }
    $img = [Drawing.Image]::FromFile($path)
    $w = [double]$img.Width; $h = [double]$img.Height
    $img.Dispose()
    if ($w -gt $maxWidth) { $h = $h * ($maxWidth / $w); $w = $maxWidth }
    # Excelの行の高さは409ポイントが上限。縦長の画像はそこに収まるよう更に縮める
    $maxHeight = 380
    if ($h -gt $maxHeight) { $w = $w * ($maxHeight / $h); $h = $maxHeight }
    $cell = $ws.Cells.Item($ctx.Row, $COL_FIRST)
    # AddPicture(ファイル, リンクしない, ブックに保存する, 左, 上, 幅, 高さ)
    [void]$ws.Shapes.AddPicture($path, $false, $true, ($cell.Left + 2), ($cell.Top + 2), $w, $h)
    # RowHeight に小数を渡すとCOMがキャスト例外を出すので整数に丸める
    $ws.Rows.Item($ctx.Row).RowHeight = [int][math]::Ceiling($h + 8)
    $ctx.Row++
    if ($caption) {
        $range = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
        $range.Merge()
        $range.Value2 = $caption
        $range.Font.Size = 9
        $range.Font.Color = (Get-Color 110 110 110)
        $range.IndentLevel = 1
        $ws.Rows.Item($ctx.Row).RowHeight = 18
        $ctx.Row++
    }
}

# CSVサンプルなどの等幅表示
function Add-Code($ws, $ctx, [string[]]$lines) {
    foreach ($l in $lines) {
        $range = $ws.Range($ws.Cells.Item($ctx.Row, $COL_FIRST), $ws.Cells.Item($ctx.Row, $COL_LAST))
        $range.Merge()
        $range.Value2 = $l
        $range.Font.Name = 'ＭＳ ゴシック'
        $range.Interior.Color = (Get-Color 242 242 242)
        $range.IndentLevel = 1
        $ws.Rows.Item($ctx.Row).RowHeight = 18
        $ctx.Row++
    }
}

function Initialize-Sheet($excel, $ws, [string]$name) {
    $ws.Name = $name
    $ws.Cells.Font.Name = 'Meiryo UI'
    $ws.Cells.Font.Size = 10
    $ws.Columns.Item(1).ColumnWidth = 2      # A列は余白
    $ws.Columns.Item(2).ColumnWidth = 26
    $ws.Columns.Item(3).ColumnWidth = 46
    $ws.Columns.Item(4).ColumnWidth = 34
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    return @{ Row = 2 }
}

# ============================================================
# マニュアル本文
# ============================================================

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$book = $excel.Workbooks.Add()
# 既定シートを1枚だけ残す
while ($book.Sheets.Count -gt 1) { $book.Sheets.Item($book.Sheets.Count).Delete() }

$sheetNames = @('はじめに', '一日の流れ', '入力画面の使い方', 'こんなときは', '記録の修正と設定', 'CSVと困ったとき')
for ($i = 2; $i -le $sheetNames.Count; $i++) { [void]$book.Sheets.Add([Type]::Missing, $book.Sheets.Item($book.Sheets.Count)) }

# ---- 1. はじめに ----
$ws = $book.Sheets.Item(1)
$ctx = Initialize-Sheet $excel $ws $sheetNames[0]
Add-Title $ws $ctx '日報入力ツール ユーザーマニュアル'
Add-Para $ws $ctx '15分ごとに「これから何をやるか」を聞いてくるので、答えていくだけで日報ができあがるツールです。あとから思い出して書く必要がなくなります。'
Add-Section $ws $ctx 'まず、これだけ覚えれば使えます'
Add-Steps $ws $ctx @(
    '朝いちばんに、デスクトップの「日報入力」をダブルクリックして「業務開始」を押す'
    '15分ごとに出る画面で、案件と作業内容を選んで「記録 (OK)」を押す'
    '帰るときに「業務終了」→ 一日の記録を確認して「締め」を押す'
)
Add-Para $ws $ctx 'これで「日報ログ」フォルダに、その日のCSVができています。月次報告書へは、そのCSVを見ながら転記してください。'
Add-Section $ws $ctx 'このマニュアルの構成'
Add-Table $ws $ctx @('シート', '内容') @(
    @('一日の流れ', '朝の業務開始から、夕方の業務終了までの流れ'),
    @('入力画面の使い方', '15分ごとに出る画面のボタンと、入力のコツ'),
    @('こんなときは', '休憩・会議・画面を見失ったときなど、状況ごとの操作'),
    @('記録の修正と設定', 'あとから記録を直す方法と、選択肢の増やし方'),
    @('CSVと困ったとき', 'できあがるCSVの見方と、トラブル対処')
)
Add-Section $ws $ctx '大事な注意'
Add-Note $ws $ctx '業務開始を押したあと、ツールを止める方法は「業務終了 → 締め」だけです。画面の × を押しても止まりません（記録が止まってしまわないための作りです）。画面を見失ったら、タスクトレイのアイコンを右クリック →「日報メニューを表示」で戻せます。'

# ---- 2. 一日の流れ ----
$ws = $book.Sheets.Item(2)
$ctx = Initialize-Sheet $excel $ws $sheetNames[1]
Add-Title $ws $ctx '一日の流れ'
Add-Section $ws $ctx '朝：業務開始'
Add-Para $ws $ctx 'デスクトップの「日報入力」アイコンをダブルクリックすると起動します。「業務開始」を押すまで記録は始まりませんので、朝の準備が済んでから押してください。押した時点の15分枠（9:07なら9:00）から記録が始まります。'
Add-Note $ws $ctx 'その日を使わないときは「今日は使わない」を押せば、何も起きずに閉じます。デスクトップにアイコンが見当たらないときは、日報記録フォルダの「日報入力.bat」をダブルクリックしてください。起動時にデスクトップのショートカットが作り直されます。'
Add-Section $ws $ctx '日中：15分ごとに答える'
Add-Para $ws $ctx '区切りのタイミング（0分・15分・30分・45分）で入力画面が出ます。音が鳴って手前に出てきます。「これから何をやるか」を答えるのがポイントです。終わった作業を報告するのではありません。'
Add-Section $ws $ctx '夕方：業務終了'
Add-Para $ws $ctx '「業務終了」を押すと、その日の記録が表で出てきます。抜けや間違いをその場で直して、「締め（保存して業務終了）」を押すと保存されてツールが終了します。やっぱり続けるときは「キャンセル」を押せば、そのまま仕事に戻れます。'
Add-Section $ws $ctx '2つの画面を使い分けます'
Add-Table $ws $ctx @('画面', '出るタイミング', 'できること') @(
    @('入力画面', '15分ごとに自動で出る', '案件と作業内容を記録する。休憩・スキップ・業務終了もここから'),
    @('日報メニュー', 'ずっと出ている小窓', '日付を選んで記録を編集、MTG予約、休憩、選択肢の編集、業務終了')
)
Add-Shot $ws $ctx '日報メニュー.png' 'こちらが日報メニュー。起動中はずっと出ています。' 340
Add-Note $ws $ctx '日報メニューは × で閉じても隠れるだけです。タスクトレイ（画面右下の時計のそば。隠れているときは「∧」をクリック）のアイコンを右クリック →「日報メニューを表示」で戻せます。'

# ---- 3. 入力画面の使い方 ----
$ws = $book.Sheets.Item(3)
$ctx = Initialize-Sheet $excel $ws $sheetNames[2]
Add-Title $ws $ctx '入力画面の使い方'
Add-Shot $ws $ctx '入力ウィンドウ.png' '15分区切りごとに、この画面が音とともに出てきます。'
Add-Section $ws $ctx '4つのボタン'
Add-Table $ws $ctx @('ボタン', 'どんなとき', '何が起きる') @(
    @('記録 (OK)', 'いつもの作業をするとき', '選んだ時間ぶんを記録して閉じる'),
    @('スキップ', '記録に残したくない時間', '記録せず飛ばす（表には空欄が残る）'),
    @('休憩', '休憩・昼食に入るとき', '休憩モードへ。戻って「休憩完了」を押すまでが休憩になる'),
    @('業務終了', '帰るとき', '一日の記録を表で確認して締める')
)
Add-Section $ws $ctx '入力のコツ'
Add-Para $ws $ctx '案件 → 作業内容の順に選びます。案件を選ぶと、その案件の作業内容だけが候補に出るので探しやすくなります。'
Add-Table $ws $ctx @('こんなとき', 'どうする') @(
    @('候補にない作業をした', 'そのまま打ち込んでOK。次回から候補に出るようになります（覚えさせる操作は不要）'),
    @('さっきの続きをやる', '「前回の入力を保持する」にチェックが入っていれば、前回と同じ内容が入った状態で出ます。そのまま「記録 (OK)」を押すだけ'),
    @('1時間の会議に出る', '「記録する時間」で60分を選ぶと、次の画面は1時間後まで出ません。この選択は毎回15分に戻ります'),
    @('いま答えられない', '× で閉じてください。記録は保留になり、次の15分区切りでもう一度聞いてきます')
)
Add-Section $ws $ctx '画面の上に出る「対象」の意味'
Add-Code $ws $ctx @('対象: 10:15 ～ 10:45（2枠 × 15分）')
Add-Para $ws $ctx 'これは「いまから記録される時間帯」です。席を外していて画面を何回か見逃していた場合、たまっていたぶんもまとめてこの内容で記録されます。'

# ---- 4. こんなときは ----
$ws = $book.Sheets.Item(4)
$ctx = Initialize-Sheet $excel $ws $sheetNames[3]
Add-Title $ws $ctx 'こんなときは'
Add-Table $ws $ctx @('状況', 'どうする') @(
    @('休憩・お昼に行く', '入力画面か日報メニューの「休憩」を押す（下に詳しい手順があります）'),
    @('これから会議がある', '日報メニューの「MTG予約」で先に登録しておく（下に詳しい手順があります）'),
    @('候補が増えすぎた／名前を直したい', '日報メニューの「選択肢の編集」で整理する（「記録の修正と設定」シート）'),
    @('画面を見失った', 'タスクトレイのアイコンを右クリック →「日報メニューを表示」'),
    @('「既に起動しています」と出る', 'まだ裏で動いています（正常）。上と同じ方法で画面を戻してください'),
    @('いま答えられない', '× で閉じる。次の15分区切りでもう一度聞いてきます')
)
Add-Section $ws $ctx '休憩・お昼に行く'
Add-Para $ws $ctx '「休憩」を押すだけです。何分休むかは決めなくて構いません。'
Add-Steps $ws $ctx @(
    '「休憩」を押すと「休憩中」の小さい窓が出て待機します（何時開始・何分経過かが表示されます）'
    '戻ってきたら「休憩完了（業務に戻る）」を押します'
    'その時点までが休憩として記録され、続けて次の入力画面が出ます'
)
Add-Shot $ws $ctx '休憩中.png' '休憩中はこの窓が出たまま待機します。' 360
Add-Note $ws $ctx '例：12:03に休憩 → 12:52に戻って完了。12:00・12:15・12:30が休憩になり、12:45からの作業を聞かれます。記録には案件「その他」・作業内容「休憩」と入ります。15分たたずに戻った場合は何も記録されません（押し間違えても大丈夫です）。'
Add-Para $ws $ctx '「スキップ」でも時間は飛ばせますが、表に何も残らないので、あとから見て「入力し忘れ」なのか「休憩」なのか分からなくなります。休憩のときは「休憩」ボタンを使ってください。'
Add-Section $ws $ctx 'これから会議がある（時間が決まっている予定）'
Add-Para $ws $ctx '先に予約しておくと、その時間帯は入力画面が出ません。会議中に画面が割り込んでこないので便利です。'
Add-Steps $ws $ctx @(
    '日報メニューの「MTG予約（当日の事前入力）」を押す'
    '開始時刻・時間（30分か60分）・案件・作業内容を入れて「予約する」を押す'
)
Add-Shot $ws $ctx 'MTG予約.png' '開始時刻と時間（30分か60分）を選ぶだけ。あとは案件と作業内容を入れて「予約する」。'
Add-Note $ws $ctx '会議が30分・60分以外のときは、いちばん近い長さで予約しておいて、終わってから「編集」で時間帯を直すのが簡単です。'
Add-Para $ws $ctx '会議が終わった次の区切りから、いつも通り入力画面が出ます。予定がなくなったら「編集」でその時間帯を「選択した行をクリア」して保存すれば、予約が取り消されてその枠から入力画面が復活します。'

# ---- 5. 記録の修正と設定 ----
$ws = $book.Sheets.Item(5)
$ctx = Initialize-Sheet $excel $ws $sheetNames[4]
Add-Title $ws $ctx '記録の修正と設定'
Add-Section $ws $ctx '一日の記録を直す'
Add-Para $ws $ctx '日報メニューで日付を選んで「編集」を押すと、その日の記録が「時刻 × 案件 × 作業内容」の表で開きます。過去の日も直せます。セルを直接書き換えるほか、下のボタンでまとめて操作できます。'
Add-Shot $ws $ctx '一覧編集.png' '業務終了のときに出る確認画面も、これと同じ表です。' 480
Add-Table $ws $ctx @('ボタン', '動き') @(
    @('選択した行に入力', '下の入力欄の内容を、選んでいる行すべてに入れる'),
    @('選択した行をクリア', '選んでいる行を空にする（その時間の記録を消す）'),
    @('この行を入力欄へコピー', 'クリックした行の内容を、下の入力欄に写す')
)
Add-Section $ws $ctx 'よくある使い方：同じ作業を別の時間にも入れる'
Add-Steps $ws $ctx @(
    'コピーしたい行をクリックする'
    '「この行を入力欄へコピー」を押す'
    '入れたい行を選ぶ（ドラッグや Ctrl+クリックで複数選択できます）'
    '「選択した行に入力」を押す'
    '最後に「保存して閉じる」を押す'
)
Add-Note $ws $ctx '案件と作業内容は両方入れてください。片方だけだと保存時に注意されます。空欄のまま保存すれば、その時間は「記録なし」になります。'
Add-Section $ws $ctx '選択肢（案件・作業内容）を整理する'
Add-Para $ws $ctx '日報メニューの「選択肢の編集（案件・作業内容）」を押すと、案件と作業内容の一覧を画面から直せます。使わなくなった作業内容を消す、名前を整える、よく使うものを上に持ってくる、といった手入れはここでできます。'
Add-Shot $ws $ctx '選択肢の編集.png' '左が案件、右がその案件の作業内容。左で案件を選ぶと右が切り替わります。' 480
Add-Table $ws $ctx @('ボタン', '動き') @(
    @('追加', '新しい案件（左）または作業内容（右）を増やす'),
    @('名前変更', '選んでいるものの名前を付け直す'),
    @('削除', '選んでいるものを候補から消す（案件を消すと、その案件の作業内容も一緒に消えます）'),
    @('↑ 上へ ／ ↓ 下へ', '並び順を入れ替える。ドロップダウンにはこの順で出ます'),
    @('別の案件へ移動', '作業内容を、別の案件（または共通）に付け替える'),
    @('保存して閉じる', 'ここまでの編集を保存する。押すまではファイルは変わりません')
)
Add-Note $ws $ctx '名前を変えたり削除したりしても、過去の記録（CSV）はそのままです。変わるのは「これから選べる候補」だけなので、過去の日報が書き換わる心配はありません。過去の記録の表記を直したいときは「編集」で該当の行を直してください。'
Add-Para $ws $ctx '左の一覧のいちばん上にある「（共通：すべての案件で表示）」は、どの案件を選んでも候補に出る作業内容です。「チームMTG」のように案件をまたいで使うものはここに置いておくと便利です。'
Add-Note $ws $ctx 'ふだんは編集不要です。入力画面で打ち込んだ作業内容は、選んでいた案件のところに自動で追加されていきます。この画面は「増えすぎた候補を整理したくなったとき」に使ってください。'
Add-Para $ws $ctx 'なお、中身は「日報設定」フォルダのテキストファイル（案件.txt ／ 作業名.txt）です。メモ帳で直接編集しても構いません（行頭が # の行はメモ書き扱いで無視されます）。どちらの方法でも、保存すれば次に画面が出たときから反映されます（ツールの再起動は不要）。'

# ---- 6. CSVと困ったとき ----
$ws = $book.Sheets.Item(6)
$ctx = Initialize-Sheet $excel $ws $sheetNames[5]
Add-Title $ws $ctx 'CSVと困ったとき'
Add-Section $ws $ctx 'できあがるCSV'
Add-Para $ws $ctx '「日報ログ」フォルダに、1日1ファイルで保存されています（例：日報ログ_2026-08-07.csv）。'
Add-Code $ws $ctx @(
    '日付,時刻,案件,作業内容'
    '"2026/08/07","9:00","案件A","API設計レビュー"'
    '"2026/08/07","9:15","案件A","API設計レビュー"'
    '"2026/08/07","9:30","案件A","PRJ-12 商品ページ改修"'
    '"2026/08/07","12:00","その他","休憩"'
)
Add-Para $ws $ctx '1行が15分です。同じ作業が続けば同じ行が並びます。Excelで開いてピボットテーブルにすれば、案件ごとの合計時間もすぐ出せます（15分 × 行数で計算できます）。'
Add-Note $ws $ctx '記録中のCSVをExcelで開いたままにしないでください。開いている間は書き込めず、「CSVに書き込めませんでした」という案内が出ます。閉じてからもう一度ボタンを押せば記録されます。なお、Excelで開いて保存すると日付の書き方が変わることがありますが、ツール側で吸収するので壊れません。'
Add-Section $ws $ctx 'フォルダの中身'
Add-Para $ws $ctx 'このツールは「日報記録」フォルダひとつで完結しています。ふだん開くのは、下の3つだけです。'
Add-Table $ws $ctx @('名前', '中身') @(
    @('日報入力.bat', '起動用。デスクトップの「日報入力」ショートカットは、これを指しています'),
    @('日報ログ（フォルダ）', '記録の出力先。1日1ファイルのCSVが貯まります'),
    @('日報設定（フォルダ）', '案件.txt と 作業名.txt。選択肢を変えたいときはここ'),
    @('ユーザーマニュアル.xlsx', 'このファイル'),
    @('日報システム紹介プレゼン.html', '社内紹介用のスライド'),
    @('Nippo-Batch.ps1', 'プログラム本体。動きを変えたいとき以外は触りません'),
    @('ユーザーマニュアル生成.ps1 ／ 画像（フォルダ）', 'このマニュアルを作り直すための道具。触らなくて大丈夫です')
)
Add-Note $ws $ctx 'フォルダごと移動しても動きます（記録も設定も、同じフォルダの中を見にいくため）。ただし移動したあとは、デスクトップのショートカットを作り直すために「日報入力.bat」から一度起動してください。'
Add-Section $ws $ctx '困ったときは'
Add-Table $ws $ctx @('こまりごと', 'どうするか') @(
    @('画面が見当たらない', 'タスクトレイのアイコンを右クリック →「日報メニューを表示」'),
    @('「既に起動しています」と出る', 'すでに動いています。上と同じ方法で画面を戻してください'),
    @('「CSVに書き込めませんでした」と出る', 'その日のCSVをExcelで開いていたら閉じて、もう一度ボタンを押す'),
    @('入力し忘れた時間がある', '日報メニューの「編集」で、その日を選んで空欄を埋める'),
    @('昨日の記録を直したい', '同じく「編集」で日付を選ぶ（過去の日も直せます）'),
    @('候補に出したくない作業がある', '日報メニューの「選択肢の編集」でその作業内容を選んで「削除」'),
    @('予約した会議がなくなった', '「編集」でその時間帯を選び「選択した行をクリア」→ 保存'),
    @('PCを再起動した／途中で落ちた', 'デスクトップの「日報入力」からもう一度起動すれば、記録済みの続きから再開します（記録は毎回すぐ保存されているので失われません）'),
    @('デスクトップのアイコンが消えた', '日報記録フォルダの「日報入力.bat」をダブルクリックして起動すると、ショートカットが作り直されます')
)
Add-Para $ws $ctx '上記で解決しないときは、日報ログフォルダのCSVを直接開いて手で直しても大丈夫です（列の並びだけ変えないでください）。'

# ---- 保存 ----
$book.Sheets.Item(1).Activate()
$excel.ActiveWindow.ScrollRow = 1
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
$book.SaveAs($outPath, 51)   # 51 = xlsx
$book.Close($false)
$excel.Quit()
foreach ($o in $ws, $book, $excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Output "作成しました: $outPath"
