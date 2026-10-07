# ============================================================
# step6-return : 複数の値を返す、関数の中の変数（スコープ）
# ------------------------------------------------------------
# 学ぶこと
#   1. 複数の値を出力して、左側に変数を並べて受け取る
#   2. 余計な出力が混ざると、受け取る値がずれる
#   3. オブジェクトにまとめて返す（名前で取り出せる。ふだんはこれを使う）
#   4. [ref] で引数に書き戻してもらう
#   5. 関数の中の変数は、外の変数と「半分だけ」別
#   6. & で呼んだスクリプトの変数（& と . と pwsh -File の違い）
#   7. Excel で使う（シートの大きさをオブジェクトで返す、関数の中から $sheet に書く）
#
# 使うファイル
#   step6-scope-child.ps1 : 6 で呼ぶ子のスクリプト
#   sample-org.xlsx       : 7 で読み取り専用で開く（保存しないので、ファイルは変わらない）
#
# 実行方法
#   .\step6-return.ps1
#
# 覚えておくこと
#   - 関数やスクリプトが出力したものは、全部まとめて戻り値になる（README の「戻り値のしくみ」）
#   - $a, $b = 関数 は「順番」で受け取る。余計な出力が 1 つ混ざると全部ずれる
#   - [pscustomobject]@{ 名前 = 値; ... } で返すと、$r.名前 で取り出せる
#   - 関数の中では、外の変数を読める。代入すると関数の中だけの変数になり、外は変わらない
#   - 配列の中身や Excel の $sheet のように、変数が指している先を書き換えると外にも反映される
#   - 関数で使う値は引数で渡し、外の変数を変えたいときは戻り値で返して呼び出し側で代入する
#   - & で呼んだスクリプトも関数と同じ。. で呼ぶと変数が親に残る。pwsh -File では親の変数は見えない
# ============================================================

# ============================================================
# 1. 複数の値を出力して、左側に変数を並べて受け取る
# ============================================================
Write-Host "--- 1. 変数を並べて受け取る"

# 最小値と最大値の 2 つを返す
function Get-MinMax($nums) {
    $min = ($nums | Measure-Object -Minimum).Minimum
    $max = ($nums | Measure-Object -Maximum).Maximum
    return $min, $max                              # 2 つ出力する
}

$lo, $hi = Get-MinMax @(3, 8, 1, 5)                # 1 つ目が $lo、2 つ目が $hi に入る
Write-Host "lo = $lo  hi = $hi"

$r = Get-MinMax @(3, 8, 1, 5)                      # 1 つの変数で受けると配列になる
Write-Host "r[0] = $($r[0])  r[1] = $($r[1])  個数 = $($r.Count)"

# 変数と値の数が合わないとき
$a, $b, $c = 1, 2                                  # 足りない分は $null
Write-Host "a = $a  b = $b  c は空 = $($null -eq $c)"
$a, $b = 1, 2, 3                                   # 余った分は最後の変数に配列で入る
Write-Host "a = $a  b = $($b -join ', ')（$($b.Count) 個）"

# ============================================================
# 2. 余計な出力が混ざると、受け取る値がずれる
# ============================================================
Write-Host "--- 2. 余計な出力が混ざる"

# ArrayList の Add() は「追加した位置」を返す。変数に入れずに呼ぶと、その値も出力になる
function Get-PairBad {
    $list = New-Object System.Collections.ArrayList
    $list.Add("x")                                 # 0 が出力に混ざる
    return "a", "b"
}
$p, $q = Get-PairBad
Write-Host "直す前 : p = $p  q = $q"               # p = 0、q に a と b が入ってしまう

# いらない出力は [void] で捨てる（| Out-Null や $null = でもよい）
function Get-PairGood {
    $list = New-Object System.Collections.ArrayList
    [void]$list.Add("x")
    return "a", "b"
}
$p, $q = Get-PairGood
Write-Host "直した後 : p = $p  q = $q"

# ============================================================
# 3. オブジェクトにまとめて返す（ふだんはこれを使う）
# ============================================================
Write-Host "--- 3. オブジェクトにまとめて返す"

# [pscustomobject]@{ 名前 = 値; ... } で、名前の付いた値をまとめて 1 つにする
function Get-Score {
    [pscustomobject]@{ Name = "太郎"; Total = 145; Average = 72.5 }
}

$s = Get-Score
Write-Host "Name = $($s.Name)  Total = $($s.Total)  Average = $($s.Average)"    # 名前で取り出す
Write-Host "まとめて表示 : $s"

# ============================================================
# 4. [ref] で引数に書き戻してもらう
# ============================================================
Write-Host "--- 4. [ref]"

# 戻り値で "ok" を返し、件数は引数に書き戻す
function Get-Info([ref]$count) {
    $count.Value = 3                               # .Value に代入すると、呼び出し側の変数が変わる
    return "ok"
}
$n = 0
$st = Get-Info ([ref]$n)                           # [ref] を付けた変数をカッコで囲んで渡す
Write-Host "st = $st  n = $n"

# ============================================================
# 5. 関数の中の変数（スコープ）
# ============================================================
Write-Host "--- 5. 関数の中の変数"
$x = "外の値"

function Test-Read   { Write-Host "読む           : $x" }                  # 外の値が読める
function Test-Write  { $x = "中で代入"; Write-Host "代入           : $x" } # 関数の中だけの $x
function Test-Add    { $x += "に追加";  Write-Host "+=             : $x" } # 空から始まる
function Test-Param($x) { Write-Host "同じ名前の引数 : $x" }               # 引数が使われる

Test-Read
Test-Write
Test-Add
Test-Param "引数の値"
Write-Host "ここまでの後の外の x = $x"                                     # 変わっていない

# $script: を付けると、スクリプト全体の変数に代入する
function Test-Script { $script:x = "script: で代入" }
Test-Script
Write-Host "script: の後の外の x = $x"

# 配列の中身を書き換えると、外にも反映される（変数ではなく、指している先を書き換えるため）
$arr = @("元", "元")
function Test-Array { $arr[0] = "書き換え" }
Test-Array
Write-Host "配列の中身 = $($arr -join ', ')"

# おすすめの形：値は引数で受け取り、結果は戻り値で返して、呼び出し側で代入する
function Add-Suffix($text, $suffix) { return $text + $suffix }
$x = Add-Suffix $x "（戻り値で更新）"
Write-Host "戻り値で更新した x = $x"

# ============================================================
# 6. & で呼んだスクリプトの変数（step6-scope-child.ps1 を 3 通りで呼ぶ）
# ============================================================
# & で呼ぶ     : 関数と同じ。親の変数を読めるが、代入しても親は変わらない。終わると子の変数と関数は消える
# . で呼ぶ     : 親の中でそのまま実行する。変数も関数も親に残り、同じ名前の変数は上書きされる
#                （step3.ps1 の . excel-lib.ps1 で関数を読み込めるのはこのため）
# pwsh -File   : 別のプロセスで動く。親の変数はまったく見えない（bat から呼んだときと同じ）
Write-Host "--- 6. & で呼んだスクリプトの変数"
$child = "$PSScriptRoot\step6-scope-child.ps1"

# 呼ぶ前の状態に戻し、呼んだ後の親の変数を表示する
function Reset-Vars {
    $script:x      = "親の値"
    $script:y      = "親の y"
    $script:arr    = @("元", "元")
    $script:newVar = $null
}
function Show-Vars {
    $func = if (Get-Command Child-Func -ErrorAction SilentlyContinue) { "ある" } else { "ない" }
    Write-Host "  親：x = [$x]  y = [$y]  newVar = [$newVar]  arr = [$($arr -join ', ')]  Child-Func = $func"
}

Write-Host "& で呼ぶ"
Reset-Vars
& $child
Show-Vars

Write-Host "pwsh -File で呼ぶ（別のプロセス）"
Reset-Vars
pwsh -NoProfile -File $child
Show-Vars

Write-Host ". で呼ぶ"
Reset-Vars
. $child
Show-Vars

# ============================================================
# 7. Excel で使う（3 と 5 を Excel のシートで試す）
# ============================================================
Write-Host "--- 7. Excel で使う"

# 最終行、最終列、使っている範囲の番地をまとめて返す
#   例：$size = Get-SheetSize $sheet
#       $size.LastRow、$size.LastCol、$size.Address
function Get-SheetSize($sheet) {
    $xlUp     = -4162
    $xlToLeft = -4159
    $lastRow = $sheet.Cells.Item($sheet.Rows.Count, 1).End($xlUp).Row
    $lastCol = $sheet.Cells.Item(1, $sheet.Columns.Count).End($xlToLeft).Column
    [pscustomobject]@{
        LastRow = $lastRow
        LastCol = $lastCol
        Address = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($lastRow, $lastCol)).Address(0, 0)
    }
}

# 指定した列の値を、最終行まで書き換える
#   引数の $sheet は呼び出し側と同じシートを指しているので、外のシートが書き換わる
function Set-Column($sheet, $col, $text) {
    $size = Get-SheetSize $sheet                   # 関数の中から別の関数を呼んでもよい
    for ($r = 1; $r -le $size.LastRow; $r++) {
        $sheet.Cells.Item($r, $col).Value2 = $text
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 読み取り専用で開く（第 3 引数 $true）。最後は保存せずに閉じる
    $book  = $excel.Workbooks.Open("E:\dev\excel\sample-org.xlsx", 0, $true)
    $sheet = $book.Worksheets.Item("Sheet1")

    # 3 と同じ：オブジェクトで返し、名前で取り出す
    $size = Get-SheetSize $sheet
    Write-Host "最終行 = $($size.LastRow)  最終列 = $($size.LastCol)  範囲 = $($size.Address)"
    Write-Host "まとめて表示 : $size"

    # 関数の中から $sheet に書くと、外のシートが書き換わる（5 の配列と同じ）
    Write-Host "書く前の A1 = $($sheet.Range('A1').Value2)"
    Set-Column $sheet 1 "済"
    Write-Host "書いた後の A1 = $($sheet.Range('A1').Value2)  A5 = $($sheet.Range('A5').Value2)"
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    if ($book) { $book.Close($false) }             # $false = 保存しない（元ファイルは変わらない）
    $excel.Quit()
    $sheet = $null
    $size  = $null
    $book  = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
