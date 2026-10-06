# ============================================================
# step9 : セルへの書き込みを詳しく
# ------------------------------------------------------------
# 学ぶこと
#   1. 範囲にまとめて書く（2次元配列、Resize、Offset）
#   2. 値の種類ごとの注意（先頭の 0、日付、勝手に変換される文字、空欄、エラー値、変数の埋め込み）
#   3. 数式の書き方（範囲にまとめて入れる、R1C1 形式、引用符と $、値だけにする）
#   4. 書式と消去（表示形式、罫線、色、配置、列幅、ClearContents と Clear）
#   5. 大量に書くときに速くする（まとめて書く、画面更新と再計算を止める）
#
# 使うファイル
#   sample-org.xlsx   : 元データ（書き換えない）
#   sample-step9.xlsx : 毎回 sample-org.xlsx からコピーして作る作業用
#
# 実行方法
#   .\step9.ps1
#
# 覚えておくこと
#   - たくさんのセルに書くときは、2次元配列を作って範囲に 1 回で書く。
#     1 セルずつ書くと、1 セルごとに Excel とやり取りするので遅い
#   - 2次元配列は New-Object 'object[,]' 行数, 列数 で作る（添字は 0 始まり）。
#     @(@(1,2), @(3,4)) は「配列の配列」で、範囲に書くとエラー値が入る（エラーにはならない）
#   - Excel は書いた文字を、手で入力したときと同じように解釈する。
#     "001" は数値の 1 に、"1/2" は日付（1月2日）になる
#   - 数式は '...'（単一引用符）で書く。"..." の中の $A$1 は PowerShell の変数とみなされる
#   - 文字列に変数を埋め込むときは "..." で書く（"$a, $b"）。計算やプロパティは $( ) で囲む
# ============================================================

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath = "E:\dev\excel\sample-org.xlsx"
$dstPath = "E:\dev\excel\sample-step9.xlsx"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")

    # ============================================================
    # 1. 範囲にまとめて書く（Sheet1 の G 列から）
    # ============================================================
    Write-Host "--- 1. 範囲にまとめて書く"

    # --- 2次元配列を作って、範囲に 1 回で書く ---
    # 配列の大きさ（3 行 3 列）と、書く範囲の大きさ（G1:I3）をそろえる
    $data = New-Object 'object[,]' 3, 3          # 添字は 0 始まり：[行, 列]
    $data[0, 0] = "名前";   $data[0, 1] = "国語"; $data[0, 2] = "数学"
    $data[1, 0] = "太郎";   $data[1, 1] = 80;     $data[1, 2] = 65
    $data[2, 0] = "花子";   $data[2, 1] = 72;     $data[2, 2] = 90
    $sheet.Range("G1:I3").Value2 = $data

    # --- 範囲の大きさを配列から決める（Resize） ---
    # Range("G5") を起点に、配列と同じ行数・列数に広げる。番地を自分で数えなくてよい
    # GetLength(0) が行数、GetLength(1) が列数
    $sheet.Range("G5").Resize($data.GetLength(0), $data.GetLength(1)).Value2 = $data

    # --- 起点からずらす（Offset） ---
    # Offset(行, 列) は「何行下・何列右か」。G1 から 1 行下・3 列右は J2
    $sheet.Range("G1").Offset(0, 3).Value2 = "合計"
    $sheet.Range("G1").Offset(1, 3).Value2 = 145
    $sheet.Range("G1").Offset(2, 3).Value2 = 162

    # --- 縦 1 列に書くときも 2次元配列にする ---
    # 1次元配列 @("a","b","c") は「横 1 行」とみなされる。縦の範囲に書くと全部 "a" になる
    $names = @("一郎", "二郎", "三郎")
    $col = New-Object 'object[,]' $names.Count, 1
    for ($i = 0; $i -lt $names.Count; $i++) { $col[$i, 0] = $names[$i] }
    $sheet.Range("K1").Resize($names.Count, 1).Value2 = $col

    # --- 同じ値を範囲全体に書く ---
    $sheet.Range("L1:L3").Value2 = "済"

    # --- 範囲をまとめて読むと、2次元配列で返る（添字は 1 始まり） ---
    # 書くときの配列は 0 始まり、読んだ配列は 1 始まりなので注意
    $back = $sheet.Range("G1:I3").Value2
    Write-Host "G1:I3 を読んだ配列の [2, 1] = $($back[2, 1])"    # 2 行目・1 列目 = 太郎

    # ============================================================
    # 2. 値の種類ごとの注意（「値の種類」シート）
    # ============================================================
    Write-Host "--- 2. 値の種類"
    $vs = $book.Worksheets.Add()
    $vs.Name = "値の種類"
    $title = New-Object 'object[,]' 1, 3         # 横 1 行も 2次元配列で作れる
    $title[0, 0] = "書いたもの"; $title[0, 1] = "入った値"; $title[0, 2] = "表示"
    $vs.Range("A1:C1").Value2 = $title

    # 書いた文字と、Excel に入った値・表示を並べる
    # A 列には '（単一引用符）を付けて、書いた文字をそのまま見せる
    $samples = @("001", "'001", "1/2", "1-2", "12345678901234567890", "TRUE", "=1+1")
    $row = 2
    foreach ($s in $samples) {
        $vs.Cells.Item($row, 1).Value2 = "'" + $s
        $vs.Cells.Item($row, 2).Value2 = $s
        $vs.Cells.Item($row, 3).Value2 = "'" + $vs.Cells.Item($row, 2).Text
        Write-Host ("{0,-22} → 入った値 {1}（{2}）" -f $s, $vs.Cells.Item($row, 2).Value2, $vs.Cells.Item($row, 2).Value2.GetType().Name)
        $row++
    }

    # --- 先頭の 0 を残す方法は 2 つ ---
    $vs.Range("E1").Value2 = "'001"                # 方法1：' を付ける（' は表示されない）
    $vs.Range("E2").NumberFormat = "@"             # 方法2：先に表示形式を「文字列」にしてから書く
    $vs.Range("E2").Value2 = "001"
    Write-Host "E1 = $($vs.Range('E1').Value2)  E2 = $($vs.Range('E2').Value2)"

    # --- 日付は Value で書く ---
    # Value2 で書くと、日付のシリアル値（ただの数値）が入り、表示も数値のまま
    # Value で書くと、日付の表示形式も付く
    $dt = [datetime]"2026-10-06 13:45"            # Get-Date -Year … だとミリ秒が実行時の値で残る
    $vs.Range("E4").Value2 = $dt
    $vs.Range("E5").Value  = $dt
    $vs.Range("E6").Value2 = $dt.ToOADate()        # 数値で書いて、表示形式を自分で付けてもよい
    $vs.Range("E6").NumberFormat = "yyyy/mm/dd hh:mm"
    $vs.Columns.Item(5).AutoFit() | Out-Null
    Write-Host "Value2 で書いた日付の表示 = $($vs.Range('E4').Text)"
    Write-Host "Value  で書いた日付の表示 = $($vs.Range('E5').Text)"
    Write-Host "数値と表示形式で書いた表示 = $($vs.Range('E6').Text)"

    # Value は「引数を取れるプロパティ」なので、読むときは Value() と書く
    # （$vs.Range("E5").Value と書くと、値ではなくプロパティの説明が返る）
    Write-Host "Value() で読む = $($vs.Range('E5').Value())（$($vs.Range('E5').Value().GetType().Name)）"

    # --- 空欄にする ---
    # $null を書くと空欄になる。"" を書いても空欄になる
    $vs.Range("E8").Value2 = "消す前"
    $vs.Range("E8").Value2 = $null
    Write-Host "E8 は空欄 = $($null -eq $vs.Range('E8').Value2)"

    # --- エラー値を読むと、整数が返る ---
    # #DIV/0! のセルを Value2 で読むと -2146826281 が返り、エラーにはならない。
    # 計算結果を使う前に、Text でエラー表示になっていないか確かめる
    $vs.Range("E9").Formula = "=1/0"
    $e = $vs.Range("E9")
    Write-Host "E9 の Value2 = $($e.Value2)  Text = $($e.Text)"
    if ($e.Text.StartsWith("#")) { Write-Host "E9 はエラー値です" }

    # --- 変数を埋め込んだ文字列を書く（G 列） ---
    # "..." の中の $変数 は値に置き換わり、1 つの文字列として 1 つのセルに入る。
    # '...' の中では置き換わらない（数式を書くときとちょうど逆）。
    # 置き換わった後の文字が数値や日付に見えると、Excel が変換するのは "001" と同じ
    $a = "りんご"; $b = "みかん"; $c = "ぶどう"
    $x = 1; $y = 2; $z = 3
    $vs.Range("G1").Value2  = "$a, $b, $c"                  # りんご, みかん, ぶどう
    $vs.Range("G2").Value2  = '$a, $b, $c'                  # 置き換わらず $a, $b, $c のまま
    $vs.Range("G3").Value2  = "$x, $y, $z"                  # 1, 2, 3（文字列のまま）
    $vs.Range("G4").Value2  = "$x,$y,$z"                    # 1,2,3（文字列のまま）
    $vs.Range("G5").Value2  = "$x,234"                      # 1,234 → 数値の 1234 になる
    $vs.Range("G6").Value2  = "$x/$y"                       # 1/2 → 日付（1月2日）になる
    $vs.Range("G7").Value2  = "'$x/$y"                      # ' を付けると文字列のまま
    $vs.Range("G8").Value2  = "合計は $($x + $y + $z) 個"   # 計算は $( ) で囲む
    $vs.Range("G9").Value2  = "シート名は $($vs.Name)"      # プロパティも $( ) で囲む
    $vs.Range("G10").Value2 = "シート名は $vs.Name"         # 囲まないと $vs だけが置き換わる
    $vs.Range("G11").Value2 = @("x", "y", "z") -join ", "   # 配列は -join で区切り文字を決めてつなぐ
    $vs.Range("G12").Value2 = "$x`n$y`n$z"                  # `n でセルの中で改行する

    # 配列の要素も $( ) で囲む。"$names[0]" は $names だけが置き換わり、[0] は文字のまま残る
    # $names は 1次元配列（一郎, 二郎, 三郎）、$data は 2次元配列（1. で作った表）
    $vs.Range("G13").Value2 = "$names[0], $names[1]"                         # 一郎 二郎 三郎[0], …
    $vs.Range("G14").Value2 = "$($names[0]), $($names[1]), $($names[2])"     # 一郎, 二郎, 三郎
    $vs.Range("G15").Value2 = $names -join ", "                              # 一郎, 二郎, 三郎
    $vs.Range("G16").Value2 = $names[0..1] -join ", "                        # 0〜1 番目だけ：一郎, 二郎
    $vs.Range("G17").Value2 = "$data[1,0], $data[1,1]"                       # 名前 国語 … 90[1,0], …
    $vs.Range("G18").Value2 = "$($data[1,0]), $($data[1,1]), $($data[1,2])"  # 太郎, 80, 65
    $vs.Range("G19").Value2 = $data -join ", "                               # 全部が 1 列に並ぶ
    $vs.Range("G20").Value2 = (0..2 | ForEach-Object { $data[1, $_] }) -join ", "   # 1 行目だけ：太郎, 80, 65
    for ($r = 1; $r -le 20; $r++) {
        $v = $vs.Cells.Item($r, 7).Value2
        Write-Host ("G{0,-2} = [{1}]（{2}）" -f $r, ($v -replace "`n", "\n"), $v.GetType().Name)
    }

    # ============================================================
    # 3. 数式の書き方（「数式」シート）
    # ============================================================
    Write-Host "--- 3. 数式"
    $fs = $book.Worksheets.Add()
    $fs.Name = "数式"
    $nums = New-Object 'object[,]' 3, 2
    $nums[0, 0] = 100; $nums[0, 1] = 3
    $nums[1, 0] = 250; $nums[1, 1] = 2
    $nums[2, 0] = 80;  $nums[2, 1] = 5
    $fs.Range("A1:B3").Value2 = $nums

    # --- 範囲にまとめて数式を入れる ---
    # 先頭のセル（C1）に書く形で 1 つ書けば、下のセルは行がずれた数式になる（C2 は =A2*B2）
    $fs.Range("C1:C3").Formula = "=A1*B1"

    # --- R1C1 形式（相対位置で書く） ---
    # RC[-1] は「同じ行の 1 列左」。どのセルにも同じ文字列で書けるので、ループで作りやすい
    $fs.Range("D1:D3").FormulaR1C1 = "=RC[-1]*2"

    # --- 引用符と $ を含む数式は '...'（単一引用符）で書く ---
    # "=SUM($A$1:$A$3)" と書くと、PowerShell が $A を変数とみなし、スクリプト全体が動かない
    # '...' の中では、数式の中の "..." もそのまま書ける
    $fs.Range("E1").Formula = '=SUM($C$1:$C$3)'
    $fs.Range("E2").Formula = '=IF(E1>1000,"多い","少ない")'

    # 変数を数式に埋め込みたいときは "..." を使い、数式の $ は `$ と書く
    $lastRow = 3
    $fs.Range("E3").Formula = "=AVERAGE(`$A`$1:`$A`$$lastRow)"

    for ($r = 1; $r -le 3; $r++) {
        Write-Host ("C{0} {1,-10} D{0} {2,-10} = {3}" -f $r, $fs.Cells.Item($r, 3).Formula, $fs.Cells.Item($r, 4).Formula, $fs.Cells.Item($r, 4).Value2)
    }
    Write-Host "E1 $($fs.Range('E1').Formula) = $($fs.Range('E1').Value2)"
    Write-Host "E2 $($fs.Range('E2').Formula) = $($fs.Range('E2').Value2)"
    Write-Host "E3 $($fs.Range('E3').Formula) = $($fs.Range('E3').Value2)"

    # --- 数式を値だけにする（計算結果で上書きする） ---
    # 読んだ値を同じ範囲に書き戻す。D 列は数式が消えて、数値だけが残る
    $d = $fs.Range("D1:D3")
    $d.Value2 = $d.Value2
    Write-Host "値にした後の D1 の Formula = $($fs.Range('D1').Formula)"

    # --- 数式の書き間違いは、0x800A03EC というエラーになる ---
    try {
        $fs.Range("F1").Formula = "=SUM("
    }
    catch {
        Write-Host "数式を入れられませんでした : $($_.Exception.Message)"
    }

    # ============================================================
    # 4. 書式と消去（Sheet1 の G1:J7）
    # ============================================================
    Write-Host "--- 4. 書式と消去"
    $table = $sheet.Range("G1:J3")
    $head  = $sheet.Range("G1:J1")

    # --- 表示形式（NumberFormat） ---
    # Excel の「セルの書式設定 → 表示形式 → ユーザー定義」と同じ文字列
    $sheet.Range("H2:J3").NumberFormat = "#,##0"

    # --- 罫線（Borders） ---
    # LineStyle 1 = 実線（xlContinuous）。Weight 2 = 細線、-4138 = 中太線
    $table.Borders.LineStyle = 1
    $table.Borders.Weight    = 2

    # --- 色（Color） ---
    # Excel の色は 赤 + 緑*256 + 青*65536 で表す（RGB の順と逆に重みが付く）
    function Get-ExcelColor($r, $g, $b) { return $r + $g * 256 + $b * 65536 }
    $head.Interior.Color = Get-ExcelColor 221 235 247     # 薄い青
    $head.Font.Color     = Get-ExcelColor 0 32 96         # 紺
    $head.Font.Bold      = $true

    # --- 配置と列幅 ---
    # HorizontalAlignment : -4108 = 中央、-4131 = 左、-4152 = 右
    $head.HorizontalAlignment = -4108
    $sheet.Columns.Item(7).ColumnWidth = 12         # G 列の幅（文字数の目安）
    $sheet.Range("H:J").Columns.AutoFit() | Out-Null

    Write-Host "J2 の表示 = $($sheet.Range('J2').Text)  背景色 = $($head.Interior.Color)"

    # --- 消去 ---
    # ClearContents : 値と数式だけ消す（書式は残る）
    # Clear         : 書式も含めてすべて消す
    # ClearFormats  : 書式だけ消す（値は残る）
    $copy = $sheet.Range("G5:I7")
    $copy.Interior.Color = Get-ExcelColor 255 242 204
    $sheet.Range("G6:I6").ClearContents() | Out-Null
    Write-Host "ClearContents 後の G6 : 値 = [$($sheet.Range('G6').Value2)]  背景色 = $($sheet.Range('G6').Interior.Color)"
    $sheet.Range("G7:I7").Clear() | Out-Null
    Write-Host "Clear 後の G7         : 値 = [$($sheet.Range('G7').Value2)]  背景色 = $($sheet.Range('G7').Interior.Color)"

    # ============================================================
    # 5. 大量に書くときに速くする（「速さ」シート）
    # ============================================================
    Write-Host "--- 5. 速さ（300 行 × 5 列）"
    $ps = $book.Worksheets.Add()
    $ps.Name = "速さ"
    $rows = 300
    $cols = 5

    # --- A: 1 セルずつ書く ---
    $t1 = Measure-Command {
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $ps.Cells.Item($r, $c).Value2 = $r * $c
            }
        }
    }

    # --- B: 画面更新と再計算を止めて、1 セルずつ書く ---
    # Calculation : -4135 = 手動、-4105 = 自動（ブックを開いているときだけ設定できる）
    $excel.ScreenUpdating = $false
    $excel.Calculation    = -4135
    try {
        $t2 = Measure-Command {
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $ps.Cells.Item($r, $c + 6).Value2 = $r * $c
                }
            }
        }
    }
    finally {
        # 必ず元に戻す。手動計算のまま保存すると、次に開いたときも手動計算になる
        $excel.Calculation    = -4105
        $excel.ScreenUpdating = $true
    }

    # --- C: 2次元配列を作って、1 回で書く ---
    $t3 = Measure-Command {
        $big = New-Object 'object[,]' $rows, $cols
        for ($r = 0; $r -lt $rows; $r++) {
            for ($c = 0; $c -lt $cols; $c++) {
                $big[$r, $c] = ($r + 1) * ($c + 1)
            }
        }
        $ps.Range("M1").Resize($rows, $cols).Value2 = $big
    }

    Write-Host ("A: 1 セルずつ                  {0,6:N0} ミリ秒" -f $t1.TotalMilliseconds)
    Write-Host ("B: 画面更新・再計算を止める    {0,6:N0} ミリ秒" -f $t2.TotalMilliseconds)
    Write-Host ("C: 2次元配列で 1 回            {0,6:N0} ミリ秒" -f $t3.TotalMilliseconds)

    # --- 保存 ---
    $sheet.Activate()                              # 開いたときに Sheet1 が出るようにする
    $book.Save()
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    if ($book) { $book.Close($false) }
    $excel.Quit()
    $sheet = $null
    $vs    = $null
    $fs    = $null
    $ps    = $null
    $table = $null
    $head  = $null
    $copy  = $null
    $d     = $null
    $e     = $null
    $book  = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
