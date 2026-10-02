# ============================================================
# step2 : 別のブック（.xlsm）に書いたマクロを呼び出す
# ------------------------------------------------------------
# 学ぶこと
#   - $excel.Run("'ブック名'!マクロ名", 引数...) でマクロを呼ぶ
#   - 引数なし / 引数あり / 戻り値ありの呼び方
#   - 配列を引数として渡す呼び方
#   - 呼び出しを短く書くための関数 Invoke-Macro
#
# 使うファイル
#   sample-org.xlsx   : 元データ（書き換えない）
#   sample-step2.xlsx : 毎回 sample-org.xlsx からコピーして作る作業用（step ごとに別の名前）
#   macros.xlsm       : マクロを書いたブック（標準モジュール Module1）
#                         Public Sub BoldHeader()
#                         Public Sub WriteText(addr As String, text As String)
#                         Public Function AddTwo(a As Double, b As Double) As Double
#                         Public Function SumArray(values As Variant) As Double
#                         Public Function SafeCalc(kind As String, x As Double) As Variant   （step7 で使う）
#                         Public Function UnsafeSqrt(x As Double) As Double                  （step7 で使う）
#
# 実行方法
#   .\step2.ps1
#
# 覚えておくこと
#   - マクロのブックは同じ $excel で開いておく（別の Excel で開いていても呼べない）
#   - 呼び出し名はファイル名で書く（フルパスではない）
#   - 呼べるのは標準モジュールの Public な Sub / Function
#   - 同じ名前のマクロが複数のモジュールにあるときは 'macros.xlsm'!Module1.BoldHeader と書く
#   - マクロ内の ActiveSheet は「そのとき選ばれているシート」、
#     ThisWorkbook は macros.xlsm 自身を指す
#   - マクロに MsgBox があると、画面非表示のまま待ち続けて止まる
#
# 関連
#   step3.ps1 : 同じ処理を PowerShell の関数だけで書く（.xlsm 不要）
#   step4.ps1 : VBA を .bas（テキスト）で管理して取り込む（.xlsm 不要）
#   step5.ps1 : アドイン（.xlam）のマクロを呼ぶ
#   step7.ps1 : マクロでエラーが起きたときの扱い方
# ============================================================

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath   = "E:\dev\excel\sample-org.xlsx"
$dstPath   = "E:\dev\excel\sample-step2.xlsx"
$macroPath = "E:\dev\excel\macros.xlsm"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

# --- マクロを短く呼ぶための関数 ---
# Invoke-Macro マクロ名 引数1 引数2 ... と書くと、
# $excel.Run("'macros.xlsm'!マクロ名", 引数1, 引数2, ...) を実行する
function Invoke-Macro {
    $target    = "'$($macroBook.Name)'!$($args[0])"             # ブック名は $macroBook から取る
    $macroArgs = @()                                            # 2個目以降がマクロの引数
    if ($args.Count -gt 1) { $macroArgs = $args[1..($args.Count - 1)] }
    # 引数の個数が決まっていないので、配列を展開して Run に渡す
    $excel.GetType().InvokeMember('Run', 'InvokeMethod', $null, $excel, @($target) + $macroArgs)
}

try {
    # --- マクロのブックと、処理対象のブックを両方開く ---
    $macroBook = $excel.Workbooks.Open($macroPath, 0, $true)    # 第3引数 $true = 読み取り専用
    $book      = $excel.Workbooks.Open($dstPath)
    $sheet     = $book.Worksheets.Item("Sheet1")

    # --- マクロが操作する対象を選んでおく ---
    # マクロは ActiveSheet に対して動くので、呼ぶ前に対象を決める
    $book.Activate()
    $sheet.Activate()

    # --- マクロを呼ぶ：'ブック名'!マクロ名 ---
    $excel.Run("'macros.xlsm'!BoldHeader")                      # 引数なし
    # 書き込み先は表（A〜E列）の外の空いているセルにして、元のデータを上書きしない
    $excel.Run("'macros.xlsm'!WriteText", "G1", "マクロから")    # 引数あり
    $sum = $excel.Run("'macros.xlsm'!AddTwo", 3, 4)             # 戻り値あり
    Write-Host "AddTwo = $sum"

    # --- 同じことを Invoke-Macro で短く書く ---
    # カッコとカンマは使わず、空白で区切る（AddTwo(5, 6) とは書かない）
    Invoke-Macro BoldHeader                                     # 引数なし
    Invoke-Macro WriteText "G2" "関数から"                     # 引数あり
    $sum2 = Invoke-Macro AddTwo 5 6                             # 戻り値あり
    Write-Host "AddTwo = $sum2"

    # --- 配列を引数として渡す ---
    # 配列は「引数1個」として渡る。マクロ側は Variant で受ける（values As Variant）
    # マクロに届いた配列の添字は 0 から始まる
    $numbers = 1, 2, 3
    $total = $excel.Run("'macros.xlsm'!SumArray", $numbers)     # $excel.Run で渡す
    Write-Host "SumArray = $total"
    $total2 = Invoke-Macro SumArray $numbers                    # Invoke-Macro で渡す
    Write-Host "SumArray = $total2"
    $total3 = Invoke-Macro SumArray @(10, 20, 30)               # その場で配列を書いてもよい
    Write-Host "SumArray = $total3"

    # --- 保存 ---
    $book.Save()
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    # $macroBook も Excel への参照なので、閉じて $null にする
    if ($book)      { $book.Close($false) }
    if ($macroBook) { $macroBook.Close($false) }                # マクロのブックは保存しない
    $excel.Quit()
    $sheet     = $null
    $book      = $null
    $macroBook = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
