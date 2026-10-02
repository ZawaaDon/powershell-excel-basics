# ============================================================
# step6-macro : step6.bat から呼ばれ、マクロで計算した値を返すスクリプト
# ------------------------------------------------------------
# やること
#   1. macros.xlsm の AddTwo で A + B を計算する
#   2. その結果を macros.xlam の SumArray に渡して、10 と 0.5 を足す
#   3. 結果を1行だけ出力する（bat がこの1行を受け取る）
#
# 引数（どちらも省略できる）
#   -A, -B : 足す数（省略時は 3 と 4）
#
# 呼び出し方
#   step6.bat : for /f で出力を受け取る
#
# 覚えておくこと
#   - bat は、このスクリプトが画面に出したものをすべて受け取る（Write-Host も含む）。
#     返したい値のほかは何も出力しない
#   - エラーで止まると何も出力されず、終了コードが 1 になる
#     （$ErrorActionPreference = "Stop" を書いているため）
#   - マクロの呼び方は step2.ps1（.xlsm）と step5.ps1（.xlam）を参照
# ============================================================
param(
    [double]$A = 3,
    [double]$B = 4
)

# --- エラーが起きたら、その場でスクリプトを止める ---
# 書かないと、エラーを表示したあと次の行へ進み、最後まで実行して「成功」で終わる（step7.ps1 を参照）
$ErrorActionPreference = "Stop"

$macroPath = "E:\dev\excel\macros.xlsm"
$addinPath = "E:\dev\excel\macros.xlam"

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- マクロのブックとアドインを開く ---
    $macroBook = $excel.Workbooks.Open($macroPath, 0, $true)    # 第3引数 $true = 読み取り専用
    $addin     = $excel.Workbooks.Open($addinPath)

    # --- マクロで計算する ---
    $sum   = $excel.Run("'macros.xlsm'!AddTwo", $A, $B)             # .xlsm のマクロ
    $total = $excel.Run("'macros.xlam'!SumArray", @($sum, 10, 0.5)) # .xlam のマクロ
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    if ($macroBook) { $macroBook.Close($false) }
    if ($addin)     { $addin.Close($false) }
    $excel.Quit()
    $macroBook = $null
    $addin     = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --- 戻り値（bat が受け取る） ---
$total
