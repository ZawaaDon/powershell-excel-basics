# ============================================================
# step6-sub : step6 から呼ばれる側のスクリプト
# ------------------------------------------------------------
# やること
#   sample-step6.xlsx の指定セルに文字を書き、結果を1行返す
#
# 引数（どちらも省略できる）
#   -Addr : 書き込み先のセル（省略時は G1）
#   -Text : 書き込む文字（省略時は「引数なし」）
#
# 呼び出し方
#   step6.ps1 : PowerShell のスクリプトから呼ぶ
#   step6.bat : bat ファイルから呼ぶ
#
# 覚えておくこと
#   - param(...) はファイルの先頭に書く（上に置けるのはコメントだけ）
#   - 「= 値」を書いた引数は、省略されたときにその値になる
#   - そのまま出力した値が呼び出し元への戻り値になる。最後の行だけでなく、途中で出力したものも全部入る
#     （return は「その値を出力して終わる」。C のように return の値だけが返るわけではない）
#   - 戻り値のあるメソッドを変数に入れずに呼ぶと、その値も混ざる。いらない出力は | Out-Null で捨てる
#     Write-Host の表示は戻り値に入らない
# ============================================================
param(
    [string]$Addr = "G1",
    [string]$Text = "引数なし"
)

# --- エラーが起きたら、その場でスクリプトを止める ---
# 書かないと、エラーを表示したあと次の行へ進み、最後まで実行して「成功」で終わる（step7.ps1 を参照）
$ErrorActionPreference = "Stop"

# --- 共通関数を読み込む（step3.ps1 を参照） ---
. "$PSScriptRoot\excel-lib.ps1"

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath = "E:\dev\excel\sample-org.xlsx"
$dstPath = "E:\dev\excel\sample-step6.xlsx"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- ブックとシートを開く ---
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")

    # --- 引数で受け取ったセルに、引数で受け取った文字を書く ---
    Write-CellText $sheet $Addr $Text

    # --- 保存 ---
    $book.Save()
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    if ($book) { $book.Close($false) }
    $excel.Quit()
    $sheet = $null
    $book  = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --- 戻り値（呼び出し元が受け取る） ---
"$Addr に「$Text」を書きました"
