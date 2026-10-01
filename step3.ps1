# ============================================================
# step3 : マクロを使わず、PowerShell の関数を別ファイルに分ける
# ------------------------------------------------------------
# 学ぶこと
#   - 共通の処理を関数にして別ファイル（excel-lib.ps1）に置く
#   - 「. ファイル」で読み込んで呼び出す（ドットソース）
#   - step2 と同じ処理を、.xlsm なしで実現する
#
# 使うファイル
#   sample-org.xlsx : 元データ（書き換えない）
#   sample.xlsx     : 毎回 sample-org.xlsx からコピーして作る作業用
#   excel-lib.ps1   : 共通関数（Set-BoldHeader / Write-CellText / Add-Two）
#
# 実行方法
#   .\step3.ps1
#
# 覚えておくこと
#   - 関数の呼び出しはカッコとカンマを使わず、空白で区切る
#     （Add-Two(3, 4) と書くと「配列1個を渡した」ことになる）
#   - $PSScriptRoot は実行中のスクリプトがあるフォルダ
#   - シートを引数で渡すので、ActiveSheet に頼らない（Activate() が不要）
#
# step2 との違い
#   - マクロのブックを開く・閉じる・解放する処理がいらない
#   - コードがすべてテキストなので、VSCode で編集でき、差分も取れる
# ============================================================

# --- 共通関数を読み込む（先頭の「. 」で別ファイルの関数が使えるようになる） ---
. "$PSScriptRoot\excel-lib.ps1"

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath = "E:\dev\excel\sample-org.xlsx"
$dstPath = "E:\dev\excel\sample.xlsx"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- ブックとシートを開く ---
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")

    # --- 関数を呼ぶ（カッコとカンマは使わず、空白で区切る） ---
    Set-BoldHeader $sheet                       # 引数はシートだけ
    Write-CellText $sheet "D1" "関数から"       # 引数あり
    $sum = Add-Two 3 4                          # 戻り値あり
    Write-Host "Add-Two = $sum"

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
