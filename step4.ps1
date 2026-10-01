# ============================================================
# step4 : VBA をテキスト（.bas）で管理し、実行時に取り込んで呼ぶ
# ------------------------------------------------------------
# 学ぶこと
#   - VBA のコードを .bas ファイルとして持つ（VSCode で編集でき、差分も取れる）
#   - 実行時に一時ブックへ取り込み、$excel.Run() で呼ぶ
#   - step2 と同じ処理を、.xlsm なしで実現する
#
# 使うファイル
#   sample-org.xlsx : 元データ（書き換えない）
#   sample.xlsx     : 毎回 sample-org.xlsx からコピーして作る作業用
#   vba\*.bas       : VBA のコード（このフォルダの .bas をすべて取り込む）
#                       Module1.bas : BoldHeader / WriteText / AddTwo
#
# 事前の設定（必須。オフだと「VBA プロジェクトにアクセスできません」で止まる）
#   Excel → ファイル → オプション → トラスト センター → トラスト センターの設定
#   → マクロの設定 →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオン
#   ※ セキュリティを下げる設定。使わなくなったらオフに戻す
#
# 実行方法
#   .\step4.ps1
#
# .bas を書くときの注意
#   - 1行目の Attribute VB_Name = "Module1" がモジュール名になる（消さない）
#   - 改行は CRLF にする（VSCode 右下の表示で確認）
#   - コメントは英語にしておく（この PC はコードページが UTF-8 で、日本語は化ける可能性がある）
#   - マクロの制約は step2.ps1 の「覚えておくこと」と同じ
#
# step3 との使い分け
#   新しく書くなら step3（PowerShell の関数）のほうが手間が少ない。
#   step4 は、すでにある VBA をそのまま使いたいとき向け。
# ============================================================

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath = "E:\dev\excel\sample-org.xlsx"
$dstPath = "E:\dev\excel\sample.xlsx"
$vbaDir  = "$PSScriptRoot\vba"                  # VBA のコード（.bas）を置くフォルダ
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- マクロを入れるための一時ブックを作り、.bas を取り込む ---
    # 一時ブックは保存せずに閉じるので、.xlsm ファイルは作られない
    $macroBook  = $excel.Workbooks.Add()
    $components = $macroBook.VBProject.VBComponents
    if ($null -eq $components) {
        # 信頼の設定がオフだと、エラーにならず $null が返る
        throw "VBA プロジェクトにアクセスできません。Excel のオプション → トラスト センター → マクロの設定 で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにしてください。"
    }
    foreach ($bas in Get-ChildItem $vbaDir -Filter *.bas) {
        $components.Import($bas.FullName) | Out-Null
    }
    $macroName = "'" + $macroBook.Name + "'!"   # 例：'Book1'!（一時ブックの名前は実行時に決まる）

    # --- 処理対象のブックを開く ---
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")

    # --- マクロが操作する対象を選んでおく ---
    # マクロは ActiveSheet に対して動くので、呼ぶ前に対象を決める
    $book.Activate()
    $sheet.Activate()

    # --- マクロを呼ぶ：'ブック名'!マクロ名 ---
    $excel.Run($macroName + "BoldHeader")                       # 引数なし
    $excel.Run($macroName + "WriteText", "D1", "マクロから")     # 引数あり
    $sum = $excel.Run($macroName + "AddTwo", 3, 4)              # 戻り値あり
    Write-Host "AddTwo = $sum"

    # --- 保存 ---
    $book.Save()
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    # $components と $macroBook も Excel への参照なので $null にする
    if ($book)      { $book.Close($false) }
    if ($macroBook) { $macroBook.Close($false) }                # 一時ブックは保存しない
    $excel.Quit()
    $sheet      = $null
    $book       = $null
    $components = $null
    $macroBook  = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
