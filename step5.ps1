# ============================================================
# step5 : アドイン（.xlam）に書いたマクロを呼び出す
# ------------------------------------------------------------
# 学ぶこと
#   - アドインを Workbooks.Open() で開いてから、$excel.Run() でマクロを呼ぶ
#   - 呼び方は step2 と同じ（'ブック名'!マクロ名）
#   - ファイルの場所は UNC パス（\\サーバー名\共有名\...）でも指定できる
#
# 使うファイル
#   sample-org.xlsx   : 元データ（書き換えない）
#   sample-step5.xlsx : 毎回 sample-org.xlsx からコピーして作る作業用（step ごとに別の名前）
#   macros.xlam       : マクロを書いたアドイン（中身は vba\Module1.bas と同じ）
#                         BoldHeader / WriteText / AddTwo / SumArray / SafeCalc / UnsafeSqrt
#
# 実行方法
#   .\step5.ps1
#
# 覚えておくこと
#   - PowerShell から起動した Excel は、Excel に登録してあるアドインを自動では読み込まない。
#     使う前に Workbooks.Open() で開く
#   - 登録済みアドインの場所は $excel.AddIns で調べられる（各要素の FullName がパス）
#   - アドインは非表示のブック。開いてもアクティブなシートは変わらず、
#     $excel.Workbooks の一覧にも出ない
#   - Open() の戻り値を変数に取っておき、最後に閉じて $null にする
#   - UNC パスで開いても、呼び出し名はファイル名だけで書く（'macros.xlam'!AddTwo）
#   - マクロの制約は step2.ps1 の「覚えておくこと」と同じ
#
# アドイン（.xlam）の作り方
#   マクロを書いたブックを Excel で開き、名前を付けて保存 →
#   ファイルの種類「Excel アドイン (*.xlam)」で保存する
# ============================================================

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath   = "E:\dev\excel\sample-org.xlsx"
$dstPath   = "E:\dev\excel\sample-step5.xlsx"
$addinPath = "E:\dev\excel\macros.xlam"
# UNC パスで書く場合の例（アドインを共有フォルダに置いたとき）
# $addinPath = "\\サーバー名\共有名\excel\macros.xlam"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- アドインと、処理対象のブックを両方開く ---
    $addin = $excel.Workbooks.Open($addinPath)                  # アドインは自分で開く
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")

    # --- マクロが操作する対象を選んでおく ---
    # マクロは ActiveSheet に対して動くので、呼ぶ前に対象を決める
    $book.Activate()
    $sheet.Activate()

    # --- マクロを呼ぶ：'アドイン名'!マクロ名 ---
    $macroName = "'" + $addin.Name + "'!"                       # 'macros.xlam'!
    $excel.Run($macroName + "BoldHeader")                       # 引数なし
    # 書き込み先は表（A〜E列）の外の空いているセルにして、元のデータを上書きしない
    $excel.Run($macroName + "WriteText", "G1", "アドインから")   # 引数あり
    $sum = $excel.Run($macroName + "AddTwo", 3, 4)              # 戻り値あり
    Write-Host "AddTwo = $sum"
    $total = $excel.Run($macroName + "SumArray", @(1, 2, 3))    # 配列を渡す
    Write-Host "SumArray = $total"

    # --- 保存 ---
    $book.Save()
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    # $addin も Excel への参照なので、閉じて $null にする
    if ($book)  { $book.Close($false) }
    if ($addin) { $addin.Close($false) }                        # アドインは保存しない
    $excel.Quit()
    $sheet = $null
    $book  = $null
    $addin = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
