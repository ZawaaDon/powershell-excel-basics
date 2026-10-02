# ============================================================
# step1 : PowerShell から Excel を操作する基本
# ------------------------------------------------------------
# 学ぶこと
#   - Excel の起動と終了（COM オブジェクト）
#   - セルの読み書き、最終行・最終列の求め方、ループ
#   - シートの追加、数式、書式
#   - 後始末（Excel のプロセスを残さない）
#
# 使うファイル
#   sample-org.xlsx   : 元データ（書き換えない）
#   sample-step1.xlsx : 毎回 sample-org.xlsx からコピーして作る作業用（step ごとに別の名前）
#
# 実行方法
#   .\step1.ps1
#
# 覚えておくこと
#   - 行・列・シート番号はすべて 1 始まり
#   - パスは絶対パスで指定する（Excel のカレントフォルダは PowerShell と別）
#   - プロパティ名を間違えて「読む」とエラーにならず $null が返る
#     （例：$book.WorkSheet.Add() は「null に対してメソッドを呼べない」で止まる）
# ============================================================

# --- 元ファイルをコピーしてから処理する（毎回同じ状態から始める） ---
$srcPath = "E:\dev\excel\sample-org.xlsx"
$dstPath = "E:\dev\excel\sample-step1.xlsx"
Copy-Item $srcPath $dstPath -Force -ErrorAction Stop    # 失敗したらここで止める

# --- Excel を起動する ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    # --- ブックとシートを開く ---
    # 変数は代入した時点のシートを指し続ける（Excel 上でどのシートが選ばれていても変わらない）
    # $book = $excel.Workbooks.Add()               # 新規ブックの場合
    $book  = $excel.Workbooks.Open($dstPath)
    $sheet = $book.Worksheets.Item("Sheet1")       # 名前 or 番号（1始まり）

    # --- 読み込み ---
    $v = $sheet.Range("A1").Value2                 # 生の値（日付はシリアル値の数値で返る）
    $t = $sheet.Range("A2").Text                   # 表示されている文字列
    Write-Host "A1 = $v"
    Write-Host "A2 = $t"

    # --- 最終行・最終列を求める ---
    # VBA の定数名（xlUp など）は使えないので、数値で指定する
    $xlUp     = -4162
    $xlToLeft = -4159
    $lastRow = $sheet.Cells.Item($sheet.Rows.Count, 1).End($xlUp).Row              # A列の最終行
    $lastCol = $sheet.Cells.Item(1, $sheet.Columns.Count).End($xlToLeft).Column    # 1行目の最終列

    # --- 1行ずつループして読む（A列とB列） ---
    Write-Host "行方向に読み出し"
    for ($r = 1; $r -le $lastRow; $r++) {
        $temp1 = $sheet.Cells.Item($r, 1).Value2
        $temp2 = $sheet.Cells.Item($r, 2).Value2
        Write-Host "$temp1 : $temp2"
    }

    # --- 1列ずつループして読む（1行目と2行目） ---
    Write-Host "列方向に読み出し"
    for ($c = 1; $c -le $lastCol; $c++) {
        $temp1 = $sheet.Cells.Item(1, $c).Value2
        $temp2 = $sheet.Cells.Item(2, $c).Value2
        Write-Host "$temp1 : $temp2"
    }

    # --- 書き込み ---
    $sheet.Cells.Item(1, 1).Value2 = "名前"        # (行, 列) で指定
    $sheet.Range("B1").Value2      = "点数"        # A1形式で指定

    # --- シートの追加・名前変更 ---
    # $sheet は Sheet1、$new は「集計」を指す
    # 同じ名前のシートがすでにあると Name の代入でエラーになる
    # （毎回 sample-org.xlsx からコピーしているので、ここでは起きない）
    $new = $book.Worksheets.Add()
    $new.Name = "集計"

    # --- 数式・書式（Sheet1） ---
    $sheet.Range("C2").Formula = "=SUM(B2:B10)"
    $sheet.Range("A1:C1").Font.Bold = $true
    $sheet.Range("A1:C1").Interior.ColorIndex = 6  # 6 = 黄色
    $sheet.Columns.AutoFit() | Out-Null            # 戻り値を画面に出さない

    # --- 数式（集計シート） ---
    $new.Range("A1").Value2  = "合計"
    $new.Range("B1").Formula = "=SUM(Sheet1!B2:B10)"   # 別シートの参照は「シート名!範囲」

    # --- 保存 ---
    $book.Save()                                   # 上書き保存
    # $book.SaveAs("E:\dev\excel\out.xlsx")        # 別名保存の場合
}
finally {
    # --- 後始末（エラーで止まってもここは必ず実行される） ---
    # Excel は自分への参照が残っている間、Quit() しても終了しない。
    # Excel の部品を受け取った変数（$sheet, $new, $book, $excel）は $null にして参照を外し、
    # [GC]::Collect() で実際に回収する。
    # $v や $t のように値を取り出しただけの変数は、ただのコピーなので何もしなくてよい。
    # 見分け方：$変数.GetType().FullName が System.__ComObject なら参照。
    # それでも残ったときは Get-Process excel | Stop-Process で終了できる。
    if ($book) { $book.Close($false) }             # $false = 保存せずに閉じる（保存は上で済み）
    $excel.Quit()
    $sheet = $null
    $new   = $null
    $book  = $null
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
