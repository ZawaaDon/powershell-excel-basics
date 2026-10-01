# ============================================================
# excel-lib : Excel 操作の共通関数
# ------------------------------------------------------------
# 使い方
#   呼び出す側のスクリプトの先頭で読み込む（step3.ps1 を参照）
#     . "$PSScriptRoot\excel-lib.ps1"
#
# 覚えておくこと
#   - このファイル単体で実行しても何も起きない（関数を定義するだけ）
#   - シートは引数で受け取る。関数の中で ActiveSheet は使わない
#   - 引数で受け取った $sheet は呼び出し側のもの。ここで $null にする必要はない
# ============================================================

# 指定シートの1行目を太字にする
#   例：Set-BoldHeader $sheet
function Set-BoldHeader($sheet) {
    $sheet.Rows.Item(1).Font.Bold = $true
}

# 指定セルに文字を書く
#   例：Write-CellText $sheet "D1" "こんにちは"
function Write-CellText($sheet, $addr, $text) {
    $sheet.Range($addr).Value2 = $text
}

# 2つの数を足して返す
#   例：$sum = Add-Two 3 4
function Add-Two($a, $b) {
    return $a + $b
}
