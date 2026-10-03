# ============================================================
# step8 : 人に渡して動かす（起動用の bat と実行環境）
# ------------------------------------------------------------
# 学ぶこと
#   - bat は入口だけにして、処理は ps に書く（ps から ps を呼ぶ）
#   - PowerShell 5.1（powershell）と 7（pwsh）の違いと、使うバージョンの決め方
#   - 実行ポリシーと -ExecutionPolicy Bypass
#
# 使うファイル
#   step8.bat         : 入口（ダブルクリックで実行する）。引数をそのまま step8.ps1 に渡す
#   step6-sub.ps1     : 呼ばれる側（指定セルに文字を書き、結果を1行返す）
#   sample-step6.xlsx : step6-sub.ps1 が毎回 sample-org.xlsx からコピーして作る作業用
#
# 実行方法
#   step8.bat をダブルクリック              引数なし
#   step8.bat G5 "from bat"                  引数あり（コマンドプロンプトから）
#   .\step8.ps1 -Addr G5 -Text "ps から"     PowerShell 7 から直接
#
# 覚えておくこと
#   - 起動するコマンドでバージョンが決まる：pwsh = 7、powershell = 5.1
#   - & "パス" で呼んだスクリプトは、呼んだ側と同じバージョンで動く
#   - #Requires -Version 7 を書くと、5.1 で実行したときに1行目も動かずに止まる
#   - 5.1 は BOM なしの .ps1 を Windows の文字コード（日本語版は Shift-JIS）で読むので、
#     日本語が化ける。7 は UTF-8 で読む
#   - 値は ps どうしでやり取りし、bat には終了コードだけ返す（for /f が要らない）
# ============================================================
#Requires -Version 7
param(
    [string]$Addr = "G1",
    [string]$Text = "step8 から"
)

# --- エラーが起きたら、その場でスクリプトを止める（終了コードは 1 になる） ---
$ErrorActionPreference = "Stop"

# --- どの PowerShell で、どの設定で動いているかを表示する ---
Write-Host "PowerShell   = $($PSVersionTable.PSVersion)（$($PSVersionTable.PSEdition)）"   # 7 なら Core、5.1 なら Desktop
Write-Host "場所         = $PSHOME"
Write-Host "実行ポリシー = $(Get-ExecutionPolicy)"     # bat から Bypass で起動すると Bypass
Write-Host "引数         = Addr: $Addr / Text: $Text"

# --- 処理の本体は別のスクリプトに任せ、戻り値を受け取る（step6.ps1 と同じ ps→ps） ---
$result = & "$PSScriptRoot\step6-sub.ps1" $Addr $Text
Write-Host "戻り値       = $result"
