# ============================================================
# step7 : エラーが起きたときの扱い方
# ------------------------------------------------------------
# 学ぶこと
#   1. PowerShell 側のエラーを try / catch で受ける
#   2. エラーにならない失敗（$null が返る）を自分で確かめる
#   3. 計算できない値（負数のルート、定義域・値域の外など）の扱い
#        - マクロ側で On Error を書き、結果かエラーの文字列を返す
#        - PowerShell の [math] はエラーにならず NaN / Infinity を返す
#   4. マクロが止まったまま返ってこないとき、時間切れで Excel を終了させる
#   5. Excel が異常終了（メモリアクセス違反など）したときの受け方
#   6. 呼び出した別のスクリプトのエラーを受ける
#   7. bat に成功 / 失敗を伝える（step7.bat）
#
# 使うファイル
#   macros.xlsm   : マクロを書いたブック（SafeCalc / UnsafeSqrt を使う）
#   step6-sub.ps1 : 6 で、わざと失敗させて呼ぶ
#   step7.bat     : このスクリプトを bat から呼び、終了コードで分岐する
#
# 実行方法
#   .\step7.ps1    （4 で 10 秒ほど待つ）
#
# 覚えておくこと
#   - PowerShell は、エラーが起きてもメッセージを出して次の行へ進むことが多い。
#     最後まで進むと終了コードは 0（成功）になり、呼び出し元は失敗に気づけない。
#     先頭に $ErrorActionPreference = "Stop" を書くと、エラーの行で止まるようになる
#   - マクロの中で起きたエラーは、PowerShell の try / catch では受けられない。
#     Excel がエラーのダイアログを出して待ち続ける（画面非表示なので見えない）。
#     マクロ側で On Error を書いて、戻り値でエラーを知らせる
#   - On Error を書けないマクロを呼ぶときは、時間切れで Excel を終了させるしかない
#   - Excel の異常終了は、マクロの On Error でも PowerShell でも防げない。
#     PowerShell からは「Excel が急にいなくなった」というエラーとして見える
#       0x800706BA : RPC サーバーを利用できません（もういない）
#       0x800706BE : リモート プロシージャ コールに失敗しました（呼び出し中にいなくなった）
#     Windows では、異常終了時の記録（コアダンプにあたるもの）はクラッシュダンプと呼ぶ
#   - Excel がいなくなった後は Quit() も失敗する。後始末にも try / catch を付ける
#   - 異常終了の後は、保存していない変更は残らない。最初からやり直す
# ============================================================

$macroPath = "E:\dev\excel\macros.xlsm"

# --- エラーが起きたら、その場でスクリプトを止める ---
# 止まったエラーは try / catch で受けられる。受けなければスクリプトが終了コード 1 で終わる
$ErrorActionPreference = "Stop"

# --- Excel のプロセス ID を調べる関数 ---
# 自分が起動した Excel だけを終了させるために使う
# （Get-Process excel | Stop-Process だと、ほかの Excel も巻き込む）
Add-Type -Namespace Win32 -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);'
function Get-ExcelPid($excel) {
    [uint32]$id = 0
    [void][Win32.User32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$id)
    return $id
}

# --- Excel を終了する関数（Excel がすでにいなくても止まらない） ---
# 呼んだ後、呼び出し側で Excel の部品を受け取った変数を $null にして [GC]::Collect() する
function Stop-ExcelSafely($excel) {
    try { $excel.Quit() } catch { }                # Excel がいないと Quit() も失敗する
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
}

# ============================================================
# 1〜3 : ふつうのエラー（Excel は動き続ける）
# ============================================================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない（デバッグ時は $true）
$excel.DisplayAlerts = $false    # 上書き確認などのダイアログを出さない

try {
    $macroBook = $excel.Workbooks.Open($macroPath, 0, $true)    # 第3引数 $true = 読み取り専用

    # --- 1. PowerShell 側のエラーを try / catch で受ける ---
    # catch の中では $_ がエラー。$_.Exception.Message がメッセージ
    Write-Host "--- 1. try / catch"
    try {
        $book = $excel.Workbooks.Open("E:\dev\excel\nothing.xlsx")     # 存在しないファイル
    }
    catch {
        Write-Host "開けませんでした : $($_.Exception.Message)"
    }

    # --- 2. エラーにならない失敗を確かめる ---
    # プロパティ名を間違えて読むと、エラーにならず $null が返る（step1.ps1 を参照）
    Write-Host "--- 2. `$null の確認"
    $sheets = $macroBook.WorkSheet                 # 正しくは Worksheets
    if ($null -eq $sheets) {
        Write-Host "シートを取得できませんでした（プロパティ名の間違い）"
    }

    # --- 3. 計算できない値：マクロ側でエラーを処理して、結果を戻り値で受け取る ---
    # SafeCalc は、計算できれば数値を、できなければ "ERROR: 番号 説明" の文字列を返す
    Write-Host "--- 3. マクロの計算エラー"
    $cases = @(
        @("sqrt", 9),       # 計算できる
        @("sqrt", -1),      # ルートの中が負
        @("log",  0),       # 対数の定義域の外
        @("acos", 2),       # 逆三角関数の定義域の外（-1〜1 のみ）
        @("exp",  1000),    # 結果が大きすぎる（値域の外、オーバーフロー）
        @("div",  0)        # 0 で割る
    )
    foreach ($case in $cases) {
        $kind, $x = $case
        $result = $excel.Run("'macros.xlsm'!SafeCalc", $kind, $x)
        if ($result -is [string]) {
            Write-Host "$kind($x) : 計算できません（$result）"
        }
        else {
            Write-Host "$kind($x) = $result"
        }
    }

    # --- 3. 計算できない値：PowerShell で計算する場合 ---
    # [math] の関数はエラーにならず、NaN（数ではない）や Infinity（無限大）を返す。
    # そのまま次の計算に使うと結果もおかしくなるので、自分で確かめる
    Write-Host "--- 3. PowerShell の計算エラー"
    $values = [ordered]@{
        "Sqrt(-1)"  = [math]::Sqrt(-1)             # NaN
        "Acos(2)"   = [math]::Acos(2)              # NaN
        "Exp(1000)" = [math]::Exp(1000)            # Infinity
        "Log(0)"    = [math]::Log(0)               # -Infinity
    }
    foreach ($name in $values.Keys) {
        $v = $values[$name]
        if ([double]::IsNaN($v) -or [double]::IsInfinity($v)) {
            Write-Host "$name : 計算できません（$v）"
        }
    }
    # 整数の 0 割りだけは、エラーになる
    try { $v = 1 / 0 } catch { Write-Host "1 / 0 : $($_.Exception.Message)" }
}
finally {
    # --- 後始末（考え方は step1.ps1 の finally を参照） ---
    if ($macroBook) { $macroBook.Close($false) }
    Stop-ExcelSafely $excel
    $sheets    = $null
    $book      = $null
    $macroBook = $null
    $excel     = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================
# 4 : マクロが返ってこない（エラーのダイアログを出して待ち続ける）
# ============================================================
# UnsafeSqrt には On Error がないので、負の数を渡すとダイアログが出て止まる。
# 見張り役を別に動かしておき、時間切れになったら Excel を終了させる
Write-Host "--- 4. 時間切れ（10 秒待つ）"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $macroBook = $excel.Workbooks.Open($macroPath, 0, $true)
    $excelPid  = Get-ExcelPid $excel

    # 見張り役：10 秒たったら、この Excel を強制終了する（別のスレッドで動く）
    $watchdog = Start-ThreadJob -ArgumentList $excelPid, 10 -ScriptBlock {
        param($id, $seconds)
        Start-Sleep -Seconds $seconds
        Stop-Process -Id $id -Force
    }

    try {
        $result = $excel.Run("'macros.xlsm'!UnsafeSqrt", -1)    # ここで止まる
        Write-Host "UnsafeSqrt(-1) = $result"
    }
    catch {
        # Excel が強制終了されると、待っていた Run() がエラーで戻ってくる
        if ($watchdog.State -eq "Completed") {
            Write-Host "時間切れで Excel を終了しました : $($_.Exception.Message)"
        }
        else {
            Write-Host "エラー : $($_.Exception.Message)"
        }
    }
    finally {
        # 時間内に終わったときは、見張り役を止める
        Stop-Job $watchdog
        Remove-Job $watchdog -Force
    }
}
finally {
    Stop-ExcelSafely $excel                        # Excel はもういないが、止まらずに進む
    $macroBook = $null
    $excel     = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================
# 5 : Excel の異常終了（メモリアクセス違反などで Excel が落ちる）
# ============================================================
# 本当に落とす代わりに、自分で起動した Excel を強制終了して同じ状況を作る
Write-Host "--- 5. Excel の異常終了"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $macroBook = $excel.Workbooks.Open($macroPath, 0, $true)
    $excelPid  = Get-ExcelPid $excel

    Stop-Process -Id $excelPid -Force              # 異常終了のかわり
    Start-Sleep -Seconds 1

    try {
        $sum = $excel.Run("'macros.xlsm'!AddTwo", 3, 4)
        Write-Host "AddTwo = $sum"
    }
    catch {
        # Excel がまだ動いているかどうかで、原因を見分ける
        if (Get-Process -Id $excelPid -ErrorAction SilentlyContinue) {
            Write-Host "エラー（Excel は動いている） : $($_.Exception.Message)"
        }
        else {
            Write-Host "Excel が異常終了しました : $($_.Exception.Message)"
        }
    }
}
finally {
    Stop-ExcelSafely $excel
    $macroBook = $null
    $excel     = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================
# 6 : 呼び出した別のスクリプトのエラーを受ける
# ============================================================
# 存在しないセル番地を渡して、step6-sub.ps1 をわざと失敗させる。
# 呼んだ先で止まったエラーは、こちらの catch に届く
# （step6-sub.ps1 の finally は先に実行されるので、Excel は残らない）
Write-Host "--- 6. 別のスクリプトのエラー"
try {
    $result = & "$PSScriptRoot\step6-sub.ps1" "???" "失敗する"
    Write-Host "戻り値 = $result"
}
catch {
    Write-Host "step6-sub.ps1 が失敗しました : $($_.Exception.Message)"
}

# ============================================================
# 7 : bat に成功 / 失敗を伝える
# ============================================================
# exit の数値が bat の %ERRORLEVEL% になる（0 = 成功、0 以外 = 失敗）。
# catch で受けなかったエラーでスクリプトが止まると、終了コードは 1 になる
# （先頭の $ErrorActionPreference = "Stop" が必要。step7.bat の2つ目の例を参照）。
# ここまでのエラーはすべて catch で受けたので、成功として終わる
exit 0
