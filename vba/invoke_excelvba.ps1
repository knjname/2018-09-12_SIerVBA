# 呼び出すマクロブック
$macroBook = ".\mymacro.xlsm"
$macroBookItem = Get-Item $macroBook

# マクロの書き込み先（ディレクトリがない場合は作成する）
$macroOutputDir = "build"
if (!(Test-Path $macroOutputDir)) {
    Mkdir $macroOutputDir | Out-Null
}
$macroOutputPath = Join-Path `
    (Get-Item $macroOutputDir).FullName `
    "result_$( Get-Date -Format "yyyyMMdd_HHmmss" ).xlsx"

# 処理の現況をPowerShellに伝えてもらうためのコンソールオブジェクト
# （クラスを定義してNewしておく）
Add-Type -Language "CSharp" "using System;
public class MacroConsole {
    public void WriteLine(string line) {
        Console.WriteLine(`"[macro] `" + line);
    }
}"
$consoleObj = New-Object MacroConsole

$excelApp = $null
try {
    # Excelの起動
    $excelApp = New-Object -Com "Excel.Application"
    # 動作がわかりやすいので、見えるようにしておく
    $excelApp.Visible = $true

    # マクロブックを開いてマクロを実行。（※それぞれ渡すパスは絶対パスにしておくこと）
    $excelApp.Workbooks.Open($macroBookItem.FullName, $false, $true) | Out-Null
    $excelApp.Run("$($macroBookItem.Name)!OutputMacro", $consoleObj, $macroOutputPath)
} finally {
    if ($excelApp) {
        # Excel を終了
        $excelApp.DisplayAlerts = $false
        $excelApp.Quit()
    }
}

