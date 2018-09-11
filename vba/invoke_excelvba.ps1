# �Ăяo���}�N���u�b�N
$macroBook = ".\mymacro.xlsm"
$macroBookItem = Get-Item $macroBook

# �}�N���̏������ݐ�i�f�B���N�g�����Ȃ��ꍇ�͍쐬����j
$macroOutputDir = "build"
if (!(Test-Path $macroOutputDir)) {
    Mkdir $macroOutputDir | Out-Null
}
$macroOutputPath = Join-Path `
    (Get-Item $macroOutputDir).FullName `
    "result_$( Get-Date -Format "yyyyMMdd_HHmmss" ).xlsx"

# �����̌�����PowerShell�ɓ`���Ă��炤���߂̃R���\�[���I�u�W�F�N�g
# �i�N���X���`����New���Ă����j
Add-Type -Language "CSharp" "using System;
public class MacroConsole {
    public void WriteLine(string line) {
        Console.WriteLine(`"[macro] `" + line);
    }
}"
$consoleObj = New-Object MacroConsole

$excelApp = $null
try {
    # Excel�̋N��
    $excelApp = New-Object -Com "Excel.Application"
    # ���삪�킩��₷���̂ŁA������悤�ɂ��Ă���
    $excelApp.Visible = $true

    # �}�N���u�b�N���J���ă}�N�������s�B�i�����ꂼ��n���p�X�͐�΃p�X�ɂ��Ă������Ɓj
    $excelApp.Workbooks.Open($macroBookItem.FullName, $false, $true) | Out-Null
    $excelApp.Run("$($macroBookItem.Name)!OutputMacro", $consoleObj, $macroOutputPath)
} finally {
    if ($excelApp) {
        # Excel ���I��
        $excelApp.DisplayAlerts = $false
        $excelApp.Quit()
    }
}

