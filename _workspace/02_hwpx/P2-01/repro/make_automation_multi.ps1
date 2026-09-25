$ErrorActionPreference = "Stop"
$hwp = New-Object -ComObject HwpFrame.HwpObject
$hwp.XHwpWindows.Item(0).Visible = $false
try {
    $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
} catch {}
try {
    $hwp.SetMessageBoxMode(0x1FFFF1)
} catch {}
$hwp.HAction.GetDefault("FileNew", $hwp.HParameterSet.HFileOpenSave.HSet)
$hwp.HAction.Execute("FileNew", $hwp.HParameterSet.HFileOpenSave.HSet)

$fieldNames = @("F007SCHOOLNAME", "F007VIA", "F007TITLE", "F007REFERENCEDOCUMENT", "F007DRAFTER")
foreach ($name in $fieldNames) {
    $hwp.HAction.GetDefault("InsertText", $hwp.HParameterSet.HInsertText.HSet)
    $hwp.HParameterSet.HInsertText.Text = ("Label " + $name + ": ")
    $hwp.HAction.Execute("InsertText", $hwp.HParameterSet.HInsertText.HSet)

    $created = $hwp.CreateField("", "", $name)
    Write-Output ("CreateField " + $name + " -> " + $created)

    $hwp.HAction.GetDefault("BreakPara", $hwp.HParameterSet.HInsertText.HSet)
    $hwp.HAction.Run("BreakPara")
}

$outPath = Join-Path $PSScriptRoot "automation_multi.hwpx"
$hwp.SaveAs($outPath, "HWPX", "")
$hwp.Clear(1)
$hwp.Quit()
Write-Output ("SAVED: " + $outPath)

