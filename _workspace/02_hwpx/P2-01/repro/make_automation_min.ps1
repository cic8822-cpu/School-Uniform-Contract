$ErrorActionPreference = "Stop"
$hwp = New-Object -ComObject HwpFrame.HwpObject
$hwp.XHwpWindows.Item(0).Visible = $false
try {
    $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
} catch {}
try {
    # Suppress all interactive message boxes (format conversion prompts etc.)
    $hwp.SetMessageBoxMode(0x1FFFF1)
} catch {}
$hwp.HAction.GetDefault("FileNew", $hwp.HParameterSet.HFileOpenSave.HSet)
$hwp.HAction.Execute("FileNew", $hwp.HParameterSet.HFileOpenSave.HSet)

$hwp.HAction.GetDefault("InsertText", $hwp.HParameterSet.HInsertText.HSet)
$hwp.HParameterSet.HInsertText.Text = "TEST VIA: "
$hwp.HAction.Execute("InsertText", $hwp.HParameterSet.HInsertText.HSet)

$created = $hwp.CreateField("", "", "TESTVIA")
Write-Output ("CreateField result: " + $created)

$outPath = Join-Path $PSScriptRoot "automation_min.hwpx"
$hwp.SaveAs($outPath, "HWPX", "")
$hwp.Clear(1)
$hwp.Quit()
Write-Output ("SAVED: " + $outPath)




