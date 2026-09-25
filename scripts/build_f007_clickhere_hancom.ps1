# Convert F-007 token placeholders to Hancom click-here fields.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$ValuesJson,
    [Parameter(Mandatory = $true)][string]$Output,
    [Parameter(Mandatory = $true)][string]$LogPath
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) { throw "Input file not found: $Source" }
if (Test-Path -LiteralPath $Output) { throw "Output already exists: $Output" }

$values = Get-Content -Raw -Encoding utf8 -LiteralPath $ValuesJson | ConvertFrom-Json
$names = @($values.psobject.Properties.Name)
if ($names.Count -ne 21) { throw "Unexpected F-007 field count: $($names.Count)" }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Output) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LogPath) | Out-Null
Copy-Item -LiteralPath $Source -Destination $Output
"START $(Get-Date -Format o)" | Set-Content -Encoding utf8 -LiteralPath $LogPath

$hwp = New-Object -ComObject HWPFrame.HwpObject
try {
    $null = $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
    if (-not $hwp.Open((Resolve-Path -LiteralPath $Output).Path, '', 'forceopen:true')) { throw 'Hancom could not open the output copy' }

    foreach ($name in $names) {
        $token = ([char]123).ToString() + ([char]123).ToString() + $name + ([char]125).ToString() + ([char]125).ToString()
        $pset = $hwp.HParameterSet.HFindReplace
        $null = $hwp.HAction.GetDefault('RepeatFind', $pset.HSet)
        $pset.FindString = $token
        $pset.Direction = $hwp.FindDir('Forward')
        $pset.FindType = 1
        $pset.MatchCase = $false
        $pset.WholeWordOnly = $false
        $pset.UseWildCards = $false
        if (-not $hwp.HAction.Execute('RepeatFind', $pset.HSet)) { throw "Token not found: $token" }
        if (-not $hwp.CreateField('', 'F-007 automated fill field', $name)) { throw "Click-here field creation failed: $name" }
        "CREATED $name $(Get-Date -Format o)" | Add-Content -Encoding utf8 -LiteralPath $LogPath
    }

    if (-not $hwp.SaveAs((Resolve-Path -LiteralPath $Output).Path, 'HWPX', '')) { throw 'HWPX save failed' }
    "SAVED $(Get-Date -Format o)" | Add-Content -Encoding utf8 -LiteralPath $LogPath
}
catch {
    "FAILED $($_.Exception.Message)" | Add-Content -Encoding utf8 -LiteralPath $LogPath
    throw
}
finally {
    $hwp.Quit()
}
