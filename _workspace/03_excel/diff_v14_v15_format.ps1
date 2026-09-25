<#
v14.xlsm(사용자 수동 편집본)와 v15.xlsm(최신 자동 빌드본)의 시트별 병합 셀·열너비·행높이를
xlsm 내부 OOXML(sheetN.xml)을 직접 파싱해 비교한다. Excel COM을 쓰지 않아 훨씬 빠르고,
셀 값이 아니라 서식(병합/간격)만 비교 대상으로 삼는다.
#>
param(
    [string]$FileA = 'artifacts/excel/교복구매_길라잡이_20260924_v14.xlsm',
    [string]$FileB = 'artifacts/excel/교복구매_길라잡이_20260924_v15.xlsm',
    [string]$OutputPath = '_workspace/03_excel/v14_v15_format_diff.md'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$relNs  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function Get-SheetMap {
    param([string]$ZipPath)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' }
        $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/_rels/workbook.xml.rels' }

        $sr = New-Object System.IO.StreamReader($wbEntry.Open())
        $wbXml = New-Object xml
        $wbXml.LoadXml($sr.ReadToEnd())
        $sr.Close()

        $sr2 = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = New-Object xml
        $relsXml.LoadXml($sr2.ReadToEnd())
        $sr2.Close()

        $relMap = @{}
        foreach ($rel in $relsXml.Relationships.Relationship) {
            $relMap[$rel.Id] = $rel.Target
        }

        $ns = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $ns.AddNamespace('m', $mainNs)

        $sheetMap = [ordered]@{}
        foreach ($sheetNode in $wbXml.SelectNodes('//m:sheets/m:sheet', $ns)) {
            $name = $sheetNode.GetAttribute('name')
            $rid = $sheetNode.GetAttribute('id', $relNs)
            $target = $relMap[$rid]
            $sheetMap[$name] = "xl/$target"
        }
        return $sheetMap
    } finally {
        $zip.Dispose()
    }
}

function Get-SheetFormatInfo {
    param([string]$ZipPath, [string]$EntryName)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $EntryName }
        if (-not $entry) { return $null }
        $sr = New-Object System.IO.StreamReader($entry.Open())
        $xmlText = $sr.ReadToEnd()
        $sr.Close()

        $xml = New-Object xml
        $xml.LoadXml($xmlText)
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('m', $mainNs)

        $merges = New-Object System.Collections.Generic.List[string]
        foreach ($mc in $xml.SelectNodes('//m:mergeCells/m:mergeCell', $ns)) {
            $merges.Add($mc.GetAttribute('ref'))
        }
        $mergesSorted = $merges | Sort-Object

        $cols = New-Object System.Collections.Generic.List[string]
        foreach ($col in $xml.SelectNodes('//m:cols/m:col', $ns)) {
            $min = $col.GetAttribute('min')
            $max = $col.GetAttribute('max')
            $w = $col.GetAttribute('width')
            $cols.Add("${min}-${max}:${w}")
        }
        $colsSorted = $cols | Sort-Object

        $rows = New-Object System.Collections.Generic.List[string]
        foreach ($row in $xml.SelectNodes('//m:sheetData/m:row', $ns)) {
            $r = $row.GetAttribute('r')
            $ht = $row.GetAttribute('ht')
            if ($ht) { $rows.Add("${r}:${ht}") }
        }
        $rowsSorted = $rows | Sort-Object

        return [PSCustomObject]@{
            Merges = $mergesSorted
            Cols   = $colsSorted
            Rows   = $rowsSorted
        }
    } finally {
        $zip.Dispose()
    }
}

$pathA = (Resolve-Path $FileA).Path
$pathB = (Resolve-Path $FileB).Path

Write-Host "v14: $pathA"
Write-Host "v15: $pathB"

$mapA = Get-SheetMap $pathA
$mapB = Get-SheetMap $pathB

$report = New-Object System.Collections.Generic.List[string]
$report.Add('# v14(수동편집) vs v15(최신 자동빌드) 서식 차이 리포트')
$report.Add('')
$report.Add('셀 값이 아니라 병합 범위·열너비·행높이만 비교함. v14에만 있는 항목이 "사용자가 추가한 수정"일 가능성이 높음.')
$report.Add('')

$allNames = @($mapA.Keys) + @($mapB.Keys) | Select-Object -Unique | Sort-Object
$diffCount = 0

foreach ($name in $allNames) {
    if (-not $mapA.Contains($name)) { $report.Add("## $name : v14에 없는 시트(비교 불가)"); $report.Add(''); continue }
    if (-not $mapB.Contains($name)) { $report.Add("## $name : v15에 없는 시트(비교 불가)"); $report.Add(''); continue }

    $infoA = Get-SheetFormatInfo $pathA $mapA[$name]
    $infoB = Get-SheetFormatInfo $pathB $mapB[$name]
    if (-not $infoA -or -not $infoB) { continue }

    $mergeOnlyA = $infoA.Merges | Where-Object { $infoB.Merges -notcontains $_ }
    $mergeOnlyB = $infoB.Merges | Where-Object { $infoA.Merges -notcontains $_ }
    $colOnlyA = $infoA.Cols | Where-Object { $infoB.Cols -notcontains $_ }
    $colOnlyB = $infoB.Cols | Where-Object { $infoA.Cols -notcontains $_ }
    $rowOnlyA = $infoA.Rows | Where-Object { $infoB.Rows -notcontains $_ }
    $rowOnlyB = $infoB.Rows | Where-Object { $infoA.Rows -notcontains $_ }

    if ($mergeOnlyA -or $mergeOnlyB -or $colOnlyA -or $colOnlyB -or $rowOnlyA -or $rowOnlyB) {
        $diffCount++
        $report.Add("## $name")
        if ($mergeOnlyA) { $report.Add('- [v14에만 있는 병합(사용자 추가 가능성)] ' + ($mergeOnlyA -join ', ')) }
        if ($mergeOnlyB) { $report.Add('- [v15에만 있는 병합(v14에서 해제됨)] ' + ($mergeOnlyB -join ', ')) }
        if ($colOnlyA) { $report.Add('- [v14 전용 열너비] ' + ($colOnlyA -join ', ')) }
        if ($colOnlyB) { $report.Add('- [v15 전용 열너비] ' + ($colOnlyB -join ', ')) }
        if ($rowOnlyA) { $report.Add('- [v14 전용 행높이] ' + ($rowOnlyA -join ', ')) }
        if ($rowOnlyB) { $report.Add('- [v15 전용 행높이] ' + ($rowOnlyB -join ', ')) }
        $report.Add('')
    }
}

$report.Insert(2, "차이가 있는 시트 수: $diffCount / $($allNames.Count)")
$report.Insert(3, '')

$report | Out-File -FilePath $OutputPath -Encoding utf8
Write-Host "리포트 생성 완료: $OutputPath (차이 시트 $diffCount 개)"
