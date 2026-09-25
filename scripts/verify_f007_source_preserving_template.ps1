<#
.SYNOPSIS
  F-007 원본 보존형 토큰 템플릿과 골든 결과를 구조·토큰·원본보호 기준으로 검증한다.
#>
[CmdletBinding()]
param(
    [string]$SourcePath = 'artifacts/hwpx/P0-02_원본사본.hwpx',
    [string]$TemplatePath = 'artifacts/hwpx/P2-01/교복매뉴얼_F007_원본보존형_템플릿.hwpx',
    [string]$OutputPath = 'artifacts/hwpx/P2-01/교복매뉴얼_F007_원본보존형_골든결과.hwpx',
    [string]$WorkDirectory = '_workspace/02_hwpx/P2-01/원본보존형_토큰검증'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

function Read-ZipText($Entry) {
    $reader = [IO.StreamReader]::new($Entry.Open(), [Text.UTF8Encoding]::new($false, $true), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Get-ZipEntryHashMap([string]$Path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $map = @{}
        foreach ($entry in $zip.Entries) {
            $sha256 = [Security.Cryptography.SHA256]::Create()
            $stream = $entry.Open()
            try { $map[$entry.FullName] = ([BitConverter]::ToString($sha256.ComputeHash($stream))).Replace('-', '') }
            finally { $stream.Dispose(); $sha256.Dispose() }
        }
        return $map
    } finally { $zip.Dispose() }
}
function Get-SectionXml([string]$Path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('Contents/section0.xml')
        if ($null -eq $entry) { throw 'Contents/section0.xml이 없습니다.' }
        return Read-ZipText $entry
    } finally { $zip.Dispose() }
}
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-OnlySection0Changed([hashtable]$Expected, [hashtable]$Actual, [string]$Description) {
    Assert-True ((($Expected.Keys | Sort-Object) -join "`n") -eq (($Actual.Keys | Sort-Object) -join "`n")) "$Description ZIP 엔트리 목록이 다릅니다."
    foreach ($name in $Expected.Keys) {
        if ($name -ne 'Contents/section0.xml') {
            Assert-True ($Expected[$name] -eq $Actual[$name]) "$Description 비-본문 ZIP 엔트리가 변경되었습니다: $name"
        }
    }
}
function ConvertTo-ParagraphWithoutLineSegArray([string]$Paragraph) {
    return [regex]::Replace(
        $Paragraph,
        '<hp:linesegarray\b[^>]*>.*?</hp:linesegarray>|<hp:linesegarray\b[^>]*/>',
        '',
        [Text.RegularExpressions.RegexOptions]::Singleline
    )
}
function ConvertTo-XmlWithoutLineSegForText([string]$Xml, [string]$Text) {
    $searchOffset = 0
    while ($true) {
        $textIndex = $Xml.IndexOf($Text, $searchOffset, [StringComparison]::Ordinal)
        if ($textIndex -lt 0) { break }
        $paragraphStart = $Xml.LastIndexOf('<hp:p', $textIndex, [StringComparison]::Ordinal)
        $paragraphEnd = $Xml.IndexOf('</hp:p>', $textIndex, [StringComparison]::Ordinal)
        if ($paragraphStart -lt 0 -or $paragraphEnd -lt $paragraphStart) { throw "토큰화 대상 텍스트의 문단 경계를 찾지 못했습니다: $Text" }
        $paragraphEnd += '</hp:p>'.Length
        $paragraph = $Xml.Substring($paragraphStart, $paragraphEnd - $paragraphStart)
        $updatedParagraph = ConvertTo-ParagraphWithoutLineSegArray $paragraph
        $Xml = $Xml.Remove($paragraphStart, $paragraph.Length).Insert($paragraphStart, $updatedParagraph)
        $searchOffset = $paragraphStart + $updatedParagraph.Length
    }
    return $Xml
}
function Convert-F007ToTemplateXml([string]$SourceXml) {
    $startAnchor = '[ 6 ] 교복 학교주관구매 구매 요청(예시)'
    $endAnchor = '[ 7 ] 교복 학교주관구매 기초금액 및 계약방법 결정(예시)'
    $start = $SourceXml.IndexOf($startAnchor, [StringComparison]::Ordinal)
    $end = $SourceXml.IndexOf($endAnchor, [StringComparison]::Ordinal)
    if ($start -lt 0 -or $end -le $start) { throw '원본 F-007 경계를 찾지 못했습니다.' }
    $prefix = $SourceXml.Substring(0, $start)
    $formXml = $SourceXml.Substring($start, $end - $start)
    $suffix = $SourceXml.Substring($end)
    foreach ($replacement in @(
        @{ From = '○○학교-○○(20○○.○.○.)'; To = '{{문서번호}}({{발행일}})'; Expected = 1 },
        @{ From = '○○○○학교-'; To = '{{문서번호}}'; Expected = 1 },
        @{ From = '○ ○ 학 교'; To = '{{학교명}}'; Expected = 1 },
        @{ From = '20○○학년도'; To = '{{학년도}}'; Expected = 2 },
        @{ From = '○○학년도'; To = '{{학년도}}'; Expected = 1 }
    )) {
        $count = [regex]::Matches($formXml, [regex]::Escape($replacement.From)).Count
        if ($count -ne $replacement.Expected) { throw "F-007 원문 표식 개수가 예상과 다릅니다: $($replacement.From)" }
        $formXml = ConvertTo-XmlWithoutLineSegForText $formXml $replacement.From
        $formXml = $formXml.Replace($replacement.From, $replacement.To)
    }
    return $prefix + $formXml + $suffix
}
function Get-F007ParagraphTextLength([string]$Xml) {
    $start = $Xml.IndexOf('[ 6 ] 교복 학교주관구매 구매 요청(예시)', [StringComparison]::Ordinal)
    $end = $Xml.IndexOf('[ 7 ] 교복 학교주관구매 기초금액 및 계약방법 결정(예시)', [StringComparison]::Ordinal)
    if ($start -lt 0 -or $end -le $start) { throw 'F-007 페이지 가드용 경계를 찾지 못했습니다.' }
    $lengths = @()
    foreach ($paragraph in [regex]::Matches($Xml.Substring($start, $end - $start), '<hp:p\b.*?</hp:p>', [Text.RegularExpressions.RegexOptions]::Singleline)) {
        $text = ''
        foreach ($node in [regex]::Matches($paragraph.Value, '<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>', [Text.RegularExpressions.RegexOptions]::Singleline)) {
            $text += $node.Groups['text'].Value
        }
        $lengths += ([regex]::Replace($text, '\s', '')).Length
    }
    return $lengths
}
function Assert-F007LayoutGuard([string]$ReferenceXml, [string]$OutputXml) {
    # 실습 자료의 page_guard 원칙을 Form ID 범위에 적용한다. 실제 한컴 PDF 비교는 P2-02에서 별도로 수행한다.
    $referenceLengths = @(Get-F007ParagraphTextLength $ReferenceXml)
    $outputLengths = @(Get-F007ParagraphTextLength $OutputXml)
    Assert-True ($referenceLengths.Count -eq $outputLengths.Count) 'F-007 페이지 가드: 문단 수가 달라졌습니다.'
    for ($index = 0; $index -lt $referenceLengths.Count; $index++) {
        $before = $referenceLengths[$index]; $after = $outputLengths[$index]
        if ($before -eq 0 -and $after -eq 0) { continue }
        $ratio = if ($before -eq 0) { 1.0 } else { [Math]::Abs($after - $before) / $before }
        Assert-True ($ratio -le 0.50) "F-007 페이지 가드: $($index + 1)번째 문단 길이 변화가 50%를 초과했습니다 ($before -> $after)."
    }
}

$sourceFullPath = (Resolve-Path -LiteralPath $SourcePath).Path
$templateFullPath = (Resolve-Path -LiteralPath $TemplatePath).Path
$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
[IO.Directory]::CreateDirectory($WorkDirectory) | Out-Null
$jsonPath = Join-Path ([IO.Path]::GetFullPath($WorkDirectory)) 'F007_원본보존형_골든.json'
$sourceHash = (Get-FileHash -LiteralPath $sourceFullPath -Algorithm SHA256).Hash
$templateHash = (Get-FileHash -LiteralPath $templateFullPath -Algorithm SHA256).Hash

$sourceEntries = Get-ZipEntryHashMap $sourceFullPath
$templateEntries = Get-ZipEntryHashMap $templateFullPath
Assert-OnlySection0Changed -Expected $sourceEntries -Actual $templateEntries -Description '원본과 템플릿'

$sourceXml = Get-SectionXml $sourceFullPath
$templateXml = Get-SectionXml $templateFullPath
$expectedTemplateXml = Convert-F007ToTemplateXml $sourceXml
Assert-True ($templateXml -ceq $expectedTemplateXml) '템플릿 section0.xml이 허용된 F-007 토큰 치환 결과와 다릅니다.'

$payload = [ordered]@{ 학교명 = '검증초등학교'; 학년도 = '2026학년도'; 문서번호 = '검증초-2026-1'; 발행일 = '2026. 3. 2.' } | ConvertTo-Json -Compress
[IO.File]::WriteAllText($jsonPath, $payload, [Text.UTF8Encoding]::new($false))
& (Join-Path $PSScriptRoot 'hwpx_token_fill.ps1') -TemplatePath $templateFullPath -JsonPath $jsonPath -OutputPath $outputFullPath
if (-not $?) { throw '원본 보존형 토큰 치환이 실패했습니다.' }

$outputXml = Get-SectionXml $outputFullPath
$expectedOutputXml = $expectedTemplateXml.Replace('{{학교명}}', '검증초등학교').Replace('{{학년도}}', '2026학년도').Replace('{{문서번호}}', '검증초-2026-1').Replace('{{발행일}}', '2026. 3. 2.')
Assert-True ($outputXml -ceq $expectedOutputXml) '골든 결과 section0.xml이 허용된 토큰 치환 결과와 다릅니다.'
Assert-F007LayoutGuard $sourceXml $outputXml
$outputEntries = Get-ZipEntryHashMap $outputFullPath
Assert-OnlySection0Changed -Expected $templateEntries -Actual $outputEntries -Description '템플릿과 골든 결과'

Assert-True ((Get-FileHash -LiteralPath $sourceFullPath -Algorithm SHA256).Hash -eq $sourceHash) '검증 중 원본 HWPX가 변경되었습니다.'
Assert-True ((Get-FileHash -LiteralPath $templateFullPath -Algorithm SHA256).Hash -eq $templateHash) '검증 중 템플릿 HWPX가 변경되었습니다.'

& npx -y kordoc@4.13.1 validate $outputFullPath
if ($LASTEXITCODE -ne 0) { throw "kordoc validate 실패(exit $LASTEXITCODE)" }
Write-Output "PASS: F-007 원본 보존형 템플릿/골든결과 ZIP·XML·토큰·경계·원본해시·validate 검증 완료 -> $outputFullPath"
