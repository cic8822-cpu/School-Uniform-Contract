<#
.SYNOPSIS
  원본 교복 매뉴얼 사본에서 F-007 구간의 허용 공통값 자리만 명시 토큰으로 바꾼다.

.DESCRIPTION
  원본 HWPX는 절대 수정하지 않는다. 복합 매뉴얼의 F-007을 억지로 별도 HWPX로
  재조립하지 않으며, 원본 전체를 복사한 뒤 Contents/section0.xml의 F-007 경계 안에서만
  승인된 예시값을 {{토큰}}으로 치환한다. 표, 이미지, 스타일, 다른 Form ID의 XML은
  그대로 보존한다.

  이 결과는 원본 보존형 토큰 경로의 검증 템플릿이다. 단일 Form ID 출력 분리는 한컴
  수동 또는 별도 안전한 분리 방식이 검증되기 전까지 이 스크립트의 책임 범위가 아니다.
#>
[CmdletBinding()]
param(
    [string]$SourcePath = 'artifacts/hwpx/P0-02_원본사본.hwpx',
    [string]$OutputPath = 'artifacts/hwpx/P2-01/교복매뉴얼_F007_원본보존형_템플릿.hwpx'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$StartAnchor = '[ 6 ] 교복 학교주관구매 구매 요청(예시)'
$EndAnchor = '[ 7 ] 교복 학교주관구매 기초금액 및 계약방법 결정(예시)'
$SectionName = 'Contents/section0.xml'

function Read-ZipText($Entry) {
    $reader = [IO.StreamReader]::new($Entry.Open(), [Text.UTF8Encoding]::new($false, $true), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-ZipText($Zip, [string]$Name, [string]$Text) {
    $old = $Zip.GetEntry($Name)
    if ($null -eq $old) { throw "ZIP 엔트리를 찾을 수 없습니다: $Name" }
    $old.Delete()
    $new = $Zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
    $writer = [IO.StreamWriter]::new($new.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Get-OccurrenceCount([string]$Text, [string]$Value) {
    return [regex]::Matches($Text, [regex]::Escape($Value)).Count
}

function ConvertTo-ParagraphWithoutLineSegArray([string]$Paragraph) {
    # 텍스트가 바뀐 문단의 줄 배치 캐시는 한컴이 다시 계산하도록 제거한다.
    # 문단 이외 XML과 모든 스타일 참조는 바이트 단위로 그대로 보존한다.
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
        if ($paragraphStart -lt 0 -or $paragraphEnd -lt $paragraphStart) {
            throw "토큰화 대상 텍스트의 문단 경계를 찾지 못했습니다: $Text"
        }
        $paragraphEnd += '</hp:p>'.Length
        $paragraph = $Xml.Substring($paragraphStart, $paragraphEnd - $paragraphStart)
        $updatedParagraph = ConvertTo-ParagraphWithoutLineSegArray $paragraph
        $Xml = $Xml.Remove($paragraphStart, $paragraph.Length).Insert($paragraphStart, $updatedParagraph)
        $searchOffset = $paragraphStart + $updatedParagraph.Length
    }
    return $Xml
}

$sourceFullPath = (Resolve-Path -LiteralPath $SourcePath).Path
$outputFullPath = [IO.Path]::GetFullPath($OutputPath)
if ($sourceFullPath -eq $outputFullPath) { throw '원본과 출력 경로는 달라야 합니다.' }
$allowedOutputRoot = [IO.Path]::GetFullPath('artifacts/hwpx/P2-01').TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $outputFullPath.StartsWith($allowedOutputRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "출력은 승인된 산출물 경로 아래여야 합니다: $allowedOutputRoot"
}
if (Test-Path -LiteralPath $outputFullPath -PathType Leaf) {
    $existingOutputPath = (Resolve-Path -LiteralPath $outputFullPath).Path
    if ($existingOutputPath -eq $sourceFullPath) { throw '기존 출력 대상이 원본과 같습니다.' }
}
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputFullPath)) | Out-Null

$sourceHash = (Get-FileHash -LiteralPath $sourceFullPath -Algorithm SHA256).Hash
$temporaryPath = Join-Path ([IO.Path]::GetDirectoryName($outputFullPath)) ('.' + [IO.Path]::GetFileName($outputFullPath) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')

try {
    [IO.File]::Copy($sourceFullPath, $temporaryPath, $true)
    $zip = [IO.Compression.ZipFile]::Open($temporaryPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $section = $zip.GetEntry($SectionName)
        if ($null -eq $section) { throw "원본에 $SectionName 이 없습니다." }
        $xml = Read-ZipText $section
        $start = $xml.IndexOf($StartAnchor, [StringComparison]::Ordinal)
        $end = $xml.IndexOf($EndAnchor, [StringComparison]::Ordinal)
        if ($start -lt 0 -or $end -le $start) { throw 'F-007 경계 표제를 원본 XML에서 찾지 못했습니다.' }

        $prefix = $xml.Substring(0, $start)
        $formXml = $xml.Substring($start, $end - $start)
        $suffix = $xml.Substring($end)
        $replacements = @(
            @{ From = '○○학교-○○(20○○.○.○.)'; To = '{{문서번호}}({{발행일}})'; Expected = 1 },
            @{ From = '○○○○학교-'; To = '{{문서번호}}'; Expected = 1 },
            @{ From = '○ ○ 학 교'; To = '{{학교명}}'; Expected = 1 },
            @{ From = '20○○학년도'; To = '{{학년도}}'; Expected = 2 },
            @{ From = '○○학년도'; To = '{{학년도}}'; Expected = 1 }
        )
        foreach ($replacement in $replacements) {
            $count = Get-OccurrenceCount $formXml $replacement.From
            if ($count -ne $replacement.Expected) {
                throw "F-007 원문 표식 개수가 예상과 다릅니다: $($replacement.From) / 예상 $($replacement.Expected), 실제 $count"
            }
            $formXml = ConvertTo-XmlWithoutLineSegForText $formXml $replacement.From
            $formXml = $formXml.Replace($replacement.From, $replacement.To)
        }
        if ($formXml -notmatch '\{\{학교명\}\}' -or $formXml -notmatch '\{\{학년도\}\}' -or $formXml -notmatch '\{\{문서번호\}\}' -or $formXml -notmatch '\{\{발행일\}\}') {
            throw '필수 F-007 토큰이 모두 삽입되지 않았습니다.'
        }
        $nextXml = $prefix + $formXml + $suffix
        if ($nextXml.Substring(0, $start) -cne $prefix -or $nextXml.Substring($nextXml.Length - $suffix.Length) -cne $suffix) {
            throw 'F-007 경계 밖 XML이 변경되었습니다.'
        }
        Write-ZipText -Zip $zip -Name $SectionName -Text $nextXml
    } finally {
        if ($zip) { $zip.Dispose() }
    }

    if ((Get-FileHash -LiteralPath $sourceFullPath -Algorithm SHA256).Hash -ne $sourceHash) {
        throw '원본 HWPX 해시가 변경되었습니다.'
    }
    Move-Item -LiteralPath $temporaryPath -Destination $outputFullPath -Force
    Write-Output "PASS: 원본 전체 보존 사본의 F-007 허용 공통값만 토큰화 -> $outputFullPath"
} finally {
    if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
    }
}
