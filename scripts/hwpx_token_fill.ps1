<#
.SYNOPSIS
  HWPX 사본의 명시적 공통 토큰({{학교명}} 등)을 JSON 값으로 채운다.

.DESCRIPTION
  이 스크립트는 원본을 절대 수정하지 않는다. XML을 다시 직렬화하지 않고
  Contents/section*.xml의 hp:t 텍스트 영역만 바이트 문자열 수준에서 바꾼다.
  토큰이 인접한 여러 hp:t 런에 나뉘어 있어도 같은 문단 안에서 찾아 채운다.
  서명·직인·연락처 등 개인정보/민감 토큰 및 허용 목록 밖 키는 거부한다.
  단, 결재권자(교장 등)의 직무상 성명은 공문서 관행상 통상 기재되는 정보이므로
  허용 목록에 포함한다(2026-09-19 사용자 결정, `입력데이터_사전.md` C-07 참조).
  위원·업체 대표자 등 그 외 개인 성명은 계속 거부한다.

  '관련문서'(D-03)·'제출기한'(B-07)은 47종 P2-04 템플릿을 실제로 스캔한 결과
  존재가 확인된 비개인정보 업무 필드라 2026-09-23 동일 원칙으로 허용 목록에
  추가함(③ Excel→HWPX 연동 스크립트, `scripts/export_hwpx_p204_from_excel.ps1`).

  JSON 계약(반드시 UTF-8 BOM 없음):
    { "학교명": "검증초등학교", "학년도": "2026학년도",
      "문서번호": "검증초-1", "발행일": "2026. 3. 2.", "교장명": "홍길동",
      "관련문서": "검증초-2026-1(2026. 3. 2.)", "제출기한": "2026. 3. 10." }

  Excel VBA에서 직접 호출하지 않는다. 호출 주체는 이 CLI의 종료 코드와 표준 오류를
  확인하고, 생성 파일만 사용해야 한다.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TemplatePath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$JsonPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

# 공통·비식별 값 + 결재권자 직무상 성명(교장명)만 허용한다. 새 항목은 개인정보 경계 검토 뒤 이 목록에 추가한다.
$AllowedTokens = @('학교명', '학년도', '문서번호', '발행일', '제목', '수신기관', '시행일', '공개구분', '교장명', '관련문서', '제출기한')
$ForbiddenTokenPattern = '성명|이름|담당자|전화|휴대|주소|이메일|서명|직인|계좌|주민|생년|학번|설문'
$TokenPattern = '\{\{([^{}]+)\}\}'

function Get-Utf8NoBomText([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "JSON은 UTF-8 BOM 없이 저장해야 합니다: $Path"
    }
    return [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
}

function ConvertTo-XmlSafeText([string]$Value) {
    if ($Value -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') { throw 'JSON 값에 XML에서 허용되지 않는 제어 문자가 있습니다.' }
    return [Security.SecurityElement]::Escape($Value)
}

function Get-JsonMap([string]$Path) {
    $raw = Get-Utf8NoBomText $Path
    try { $obj = $raw | ConvertFrom-Json } catch { throw "JSON 파싱 실패: $($_.Exception.Message)" }
    $properties = @($obj.PSObject.Properties)
    if ($null -eq $obj -or $properties.Count -eq 0) { throw 'JSON 객체에는 하나 이상의 허용 토큰 값이 있어야 합니다.' }
    $map = @{}
    foreach ($property in $properties) {
        $key = $property.Name
        if ($key -notin $AllowedTokens -or $key -match $ForbiddenTokenPattern) { throw "허용되지 않거나 민감한 토큰 키입니다: $key" }
        if ($null -eq $property.Value) { throw "토큰 값은 null일 수 없습니다: $key" }
        if ($property.Value -isnot [string]) { throw "토큰 값은 문자열이어야 합니다: $key" }
        $map[$key] = ConvertTo-XmlSafeText $property.Value
    }
    return $map
}

function Assert-Xml([string]$Xml, [string]$EntryName) {
    try {
        $settings = [Xml.XmlReaderSettings]::new(); $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $reader = [Xml.XmlReader]::Create([IO.StringReader]::new($Xml), $settings)
        while ($reader.Read()) { }
        $reader.Dispose()
    } catch { throw "XML 유효성 실패($EntryName): $($_.Exception.Message)" }
}

function ConvertTo-ParagraphWithoutLineSegArray([string]$Paragraph) {
    # 값이 달라진 문단의 줄 배치 캐시는 유효하지 않다. 스타일·문단·표 구조는 건드리지 않고
    # 해당 문단의 캐시만 제거해 한컴이 열 때 줄 배치를 재계산하도록 한다.
    return [regex]::Replace(
        $Paragraph,
        '<hp:linesegarray\b[^>]*>.*?</hp:linesegarray>|<hp:linesegarray\b[^>]*/>',
        '',
        [Text.RegularExpressions.RegexOptions]::Singleline
    )
}

function Convert-TokenizedParagraph([string]$Paragraph, [hashtable]$Map, [string]$EntryName) {
    $nodePattern = '<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>'
    $nodes = @([regex]::Matches($Paragraph, $nodePattern, [Text.RegularExpressions.RegexOptions]::Singleline))
    if ($nodes.Count -eq 0) { return $Paragraph }
    $joined = [Text.StringBuilder]::new(); $starts = @(); $lengths = @()
    foreach ($node in $nodes) { $starts += $joined.Length; $lengths += $node.Groups['text'].Value.Length; [void]$joined.Append($node.Groups['text'].Value) }
    $tokenMatches = @([regex]::Matches($joined.ToString(), $TokenPattern))
    if ($tokenMatches.Count -eq 0) { return $Paragraph }
    foreach ($match in $tokenMatches) {
        $key = $match.Groups[1].Value
        if ($key -match $ForbiddenTokenPattern -or $key -notin $AllowedTokens) { throw "허용되지 않거나 민감한 문서 토큰($EntryName): {{$key}}" }
        if (-not $Map.ContainsKey($key)) { throw "JSON에 값이 없는 문서 토큰($EntryName): {{$key}}" }
    }
    $texts = @($nodes | ForEach-Object { $_.Groups['text'].Value })
    # 뒤에서 앞으로 바꾸므로 같은 런 안에 여러 토큰이 있어도 원본 위치가 유지된다.
    for ($m = $tokenMatches.Count - 1; $m -ge 0; $m--) {
        $match = $tokenMatches[$m]; $first = -1; $last = -1
        for ($i = 0; $i -lt $nodes.Count; $i++) {
            if ($starts[$i] -le $match.Index -and $match.Index -lt ($starts[$i] + $lengths[$i])) { $first = $i }
            $endIndex = $match.Index + $match.Length - 1
            if ($starts[$i] -le $endIndex -and $endIndex -lt ($starts[$i] + $lengths[$i])) { $last = $i; break }
        }
        if ($first -lt 0 -or $last -lt 0) { throw "토큰 런 위치 계산 실패($EntryName): $($match.Value)" }
        $firstOffset = $match.Index - $starts[$first]
        $lastEndOffset = ($match.Index + $match.Length) - $starts[$last]
        $replacement = $Map[$match.Groups[1].Value]
        if ($first -eq $last) {
            $texts[$first] = $texts[$first].Remove($firstOffset, $match.Length).Insert($firstOffset, $replacement)
        } else {
            $texts[$first] = $texts[$first].Substring(0, $firstOffset) + $replacement
            for ($i = $first + 1; $i -lt $last; $i++) { $texts[$i] = '' }
            $texts[$last] = $texts[$last].Substring($lastEndOffset)
        }
    }
    $result = $Paragraph
    for ($i = $nodes.Count - 1; $i -ge 0; $i--) {
        $node = $nodes[$i]; $result = $result.Remove($node.Groups['text'].Index, $node.Groups['text'].Length).Insert($node.Groups['text'].Index, $texts[$i])
    }
    return ConvertTo-ParagraphWithoutLineSegArray $result
}

function Convert-TokenizedXml([string]$Xml, [hashtable]$Map, [string]$EntryName) {
    # 토큰은 문단 경계를 넘을 수 없다. 문단 밖 hp:t도 동일 규칙으로 별도 처리한다.
    $paragraphPattern = '<hp:p\b[^>]*>.*?</hp:p>'
    $updated = [regex]::Replace($Xml, $paragraphPattern, { param($m) Convert-TokenizedParagraph -Paragraph $m.Value -Map $Map -EntryName $EntryName }, [Text.RegularExpressions.RegexOptions]::Singleline)
    $remaining = @()
    foreach ($node in [regex]::Matches($updated, '<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>', [Text.RegularExpressions.RegexOptions]::Singleline)) {
        $remaining += [regex]::Matches($node.Groups['text'].Value, $TokenPattern)
    }
    if ($remaining.Count -gt 0) { throw "치환 뒤 토큰이 남았습니다($EntryName): $($remaining[0].Value)" }
    Assert-Xml $updated $EntryName
    return $updated
}

function Read-ZipEntryText($Entry) {
    $reader = [IO.StreamReader]::new($Entry.Open(), [Text.UTF8Encoding]::new($false, $true), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

$template = (Resolve-Path -LiteralPath $TemplatePath).Path
$absoluteOutput = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetFullPath($template) -eq $absoluteOutput) { throw '출력 경로는 원본 사본 경로와 달라야 합니다.' }
$map = Get-JsonMap $JsonPath
$outputDirectory = Split-Path -Parent $absoluteOutput
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) { [IO.Directory]::CreateDirectory($outputDirectory) | Out-Null }
$temporaryOutput = Join-Path ([IO.Path]::GetDirectoryName($absoluteOutput)) ('.' + [IO.Path]::GetFileName($absoluteOutput) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
try {
    [IO.File]::Copy($template, $temporaryOutput, $false)

$zip = [IO.Compression.ZipFile]::Open($temporaryOutput, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $sections = @($zip.Entries | Where-Object { $_.FullName -match '(^|/)Contents/section\d+\.xml$' })
    if ($sections.Count -eq 0) { throw 'Contents/section*.xml 엔트리를 찾지 못했습니다.' }
    foreach ($section in $sections) {
        $original = Read-ZipEntryText $section
        $updated = Convert-TokenizedXml -Xml $original -Map $map -EntryName $section.FullName
        if ($updated -ne $original) {
            $entryName = $section.FullName; $section.Delete()
            $newEntry = $zip.CreateEntry($entryName, [IO.Compression.CompressionLevel]::Optimal)
            $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
            try { $writer.Write($updated) } finally { $writer.Dispose() }
        }
    }
} finally { $zip.Dispose() }

# 생성 후 ZIP을 다시 열어 모든 XML이 파싱되는지와 잔존 토큰이 없는지 확인한다.
$verifyZip = [IO.Compression.ZipFile]::OpenRead($temporaryOutput)
try {
    foreach ($entry in $verifyZip.Entries) {
        if ($entry.FullName -match '\.xml$') {
            $xml = Read-ZipEntryText $entry; Assert-Xml $xml $entry.FullName
            foreach ($node in [regex]::Matches($xml, '<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>', [Text.RegularExpressions.RegexOptions]::Singleline)) {
                if ([regex]::IsMatch($node.Groups['text'].Value, $TokenPattern)) { throw "출력 HWPX에 미치환 토큰이 남았습니다: $($entry.FullName)" }
            }
        }
    }
} finally { $verifyZip.Dispose() }

# 모든 검증이 끝난 뒤에만 대상 파일을 원자적으로 교체한다.
if ([IO.File]::Exists($absoluteOutput)) {
    $backupOutput = $absoluteOutput + '.' + [guid]::NewGuid().ToString('N') + '.bak'
    [IO.File]::Replace($temporaryOutput, $absoluteOutput, $backupOutput)
    if ([IO.File]::Exists($backupOutput)) {
        try { [IO.File]::Delete($backupOutput) }
        catch { Write-Warning "원자 교체는 완료됐지만 이전 결과 백업을 즉시 삭제하지 못했습니다. 나중에 정리할 수 있습니다: $backupOutput" }
    }
} else { [IO.File]::Move($temporaryOutput, $absoluteOutput) }

    Write-Output "PASS: HWPX 토큰 치환 완료 -> $absoluteOutput"
} finally {
    # 중간 XML/ZIP 검증이 실패해도 호출자가 파일 존재만으로 성공으로 오인하지 않도록
    # 남은 임시 결과를 반드시 제거한다. 원자 교체가 끝난 성공 경로에서는 파일이 이미 없다.
    if (Test-Path -LiteralPath $temporaryOutput -PathType Leaf) {
        Remove-Item -LiteralPath $temporaryOutput -Force -ErrorAction SilentlyContinue
    }
}
