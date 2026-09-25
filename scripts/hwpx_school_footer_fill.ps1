[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$HwpxPath,
    # 학교정보 시트에 주소·전화가 없는 학교(표본 미포함 등)를 선택하면 호출자가 빈
    # 문자열을 넘긴다. Mandatory 문자열 매개변수는 빈 문자열도 거부해 "인수를 매개변수에
    # 바인딩할 수 없습니다" 오류로 HWPX 작성 전체가 실패했다(실측 확인, 2026-09-24,
    # HWPX 작성 버튼 종료 코드 1 버그의 두 번째 원인). AllowEmptyString으로 허용한다.
    [Parameter(Mandatory)][AllowEmptyString()][string]$SchoolAddress,
    [Parameter(Mandatory)][AllowEmptyString()][string]$SchoolPhone,
    # 2026-09-24 사용자 요청: 하단 결재란의 담당자·교장 성명도 기초자료입력(C10·F10)
    # 값으로 자동반영한다. 기초자료입력에 아직 입력하지 않은 학교는 빈 문자열을 넘긴다
    # (이 경우 원본 플레이스홀더를 그대로 두어 빈칸임을 알 수 있게 한다).
    [Parameter()][AllowEmptyString()][string]$ManagerName = '',
    [Parameter()][AllowEmptyString()][string]$PrincipalName = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

function Set-CellText([Xml.XmlElement]$cell, [Xml.XmlNamespaceManager]$ns, [string]$value) {
    $texts = @($cell.SelectNodes('.//hp:t', $ns))
    if ($texts.Count -eq 0) {
        # 서명란처럼 완전히 빈 칸은 <hp:run/>만 있고 <hp:t> 자체가 없어(2026-09-24 F-007
        # "서명.3"/교장 이름 칸 실측) 채울 텍스트 노드가 없다. 첫 <hp:run> 안에 <hp:t>를
        # 새로 만들어 넣는다.
        $run = $cell.SelectSingleNode('.//hp:run', $ns)
        if (-not $run) { return $false }
        $newText = $cell.OwnerDocument.CreateElement('hp', 't', 'http://www.hancom.co.kr/hwpml/2011/paragraph')
        $newText.InnerText = $value
        [void]$run.AppendChild($newText)
        return $true
    }
    $texts[0].InnerText = $value
    for ($i = 1; $i -lt $texts.Count; $i++) { $texts[$i].InnerText = '' }
    foreach ($paragraph in @($cell.SelectNodes('.//hp:p', $ns))) {
        foreach ($lineSeg in @($paragraph.SelectNodes('./hp:linesegarray', $ns))) { [void]$paragraph.RemoveChild($lineSeg) }
    }
    return $true
}

function Get-CellText([Xml.XmlElement]$cell, [Xml.XmlNamespaceManager]$ns) {
    $texts = @($cell.SelectNodes('.//hp:t', $ns))
    return (($texts | ForEach-Object { $_.InnerText }) -join '').Trim()
}

function Get-CellAddr([Xml.XmlElement]$cell, [Xml.XmlNamespaceManager]$ns) {
    $addr = $cell.SelectSingleNode('.//hp:cellAddr', $ns)
    if (-not $addr) { return $null }
    return [pscustomobject]@{ Row = [int]$addr.GetAttribute('rowAddr'); Col = [int]$addr.GetAttribute('colAddr') }
}

# 담당자·교장 이름은 표마다 정확한 셀 name 속성(예: 직위.2, 서명.3)이 서식별로 달라
# 하드코딩할 수 없다(2026-09-24 F-007 템플릿 실측: "담당자" 라벨 다음 칸은 비어 있고
# 그다음 칸에 "○○○" 플레이스홀더가, "교장" 라벨 다음 칸은 플레이스홀더 없이 바로
# 비어 있음). 따라서 라벨 텍스트("담당자"/"교장")가 정확히 일치하는 칸을 찾은 뒤,
# 같은 행에서 1) "○○○" 플레이스홀더 칸을 우선 채우고, 없으면 2) name 속성이 있는
# 빈 칸 중 라벨에 가장 가까운 칸을 채우는 2단계 규칙으로 서식마다 다른 구조에 대응한다.
function Set-NameNearLabel([Xml.XmlElement]$table, [Xml.XmlNamespaceManager]$ns, [string]$label, [string]$value) {
    if ([string]::IsNullOrEmpty($value)) { return 0 }
    $cells = @($table.SelectNodes('.//hp:tc', $ns))
    $labelCell = $null
    foreach ($c in $cells) { if ((Get-CellText $c $ns) -eq $label) { $labelCell = $c; break } }
    if (-not $labelCell) { return 0 }
    $labelAddr = Get-CellAddr $labelCell $ns
    if (-not $labelAddr) { return 0 }
    $rowCells = @()
    foreach ($c in $cells) {
        $addr = Get-CellAddr $c $ns
        if ($addr -and $addr.Row -eq $labelAddr.Row -and $addr.Col -gt $labelAddr.Col) { $rowCells += [pscustomobject]@{ Cell = $c; Col = $addr.Col } }
    }
    $rowCells = @($rowCells | Sort-Object Col)
    foreach ($rc in $rowCells) {
        if ((Get-CellText $rc.Cell $ns) -eq '○○○') { if (Set-CellText $rc.Cell $ns $value) { return 1 } }
    }
    foreach ($rc in $rowCells) {
        if ((Get-CellText $rc.Cell $ns) -eq '' -and $rc.Cell.GetAttribute('name') -ne '') { if (Set-CellText $rc.Cell $ns $value) { return 1 } }
    }
    return 0
}

$source = (Resolve-Path -LiteralPath $HwpxPath).Path
if ($SchoolAddress -match '[\r\n\x00-\x1F]' -or $SchoolPhone -match '[\r\n\x00-\x1F]') { throw '학교 주소·전화에 줄바꿈 또는 제어 문자가 있습니다.' }
if ($ManagerName -match '[\r\n\x00-\x1F]' -or $PrincipalName -match '[\r\n\x00-\x1F]') { throw '담당자·교장 성명에 줄바꿈 또는 제어 문자가 있습니다.' }
$temporary = Join-Path (Split-Path -Parent $source) ('.' + [IO.Path]::GetFileName($source) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
[IO.File]::Copy($source, $temporary, $false)
$zip = $null
try {
    $zip = [IO.Compression.ZipFile]::Open($temporary, [IO.Compression.ZipArchiveMode]::Update)
    $updated = 0
    foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -match '^Contents/section\d+\.xml$' })) {
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.UTF8Encoding]::new($false, $true), $true)
        try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $doc = [Xml.XmlDocument]::new(); $doc.PreserveWhitespace = $true; $doc.LoadXml($xmlText)
        $ns = [Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('hp', 'http://www.hancom.co.kr/hwpml/2011/paragraph')
        $changed = $false
        foreach ($table in @($doc.SelectNodes('//hp:tbl', $ns))) {
            $postCodeCells = @($table.SelectNodes('.//hp:tc[@name="우편번호"]', $ns))
            if ($postCodeCells.Count -gt 0) {
                foreach ($addressCell in @($table.SelectNodes('.//hp:tc[@name="주소"]', $ns))) {
                    if (Set-CellText $addressCell $ns $SchoolAddress) { $updated++; $changed = $true }
                }
                foreach ($phoneCell in @($table.SelectNodes('.//hp:tc[@name="전화"]', $ns))) {
                    if (Set-CellText $phoneCell $ns $SchoolPhone) { $updated++; $changed = $true }
                }
            }
            # 담당자·교장 결재란은 우편번호와 다른 표(또는 같은 표의 다른 행)에 있을 수 있어
            # 우편번호 존재 여부와 무관하게 모든 표에서 라벨을 찾는다.
            if ((Set-NameNearLabel $table $ns '담당자' $ManagerName) -gt 0) { $updated++; $changed = $true }
            if ((Set-NameNearLabel $table $ns '교장' $PrincipalName) -gt 0) { $updated++; $changed = $true }
        }
        if ($changed) {
            [void]$zip.GetEntry($entry.FullName).Delete()
            $newEntry = $zip.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
            $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
            try { $writer.Write($doc.OuterXml) } finally { $writer.Dispose() }
        }
    }
} finally {
    if ($zip) { $zip.Dispose() }
}
# Windows PowerShell 5.1(.NET Framework)에는 overwrite 매개변수가 있는
# File.Move(string,string,bool) 3-인자 오버로드가 없어(.NET Core/5+ 전용)
# "Move 메서드의 인자 개수(3)를 찾을 수 없습니다" 오류로 항상 실패했다(실측
# 확인, 2026-09-24, HWPX 작성 버튼 종료 코드 1 버그의 실제 원인). 대상 파일을
# 먼저 지우고 2-인자 Move를 쓰면 같은 효과를 낸다.
if ([IO.File]::Exists($source)) { [IO.File]::Delete($source) }
[IO.File]::Move($temporary, $source)
Write-Output "PASS: HWPX 하단 결재란 학교 주소·전화·담당자·교장 성명 반영 ($updated 셀) -> $source"
