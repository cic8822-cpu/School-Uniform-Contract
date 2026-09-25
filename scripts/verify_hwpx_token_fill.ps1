<#
.SYNOPSIS
  F-007 사본으로 HWPX 명시적 토큰 치환의 자동 회귀 검증을 수행한다.
#>
[CmdletBinding()]
param(
    [string]$SourcePath = 'artifacts/hwpx/P2-01/F007_구매요청_템플릿.hwpx',
    [string]$WorkDirectory = '_workspace/02_hwpx/P2-01/토큰치환_검증산출'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

function Read-ZipText($Entry) { $r=[IO.StreamReader]::new($Entry.Open(),[Text.UTF8Encoding]::new($false,$true),$true); try { $r.ReadToEnd() } finally { $r.Dispose() } }
function Write-ZipText($Zip,[string]$Name,[string]$Text) { $old=$Zip.GetEntry($Name); $old.Delete(); $new=$Zip.CreateEntry($Name,[System.IO.Compression.CompressionLevel]::Optimal); $w=[IO.StreamWriter]::new($new.Open(),[Text.UTF8Encoding]::new($false)); try { $w.Write($Text) } finally { $w.Dispose() } }
function Assert-True([bool]$Condition,[string]$Message) { if (-not $Condition) { throw $Message } }
function Get-EntryMap([string]$Path) { $z=[IO.Compression.ZipFile]::OpenRead($Path); try { $h=@{}; foreach($e in $z.Entries){$h[$e.FullName]=$e.Length}; return $h } finally {$z.Dispose()} }
$KordocVersion = '4.13.1'
function Invoke-Kordoc([string[]]$KordocArguments) { & npx -y "kordoc@$KordocVersion" @KordocArguments; if ($LASTEXITCODE -ne 0) { throw "kordoc $KordocVersion 실패(exit $LASTEXITCODE): $($KordocArguments -join ' ')" } }

$source = (Resolve-Path -LiteralPath $SourcePath).Path
$beforeHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
[IO.Directory]::CreateDirectory($WorkDirectory) | Out-Null
$tokenTemplate = Join-Path $WorkDirectory 'F007_명시적토큰_템플릿.hwpx'
$inputJson = Join-Path $WorkDirectory 'F007_공통값_골든.json'
$output = Join-Path $WorkDirectory 'F007_명시적토큰_골든결과.hwpx'
$render = Join-Path $WorkDirectory 'F007_명시적토큰_골든결과.svg'
[IO.File]::Copy($source, $tokenTemplate, $true)

# 사본의 고정 예시 문자열만 토큰화한다. 원본 및 기준 템플릿은 수정하지 않는다.
$zip=[IO.Compression.ZipFile]::Open($tokenTemplate,[System.IO.Compression.ZipArchiveMode]::Update)
try {
  $sections=@($zip.Entries | Where-Object {$_.FullName -match '(^|/)Contents/section\d+\.xml$'})
  foreach($section in $sections){
    $xml=Read-ZipText $section
    $xml=$xml.Replace('○○학교-○○ (20○○. ○. ○.)','{{문서번호}} ({{발행일}})')
    $xml=$xml.Replace('○○학교','{{학교명}}')
    $xml=$xml.Replace('20○○학년도','{{학년도}}')
    $xml=$xml.Replace('○○학년도','{{학년도}}')
    Write-ZipText $zip $section.FullName $xml
  }
} finally {$zip.Dispose()}

# {{학교명}} 하나를 서로 다른 hp:t 런으로 나눈다. 동일 charPr 속성을 복제해 스타일을 보존한다.
$zip=[IO.Compression.ZipFile]::Open($tokenTemplate,[System.IO.Compression.ZipArchiveMode]::Update)
try {
  $didSplit=$false
  foreach($section in @($zip.Entries | Where-Object {$_.FullName -match '(^|/)Contents/section\d+\.xml$'})){
    $xml=Read-ZipText $section
    if(-not $didSplit){
      $pattern='(?<head><hp:run(?<attrs>\s[^>]*)><hp:t>[^<]*)\{\{학교명\}\}(?<tail>[^<]*</hp:t></hp:run>)'
      $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $match.Groups['head'].Value + '{{학</hp:t></hp:run><hp:run' + $match.Groups['attrs'].Value + '><hp:t>교명}}' + $match.Groups['tail'].Value }
      $next=[regex]::Replace($xml,$pattern,$evaluator,1)
      if($next -ne $xml){$xml=$next;$didSplit=$true}
    }
    Write-ZipText $zip $section.FullName $xml
  }
  Assert-True $didSplit '분절 토큰 시험용 {{학교명}} 런을 만들지 못했습니다.'
} finally {$zip.Dispose()}

$json = @{ 학교명='검증초등학교'; 학년도='2026학년도'; 문서번호='검증초-2026-1'; 발행일='2026. 3. 2.' } | ConvertTo-Json -Compress
[IO.File]::WriteAllText($inputJson,$json,[Text.UTF8Encoding]::new($false))
& (Join-Path $PSScriptRoot 'hwpx_token_fill.ps1') -TemplatePath $tokenTemplate -JsonPath $inputJson -OutputPath $output
if(-not $?) { throw '토큰 치환 스크립트가 실패했습니다.' }

Assert-True ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq $beforeHash) '원본 F-007 사본의 해시가 바뀌었습니다.'
$templateEntries=Get-EntryMap $tokenTemplate; $resultEntries=Get-EntryMap $output
$templateEntryNames = (($templateEntries.Keys | Sort-Object) -join "`n")
$resultEntryNames = (($resultEntries.Keys | Sort-Object) -join "`n")
Assert-True ($templateEntryNames -eq $resultEntryNames) 'ZIP 엔트리 목록이 변경되었습니다.'
$z=[IO.Compression.ZipFile]::OpenRead($output)
try {
  $all=''; foreach($e in $z.Entries | Where-Object {$_.FullName -match '\.xml$'}){$text=Read-ZipText $e; [xml]$null=$text; $all+=$text}
  $textTokens = @(); foreach($node in [regex]::Matches($all,'<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>',[Text.RegularExpressions.RegexOptions]::Singleline)){$textTokens += [regex]::Matches($node.Groups['text'].Value,'\{\{[^{}]+\}\}')}
  Assert-True ($textTokens.Count -eq 0) '미치환 토큰이 남았습니다.'
  foreach($v in @('검증초등학교','2026학년도','검증초-2026-1','2026. 3. 2.')){Assert-True $all.Contains($v) "치환값 누락: $v"}
  $changedParagraphs = 0
  foreach($paragraph in [regex]::Matches($all,'<hp:p\b.*?</hp:p>',[Text.RegularExpressions.RegexOptions]::Singleline)) {
    if($paragraph.Value -match '검증초등학교|2026학년도|검증초-2026-1|2026\. 3\. 2\.') {
      Assert-True ($paragraph.Value -notmatch '<hp:linesegarray\b') '값 치환 문단에 오래된 linesegarray가 남았습니다.'
      $changedParagraphs++
    }
  }
  Assert-True ($changedParagraphs -gt 0) '줄 배치 캐시 검증 대상 문단을 찾지 못했습니다.'
} finally {$z.Dispose()}

# kordoc은 한컴/PDF 수동 검토를 대체하지 않지만, 재열기 가능한 구조와 SVG 렌더 경로를 자동 확인한다.
Invoke-Kordoc @('validate',$output)
Invoke-Kordoc @('render',$output,'-o',$render)
Assert-True (Test-Path -LiteralPath $render) 'SVG 렌더 산출물이 생성되지 않았습니다.'
Write-Output "PASS: ZIP/XML/토큰/원본해시/kordoc validate·render 검증 완료 -> $WorkDirectory"
