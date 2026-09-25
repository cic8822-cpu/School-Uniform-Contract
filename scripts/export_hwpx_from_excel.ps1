<#
.SYNOPSIS
  읽기 전용 XLSM의 허용 공통값으로 명시 토큰 HWPX 사본을 생성한다.

.DESCRIPTION
  허용 입력은 기초자료입력 C4(C-01 학교명), C5(C-02 학년도), C8(C-05 발행일),
  C9(C-06 문서번호), C21(F-007 수신기관), C22(F-007 제목)뿐이다. 교장명, 성명,
  서명, 연락처와 다른 셀/시트는 읽지 않는다. Excel VBA를 실행하거나 수정하지 않는다.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ExcelPath,
    [Parameter(Mandatory)][string]$TokenTemplatePath,
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$WorkingDirectory = '_workspace/02_hwpx/P2-01/excel_hwpx_export'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$KordocVersion = '4.13.1'

function Get-CellText($Sheet, [string]$Address) { return (($Sheet.Range($Address).Text -as [string]).Trim()) }
function Assert-SafeValue([string]$Name, [string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "필수 허용값이 비어 있습니다: $Name" }
    if ($Value -match '[\r\n\x00-\x1F]') { throw "허용값에 줄바꿈/제어 문자가 있습니다: $Name" }
}
function Read-ZipEntryText($Entry) { $reader=[IO.StreamReader]::new($Entry.Open(),[Text.UTF8Encoding]::new($false,$true),$true); try{$reader.ReadToEnd()}finally{$reader.Dispose()} }
function Assert-HwpxOutput([string]$Path) {
    $zip=[IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach($entry in $zip.Entries | Where-Object {$_.FullName -match '\.xml$'}) {
            $xml=Read-ZipEntryText $entry; [xml]$null=$xml
            foreach($node in [regex]::Matches($xml,'<hp:t(?:\s[^>]*)?>(?<text>.*?)</hp:t>',[Text.RegularExpressions.RegexOptions]::Singleline)) {
                if([regex]::IsMatch($node.Groups['text'].Value,'\{\{[^{}]+\}\}')) { throw "출력 HWPX에 미치환 토큰이 남았습니다: $($entry.FullName)" }
            }
        }
    } finally { $zip.Dispose() }
}
function Invoke-Kordoc([string[]]$KordocArguments) { & npx -y "kordoc@$KordocVersion" @KordocArguments; if($LASTEXITCODE -ne 0){throw "kordoc $KordocVersion 실패(exit $LASTEXITCODE): $($KordocArguments -join ' ')"} }
function Get-SharedFileHash([string]$Path) {
    # Get-FileHash는 FileShare.Read로 파일을 여는데, Excel은 통합문서를 열 때
    # FileAccess.ReadWrite로 잠가 다른 프로세스가 쓰기 공유까지 허용해야 열 수 있게
    # 한다. 이 스크립트는 HWPX 작성 버튼을 누른 바로 그 통합문서($ExcelPath)를
    # 대상으로 해시를 구하므로, 표준 Get-FileHash를 쓰면 "다른 프로세스에서 사용
    # 중이므로 액세스할 수 없습니다"(종료 코드 1)로 항상 실패한다(2026-09-24 실측).
    # FileShare.ReadWrite로 직접 스트림을 열어 계산하면 통합문서가 열려 있어도 안전하게
    # 읽기 전용으로 해시를 구할 수 있다.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try {
        $sha256=[Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($sha256.ComputeHash($stream)) -replace '-','') }
        finally { $sha256.Dispose() }
    } finally { $stream.Dispose() }
}

$excelFullPath=(Resolve-Path -LiteralPath $ExcelPath).Path
$templateFullPath=(Resolve-Path -LiteralPath $TokenTemplatePath).Path
$excelHash=Get-SharedFileHash $excelFullPath
$templateHash=Get-SharedFileHash $templateFullPath
$outputFullPath=[IO.Path]::GetFullPath($OutputPath)
[IO.Directory]::CreateDirectory($WorkingDirectory)|Out-Null
$jsonPath=Join-Path ([IO.Path]::GetFullPath($WorkingDirectory)) ('.hwpx_export_' + [guid]::NewGuid().ToString('N') + '.json')
$renderPath=[IO.Path]::ChangeExtension($outputFullPath,'.svg')
$excel=$null; $workbook=$null; $sheet=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AutomationSecurity=3
    # UpdateLinks=0, ReadOnly=$true: 외부 연결과 VBA/수정 저장을 차단한다.
    $workbook=$excel.Workbooks.Open($excelFullPath,0,$true)
    $sheet=$workbook.Worksheets.Item('기초자료입력')
    $schoolName=Get-CellText $sheet 'C4'
    $schoolYearRaw=Get-CellText $sheet 'C5'
    $issueDate=Get-CellText $sheet 'C8'
    $documentNumber=Get-CellText $sheet 'C9'
    $recipient=Get-CellText $sheet 'C21'
    $title=Get-CellText $sheet 'C22'
    foreach($pair in @(@('학교명',$schoolName),@('학년도',$schoolYearRaw),@('발행일',$issueDate),@('문서번호',$documentNumber))){Assert-SafeValue $pair[0] $pair[1]}
    if($schoolYearRaw -notmatch '^\d{4}(학년도)?$'){throw "학년도(C-02)는 4자리 연도여야 합니다: $schoolYearRaw"}
    $schoolYear=if($schoolYearRaw.EndsWith('학년도')){$schoolYearRaw}else{$schoolYearRaw+'학년도'}
    if([string]::IsNullOrWhiteSpace($recipient)){$recipient='내부결재'}
    if([string]::IsNullOrWhiteSpace($title)){$title=$schoolYear+' 교복 학교주관구매 요청'}
    $payload=[ordered]@{학교명=$schoolName;학년도=$schoolYear;발행일=$issueDate;문서번호=$documentNumber;수신기관=$recipient;제목=$title}|ConvertTo-Json -Compress
    [IO.File]::WriteAllText($jsonPath,$payload,[Text.UTF8Encoding]::new($false))
} finally {
    if($workbook){$workbook.Close($false);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)|Out-Null}
    if($sheet){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)|Out-Null}
    if($excel){$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
try {
    & (Join-Path $PSScriptRoot 'hwpx_token_fill.ps1') -TemplatePath $templateFullPath -JsonPath $jsonPath -OutputPath $outputFullPath
    if(-not $?) {throw 'HWPX 토큰 생성기가 실패했습니다.'}
    Assert-HwpxOutput $outputFullPath
    Invoke-Kordoc @('validate',$outputFullPath)
    Invoke-Kordoc @('render',$outputFullPath,'-o',$renderPath)
    if(-not(Test-Path -LiteralPath $renderPath)){throw 'SVG 렌더 산출물이 생성되지 않았습니다.'}
    if((Get-SharedFileHash $excelFullPath) -ne $excelHash){throw '입력 XLSM 해시가 변경되었습니다.'}
    if((Get-SharedFileHash $templateFullPath) -ne $templateHash){throw '입력 HWPX 템플릿 해시가 변경되었습니다.'}
    Write-Output "PASS: Excel 허용값→HWPX 생성/ZIP/XML/토큰0/kordoc validate·render -> $outputFullPath"
} finally {
    if(Test-Path -LiteralPath $jsonPath -PathType Leaf){Remove-Item -LiteralPath $jsonPath -Force -ErrorAction SilentlyContinue}
}
