param(
    [string]$OutputPath,
    [string]$InventoryPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot '..\artifacts\excel\교복구매_길라잡이_MVP_초안.xlsm'
}
if ([string]::IsNullOrWhiteSpace($InventoryPath)) {
    $InventoryPath = Join-Path $PSScriptRoot '..\서식_인벤토리.md'
}

function Format-TitleStyle($range) {
    $range.Font.Bold = $true
    $range.Font.Size = 16
    $range.Font.Color = 16777215
    $range.Interior.Color = 10498160
    $range.HorizontalAlignment = -4108
}

function Format-HeaderStyle($range) {
    $range.Font.Bold = $true
    $range.Font.Color = 16777215
    $range.Interior.Color = 12611584
    $range.HorizontalAlignment = -4108
}

function Add-ListValidation($range, [string]$formula) {
    $range.Validation.Delete()
    $range.Validation.Add(3, 1, 1, $formula)
    $range.Validation.IgnoreBlank = $true
    $range.Validation.InCellDropdown = $true
}

$inventoryRows = @()
$inFormIdList = $false
foreach ($line in Get-Content -LiteralPath $InventoryPath -Encoding UTF8) {
    if ($line -eq '## 2. Form ID 목록') {
        $inFormIdList = $true
        continue
    }
    if ($inFormIdList -and $line -match '^## ') {
        break
    }
    if ($inFormIdList -and $line -match '^\| F-0(?:0[1-9]|[1-4][0-9]|5[0-2]) \|') {
        $parts = @($line.Split('|') | ForEach-Object { $_.Trim() })
        $inventoryRows += [PSCustomObject]@{ FormId = $parts[1]; Name = $parts[2]; Phase = $parts[4]; Type = $parts[5] }
    }
}
if ($inventoryRows.Count -ne 52) { throw "출력 대상 Form ID 수가 52개가 아닙니다: $($inventoryRows.Count)" }

$directory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $directory | Out-Null
if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    $sheetNames = @('01_교복구매_워크플로우', '02_기초자료_입력', '03_서식선택_출력', 'DB', 'Ref_data', '학교정보', '서식Metadata', '사용설명서')
    $workbook.Worksheets.Item(1).Name = $sheetNames[0]
    for ($i = 1; $i -lt $sheetNames.Count; $i++) {
        $sheet = $workbook.Worksheets.Add()
        $sheet.Name = $sheetNames[$i]
    }

    $workflow = $workbook.Worksheets.Item('01_교복구매_워크플로우')
    $workflow.Range('A1:F1').Merge()
    $workflow.Range('A1').Value2 = '교복 학교주관구매 업무 흐름'
    Format-TitleStyle $workflow.Range('A1:F1')
    $headers = @('순서', '업무 단계', '주요 확인 사항', '관련 Form ID', '완료 여부', '비고')
    for ($c = 1; $c -le $headers.Count; $c++) { $workflow.Cells.Item(3, $c).Value2 = $headers[$c - 1] }
    Format-HeaderStyle $workflow.Range('A3:F3')
    $stages = @(
        @('1', '기초자료 입력', '학교·사업·일정을 입력하고 필수값을 확인', '공통', '미완료', ''),
        @('2', '구매계획·위원회', '계획 및 위원회 관련 서식을 확인', 'F-001~F-006', '미완료', ''),
        @('3', '입찰·규격', '구매요청·규격·공고 관련 서식을 확인', 'F-007~F-033', '미완료', ''),
        @('4', '평가·선정', '평가표·결과 서식을 확인', 'F-034~F-042', '미완료', ''),
        @('5', '계약·납품', '계약·검수 서식을 확인', 'F-043', '미완료', ''),
        @('6', '안내·조사', '개인정보 자동입력 제외 원칙을 확인', 'F-044~F-052', '미완료', '')
    )
    $row = 4
    foreach ($stage in $stages) {
        for ($c = 1; $c -le $stage.Count; $c++) { $workflow.Cells.Item($row, $c).Value2 = $stage[$c - 1] }
        Add-ListValidation $workflow.Cells.Item($row, 5) '"미완료,진행 중,완료"'
        $row++
    }
    $workflow.Range('A3:F9').Borders.LineStyle = 1
    $workflow.Columns('A:F').AutoFit() | Out-Null
    $workflow.Range('A11:F11').Merge()
    $workflow.Range('A11').Value2 = '주의: 학생·학부모 개인정보, 서명·직인은 자동 입력·저장·출력 대상에서 제외함.'
    $workflow.Range('A11').Font.Bold = $true
    $workflow.Range('A11').Font.Color = 192

    $inputSheet = $workbook.Worksheets.Item('02_기초자료_입력')
    $inputSheet.Range('A1:D1').Merge()
    $inputSheet.Range('A1').Value2 = '기초자료 입력 — 보호 필드는 입력하지 않음'
    Format-TitleStyle $inputSheet.Range('A1:D1')
    $headers = @('필드 ID', '입력 항목', '값', '검증·안내')
    for ($c = 1; $c -le $headers.Count; $c++) { $inputSheet.Cells.Item(3, $c).Value2 = $headers[$c - 1] }
    Format-HeaderStyle $inputSheet.Range('A3:D3')
    $fields = @(
        @('C-01', '학교명', '필수, 2~100자'), @('C-02', '학년도', '필수, 2000~2100'), @('C-03', '담당부서', '조건부, 2~100자'),
        @('C-04', '담당자 직위', '조건부, 2~50자'), @('C-05', '문서 발행일', '조건부, 날짜'), @('C-06', '문서번호', '조건부, 기관 규칙 확인'),
        @('B-01', '구매명', '필수, 2~150자'), @('B-02', '구매 학년', '조건부, 2~100자'), @('B-03', '계약방식', '조건부, 승인 목록 선택'),
        @('B-04', '기초금액', '조건부, 0 이상 정수'), @('B-05', '예정수량', '조건부, 0 이상 정수'), @('B-06', '납품기한', '조건부, 날짜 또는 기간'), @('B-07', '제출기한', '조건부, 날짜·시간')
    )
    $row = 4
    foreach ($field in $fields) {
        $inputSheet.Cells.Item($row, 1).Value2 = $field[0]
        $inputSheet.Cells.Item($row, 2).Value2 = $field[1]
        $inputSheet.Cells.Item($row, 4).Value2 = $field[2]
        $inputSheet.Cells.Item($row, 3).Interior.Color = 13434879
        $row++
    }
    Add-ListValidation $inputSheet.Range('C12') '=DB!$A$2:$A$4'
    $inputSheet.Range('C5').Validation.Add(1, 1, 1, 2000, 2100)
    $inputSheet.Range('C5').Validation.IgnoreBlank = $false
    $inputSheet.Range('C13:C14').NumberFormatLocal = '#,##0'
    $inputSheet.Range('A3:D16').Borders.LineStyle = 1
    $inputSheet.Columns('A:D').AutoFit() | Out-Null
    $inputSheet.Columns('C').ColumnWidth = 28
    $inputSheet.Range('A19:D19').Merge()
    $inputSheet.Range('A19').Value2 = '입력 금지: 학생·학부모 성명·연락처·치수·개별 주문/설문 응답, 직인·서명'
    $inputSheet.Range('A19').Font.Bold = $true
    $inputSheet.Range('A19').Font.Color = 192

    $selection = $workbook.Worksheets.Item('03_서식선택_출력')
    $selection.Range('A1:G1').Merge()
    $selection.Range('A1').Value2 = '서식 선택 및 출력 준비 — HWPX 출력은 P2 검증 완료 전 비활성'
    Format-TitleStyle $selection.Range('A1:G1')
    $headers = @('선택', 'Form ID', '서식명', '업무 단계', '유형', '필수값 상태', '출력 상태')
    for ($c = 1; $c -le $headers.Count; $c++) { $selection.Cells.Item(3, $c).Value2 = $headers[$c - 1] }
    Format-HeaderStyle $selection.Range('A3:G3')
    $row = 4
    foreach ($item in $inventoryRows) {
        $selection.Cells.Item($row, 1).Value2 = '아니오'
        $selection.Cells.Item($row, 2).Value2 = $item.FormId
        $selection.Cells.Item($row, 3).Value2 = $item.Name
        $selection.Cells.Item($row, 4).Value2 = $item.Phase
        $selection.Cells.Item($row, 5).Value2 = $item.Type
        $selection.Cells.Item($row, 6).Formula = '=IF(AND(''02_기초자료_입력''!C4<>"",''02_기초자료_입력''!C5<>"",''02_기초자료_입력''!C10<>""),"준비","필수값 확인")'
        $selection.Cells.Item($row, 7).Formula = '=IF(F' + $row + '="준비","Excel/PDF 준비","출력 보류")'
        Add-ListValidation $selection.Cells.Item($row, 1) '"아니오,예"'
        $row++
    }
    $selection.Range("A3:G$($row - 1)").Borders.LineStyle = 1
    $selection.Range("A3:G$($row - 1)").AutoFilter(1) | Out-Null
    $selection.Columns('A:G').AutoFit() | Out-Null
    $selection.Columns('C').ColumnWidth = 35

    $db = $workbook.Worksheets.Item('DB')
    $db.Range('A1').Value2 = '계약방식'
    $db.Range('A2').Value2 = '일반경쟁입찰'
    $db.Range('A3').Value2 = '제한경쟁입찰'
    $db.Range('A4').Value2 = '수의계약'
    Format-HeaderStyle $db.Range('A1')
    $db.Visible = 0

    $refData = $workbook.Worksheets.Item('Ref_data')
    $headers = @('필드 ID', '필드명', '구분', '자동 처리 원칙')
    for ($c = 1; $c -le $headers.Count; $c++) { $refData.Cells.Item(1, $c).Value2 = $headers[$c - 1] }
    Format-HeaderStyle $refData.Range('A1:D1')
    $calculations = @(@('K-01', '품목별 금액', '계산', '수량×단가'), @('K-02', '합계금액', '계산', '품목별 금액 합계'), @('K-03', '평가 총점', '계산', '평가항목 점수 합계'), @('K-04', '만족도 평균', '계산', '승인된 비식별 집계값만 사용'))
    $row = 2
    foreach ($calculation in $calculations) { for ($c = 1; $c -le 4; $c++) { $refData.Cells.Item($row, $c).Value2 = $calculation[$c - 1] }; $row++ }
    $refData.Columns('A:D').AutoFit() | Out-Null
    $refData.Visible = 0

    $schools = $workbook.Worksheets.Item('학교정보')
    $schools.Range('A1').Value2 = '학교명'; $schools.Range('B1').Value2 = '비고'
    Format-HeaderStyle $schools.Range('A1:B1')
    $schools.Range('A2').Value2 = '직접 입력'; $schools.Range('B2').Value2 = '학교별 기준정보는 기관 확인 후 별도 관리'
    $schools.Columns('A:B').AutoFit() | Out-Null
    $schools.Visible = 0

    $metadata = $workbook.Worksheets.Item('서식Metadata')
    $headers = @('Form ID', '서식명', '업무 단계', '유형')
    for ($c = 1; $c -le $headers.Count; $c++) { $metadata.Cells.Item(1, $c).Value2 = $headers[$c - 1] }
    Format-HeaderStyle $metadata.Range('A1:D1')
    $row = 2
    foreach ($item in $inventoryRows) { $metadata.Cells.Item($row, 1).Value2 = $item.FormId; $metadata.Cells.Item($row, 2).Value2 = $item.Name; $metadata.Cells.Item($row, 3).Value2 = $item.Phase; $metadata.Cells.Item($row, 4).Value2 = $item.Type; $row++ }
    $metadata.Columns('A:D').AutoFit() | Out-Null
    $metadata.Visible = 0

    $guide = $workbook.Worksheets.Item('사용설명서')
    $guide.Range('A1:D1').Merge(); $guide.Range('A1').Value2 = '교복구매 길라잡이 Excel MVP 사용설명서'
    Format-TitleStyle $guide.Range('A1:D1')
    $guide.Range('A3').Value2 = '순서'; $guide.Range('B3').Value2 = '사용 방법'; Format-HeaderStyle $guide.Range('A3:B3')
    $instructions = @('02_기초자료_입력 시트의 노란색 셀에 업무용 공통값만 입력', '03_서식선택_출력 시트에서 필요한 Form ID를 예로 선택', '필수값 상태가 준비인지 확인한 뒤 Excel/PDF 출력 준비', '학생·학부모 개인정보, 서명, 직인 및 개별 설문·치수는 입력하지 않음', 'HWPX 출력은 별도 P2 검증 완료 전까지 사용하지 않음')
    $row = 4
    foreach ($instruction in $instructions) { $guide.Cells.Item($row, 1).Value2 = "$($row - 3)"; $guide.Cells.Item($row, 2).Value2 = $instruction; $row++ }
    $guide.Columns('A:B').AutoFit() | Out-Null
    $guide.Columns('B').ColumnWidth = 85
    $guide.Range('B4:B8').WrapText = $true

    foreach ($sheet in $workbook.Worksheets) {
        $sheet.PageSetup.Orientation = 2
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.LeftMargin = $excel.CentimetersToPoints(1.2)
        $sheet.PageSetup.RightMargin = $excel.CentimetersToPoints(1.2)
    }
    $workflow.Activate()
    $workbook.SaveAs($OutputPath, 52)
    Write-Output "CREATED=$OutputPath"
}
finally {
    if ($workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
