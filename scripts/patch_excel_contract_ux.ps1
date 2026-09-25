[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [switch]$InPlace
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$inputFull = $InputPath
if (-not (Test-Path -LiteralPath $inputFull -PathType Leaf)) { throw "입력 XLSM을 찾지 못했습니다: $inputFull" }
$outputFull = [IO.Path]::GetFullPath($OutputPath)
if (-not $InPlace) {
    [IO.Directory]::CreateDirectory((Split-Path -Parent $outputFull)) | Out-Null
    Copy-Item -LiteralPath $inputFull -Destination $outputFull -Force
}

$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($outputFull, 0, $false)

    $guide = $workbook.Worksheets.Item('사용설명서')
    $guide.Range('A1').Value2 = '교복 학교주관구매 길라잡이'
    $guide.Range('A1').Font.Size = 18
    $guide.Range('A1').Font.Bold = $true
    $guide.Range('A3').Value2 = '처음 사용하시나요? 아래 5단계만 따라 하세요.'
    $guide.Range('A4').Value2 = '1단계  파일을 열고 [콘텐츠 사용]을 누릅니다. 버튼을 사용하려면 매크로 허용이 필요합니다.'
    $guide.Range('A5').Value2 = '2단계  [입찰계약절차안내]에서 지금 해야 할 업무와 필요한 서식을 확인합니다.'
    $guide.Range('A6').Value2 = '3단계  [기초자료입력]에 학교명·학년도·구매명 등을 입력하고 [저장하기]를 누릅니다.'
    $guide.Range('A7').Value2 = '4단계  [서식선택_출력]에서 필요한 서식의 [미리보기] 또는 [PDF]를 누릅니다. 입력 전에도 빈 양식을 볼 수 있습니다.'
    $guide.Range('A8').Value2 = '5단계  HWPX가 필요하면 [HWPX 작성]을 누릅니다. 개인정보 가능 서식은 빈 양식으로만 작성합니다.'
    $guide.Range('A3:A8').WrapText = $true
    $guide.Range('A3:A8').Interior.Color = 16777164
    $guide.Range('A3').Font.Bold = $true

    try { $flow = $workbook.Worksheets.Item('입찰계약절차안내') } catch { $flow = $workbook.Worksheets.Item('절차안내'); $flow.Name = '입찰계약절차안내' }
    $flow.Range('A1').Value2 = '입찰·계약 절차 안내'
    $method = $workbook.Worksheets.Item('계약방법안내')
    $method.Range('A2').Value2 = '아래 네 가지 계약방식 중 현재 업무에 맞는 방식을 누르세요. 선택 즉시 기초자료와 출력 서식에 반영되고, 이어서 절차 안내 팝업에서 필요한 단계와 서식을 확인할 수 있습니다.'
    $method.Range('A6').Value2 = '선택 방법의 절차 안내'
    $method.Range('B5').Validation.Delete()
    $method.Range('B5').Validation.Add(3, 1, 1, '1인견적 수의계약,2인견적 수의계약,2단계 입찰(규격·가격 동시),일반경쟁입찰')
    $method.Range('B6').Value2 = '버튼을 누르면 즉시 반영'
    $method.Range('B7').Formula = '=IF(B5<>"",B5,"")'
    $method.Range('C5').Value2 = '계약방식 버튼을 누르면 B-03에 즉시 반영합니다. 추정금액에 따른 자동 판정은 제공하지 않습니다.'
    $method.Range('C6').Formula = '=IF(B5="2단계 입찰(규격·가격 동시)","규격·가격 동시입찰: 규격(적격) 심사 후 가격 개찰",IF(OR(B5="1인견적 수의계약",B5="2인견적 수의계약"),"견적서와 수의계약 사유를 확인",IF(B5="일반경쟁입찰","공고·입찰·개찰·계약 절차를 확인","")))'
    $method.Range('C7').Formula = '=IF(B7<>"","현재 기초자료(B-03): "&B7,"아직 선택하지 않았습니다")'
    $method.Range('A10').Value2 = '아래 카드를 누르면 계약방식이 즉시 반영되고 해당 절차 안내 팝업이 열립니다. 미리보기는 기초자료가 비어 있어도 빈 양식으로 확인할 수 있습니다.'
    $method.Range('A12').Value2 = '계약 방식을 선택하세요'

    foreach ($shapeName in @('btn계약절차_1인견적', 'btn계약절차_2인견적', 'btn계약절차_입찰계약', 'btn계약절차_2단계입찰', 'btn계약절차_일반입찰')) {
        try { $method.Shapes.Item($shapeName).Delete() } catch { }
    }
    $cardTop = $method.Rows.Item(13).Top
    $cardSpecs = @(
        [pscustomobject]@{ Name='btn계약절차_1인견적'; Caption=('STEP 1' + [char]10 + '1인견적 수의계약'); Action='계약방식선택_1인견적'; Left=20; Color=13551615 },
        [pscustomobject]@{ Name='btn계약절차_2인견적'; Caption=('STEP 2' + [char]10 + '2인견적 수의계약'); Action='계약방식선택_2인견적'; Left=180; Color=13421823 },
        [pscustomobject]@{ Name='btn계약절차_2단계입찰'; Caption=('STEP 3' + [char]10 + '2단계 입찰'); Action='계약방식선택_2단계입찰'; Left=340; Color=16763955 },
        [pscustomobject]@{ Name='btn계약절차_일반입찰'; Caption=('STEP 4' + [char]10 + '일반경쟁입찰'); Action='계약방식선택_일반입찰'; Left=500; Color=16768870 }
    )
    foreach ($card in $cardSpecs) {
        $shape = $method.Shapes.AddShape(5, [single]$card.Left, [single]$cardTop, 150, 46)
        $shape.Name = [string]$card.Name
        $shape.OnAction = [string]$card.Action
        $shape.TextFrame.Characters().Text = [string]$card.Caption
        $shape.TextFrame.HorizontalAlignment = -4108
        $shape.Fill.ForeColor.RGB = [int]$card.Color
        $shape.Line.Visible = 0
    }

    $selection = $workbook.Worksheets.Item('서식선택_출력')
    try { $selection.Shapes.Item('btnHWPX작성').Delete() } catch { }
    $hwpxButton = $selection.Shapes.AddShape(5, 640, 10, 150, 24)
    $hwpxButton.Name = 'btnHWPX작성'
    $hwpxButton.OnAction = '선택서식_HWPX작성'
    $hwpxButton.TextFrame.Characters().Text = '선택 서식 HWPX 작성'
    $hwpxButton.Fill.ForeColor.RGB = 16768870
    $hwpxButton.Line.Visible = 0
    $inputSheet = $workbook.Worksheets.Item('기초자료입력')
    try { $inputSheet.Shapes.Item('btn빠른HWPX').Delete() } catch { }
    $quickHwpx = $inputSheet.Shapes.AddShape(5, 860, $inputSheet.Rows.Item(2).Top, 150, 24)
    $quickHwpx.Name = 'btn빠른HWPX'
    $quickHwpx.OnAction = '선택서식_HWPX작성'
    $quickHwpx.TextFrame.Characters().Text = '선택 서식 HWPX 작성'
    $quickHwpx.Fill.ForeColor.RGB = 16768870
    $quickHwpx.Line.Visible = 0

    # 제목 셀만 조정한다. 병합된 제목의 보조 셀은 Excel COM이 서식 변경을 거부할 수 있으므로
    # 실제 제목값이 있는 셀만 처리해 기존 병합·테두리·인쇄영역을 보존한다.
    foreach ($formSheet in $workbook.Worksheets) {
        if ($formSheet.Visible -eq -1 -and $formSheet.Name -match '^F-') {
            for ($rowIndex = 1; $rowIndex -le 8; $rowIndex++) {
                $titleCell = $formSheet.Cells.Item($rowIndex, 2)
                $titleText = [string]$titleCell.Text
                if ($titleText.Length -ge 18 -and [double]$titleCell.Font.Size -ge 15) {
                    try {
                        $titleCell.WrapText = $true
                        $titleCell.HorizontalAlignment = -4108
                        $titleCell.VerticalAlignment = -4108
                        $titleCell.Font.Size = 13
                        if ($formSheet.Rows.Item($rowIndex).RowHeight -lt 32) { $formSheet.Rows.Item($rowIndex).RowHeight = 32 }
                    } catch { }
                }
            }
        }
    }

    foreach ($noticeSheetName in @('F-044_사전안내가정통신문안내', 'F-046_수요조사가정통신문안내', 'F-049_만족도설문조사실시')) {
        $noticeSheet = $workbook.Worksheets.Item($noticeSheetName)
        $noticeRange = $noticeSheet.Range('A1:H1')
        if (-not $noticeRange.MergeCells) { $noticeRange.Merge() }
        $noticeSheet.Range('A1').WrapText = $true
        $noticeSheet.Range('A1').HorizontalAlignment = -4131
        $noticeSheet.Rows.Item(1).RowHeight = 28
    }


    $vbProject = $workbook.VBProject
    try { $vbProject.VBComponents.Remove($vbProject.VBComponents.Item('Module_계약방식_직접선택')) } catch { }
    $contractModule = $vbProject.VBComponents.Add(1)
    $contractModule.Name = 'Module_계약방식_직접선택'
    $contractCode = @'
Option Explicit

Public Sub 계약방식선택_1인견적()
    계약방식적용 "1인견적 수의계약", "1인"
End Sub

Public Sub 계약방식선택_2인견적()
    계약방식적용 "2인견적 수의계약", "2인"
End Sub

Public Sub 계약방식선택_2단계입찰()
    계약방식적용 "2단계 입찰(규격·가격 동시)", "입찰"
End Sub

Public Sub 계약방식선택_일반입찰()
    계약방식적용 "일반경쟁입찰", "입찰"
End Sub

Public Sub 검증_계약방식선택(ByVal methodName As String)
    Sheets("계약방법안내").Range("B5").Value = methodName
    Sheets("기초자료입력").Range("C14").Value = methodName
    Application.Calculate
End Sub

Private Sub 계약방식적용(ByVal methodName As String, ByVal guideType As String)
    Sheets("계약방법안내").Range("B5").Value = methodName
    Sheets("기초자료입력").Range("C14").Value = methodName
    Application.Calculate
    If guideType = "1인" Then
        계약절차안내_1인견적
    ElseIf guideType = "2인" Then
        계약절차안내_2인견적
    Else
        계약절차안내_입찰계약
    End If
End Sub

Public Sub 선택서식_HWPX작성()
    Dim rootPath As String, scriptPath As String, outputPath As String, commandLine As String
    rootPath = ThisWorkbook.Path & "\..\.."
    scriptPath = rootPath & "\scripts\export_hwpx_p204_from_excel.ps1"
    outputPath = rootPath & "\artifacts\hwpx\P2-04_생성"
    If Dir(scriptPath) = "" Then
        MsgBox "HWPX 작성 도구를 찾지 못했습니다. 프로젝트 artifacts\excel 폴더의 배포본에서 실행하세요.", vbExclamation
        Exit Sub
    End If
    commandLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -ExcelPath """ & ThisWorkbook.FullName & """ -OutputDir """ & outputPath & """"
    Shell commandLine, vbNormalFocus
    MsgBox "HWPX 작성을 시작했습니다." & vbCrLf & "완료 파일: " & outputPath, vbInformation
End Sub

Public Sub 선택서식_PDF저장_열기()
    Dim folderPath As String, fileName As String, newestFile As String, newestTime As Date
    Call 선택서식_PDF저장
    folderPath = ThisWorkbook.Path & "\output\"
    fileName = Dir(folderPath & "*.pdf")
    Do While fileName <> ""
        If FileDateTime(folderPath & fileName) >= newestTime Then
            newestTime = FileDateTime(folderPath & fileName)
            newestFile = folderPath & fileName
        End If
        fileName = Dir()
    Loop
    If newestFile <> "" Then ThisWorkbook.FollowHyperlink newestFile
End Sub
'@
    $contractModule.CodeModule.AddFromString($contractCode)

    try { $selection.Shapes.Item('btnPDF저장').OnAction = '선택서식_PDF저장_열기' } catch { }
    try { $inputSheet.Shapes.Item('btn빠른PDF').OnAction = '선택서식_PDF저장_열기' } catch { }

    # 저장 검증은 그대로 두되, 기초자료 공통값이 비어 있는 경우에도 빈 양식 미리보기·PDF를
    # 열 수 있게 출력 경로만 완화한다. F-050은 원자료가 없을 때만 계속 허용한다.
    $outputModule = $vbProject.VBComponents.Item('Module_출력').CodeModule
    $outputCode = $outputModule.Lines(1, $outputModule.CountOfLines)
    $oldOutputGate = '            If fid = "F-050" Then'
    $newOutputGate = "            If fid <> ""F-050"" And Not 검증_필수값검증() Then`r`n                cnt = cnt + 1`r`n                tmp(cnt) = wsSel.Cells(r, 6).Value`r`n            ElseIf fid = ""F-050"" Then"
    if ($outputCode.Contains($oldOutputGate)) {
        $outputCode = $outputCode.Replace($oldOutputGate, $newOutputGate)
        $outputCode = $outputCode.Replace('                If 검증_필수값검증() And F050원자료없음(f050Reason) Then', '                If F050원자료없음(f050Reason) Then')
        $outputModule.DeleteLines(1, $outputModule.CountOfLines)
        $outputModule.AddFromString($outputCode)
    }

    $flow.Move([Type]::Missing, $method)
    $selection.Move([Type]::Missing, $flow)
    $guide.Activate()
    $workbook.Save()
} finally {
    if ($workbook) { $workbook.Close($false); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output "PASS: 기존 검증본에 계약방식·절차안내 UX 구조 반영 -> $outputFull"
