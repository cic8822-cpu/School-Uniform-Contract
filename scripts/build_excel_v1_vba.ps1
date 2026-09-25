# P3-01 클린룸 재구현 — VBA 빌더(2/2): 기초자료 신규/저장/수정/불러오기, 서식선택 출력(PDF) 매크로 추가
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$defaultTargetPath = Join-Path $root 'artifacts\excel\build.xlsm'
$targetPath = if ([string]::IsNullOrWhiteSpace($env:UNIFORM_EXCEL_BUILD_PATH)) { $defaultTargetPath } else { $env:UNIFORM_EXCEL_BUILD_PATH }
$logPath = Join-Path $root '_workspace\03_excel\build_vba_log.txt'

# Quit()·ReleaseComObject·GC만으로는 남은 스크립트 지역변수가 COM 참조를 계속
# 살려 두어 EXCEL.EXE가 좀비로 남을 수 있음. 생성 전후 프로세스 목록을 비교해
# 새로 뜬 PID를 기록해 두고, finally에서 종료가 확인되지 않으면 이 PID만
# 강제 종료해 고아 프로세스가 다음 실행을 막지 않게 함.
$excelPidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 300
$excelProcessId = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $excelPidsBefore -notcontains $_.Id } |
    Select-Object -First 1 -ExpandProperty Id
$excel.Visible = $false
$excel.DisplayAlerts = $false
$log = New-Object System.Text.StringBuilder
function L($s) { [void]$log.AppendLine($s) }

$wb = $null
$saveSucceeded = $false
try {
    $wb = $excel.Workbooks.Open($targetPath, [Type]::Missing, $false)  # ReadOnly=False (편집)

    # ---- 레코드 추적 셀 추가 (기초자료입력!G1/H1) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    $wsIn.Range("G1").Value2 = "불러온 레코드 순번"
    $wsIn.Range("G1").Font.Size = 8
    $wsIn.Range("H1").Value2 = ""
    L "기초자료입력!G1/H1 레코드 추적 셀 추가 완료"

    # 반복행 저장 시트는 구조 빌더가 중단된 경우에도 VBA 빌더 재실행만으로 복구할 수 있게 보장함.
    try {
        $wsItems = $wb.Worksheets.Item("DB_품목")
    } catch {
        $wsItems = $wb.Worksheets.Add()
        $wsItems.Name = "DB_품목"
        $itemHeaders = @("레코드순번", "행번호", "품목명", "수량", "단가", "금액")
        for ($i = 0; $i -lt $itemHeaders.Count; $i++) {
            $wsItems.Cells.Item(1, $i + 1).Value2 = $itemHeaders[$i]
            $wsItems.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsItems.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_품목 시트 복구 생성 완료"
    }
    try {
        $wsVendors = $wb.Worksheets.Item("DB_업체")
    } catch {
        $wsVendors = $wb.Worksheets.Add()
        $wsVendors.Name = "DB_업체"
        $vendorHeaders = @("레코드순번", "행번호", "업체명")
        for ($i = 0; $i -lt $vendorHeaders.Count; $i++) {
            $wsVendors.Cells.Item(1, $i + 1).Value2 = $vendorHeaders[$i]
            $wsVendors.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsVendors.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_업체 시트 복구 생성 완료"
    }
    try {
        $wsCommittee = $wb.Worksheets.Item("DB_위원")
    } catch {
        $wsCommittee = $wb.Worksheets.Add()
        $wsCommittee.Name = "DB_위원"
        $committeeHeaders = @("레코드순번", "행번호", "역할직위", "마스킹식별표시")
        for ($i = 0; $i -lt $committeeHeaders.Count; $i++) {
            $wsCommittee.Cells.Item(1, $i + 1).Value2 = $committeeHeaders[$i]
            $wsCommittee.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsCommittee.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_위원 시트 복구 생성 완료"
    }
    try {
        $wsScore = $wb.Worksheets.Item("DB_평가")
    } catch {
        $wsScore = $wb.Worksheets.Add()
        $wsScore.Name = "DB_평가"
        $scoreHeaders = @("레코드순번", "행번호", "평가항목", "배점", "점수")
        for ($i = 0; $i -lt $scoreHeaders.Count; $i++) {
            $wsScore.Cells.Item(1, $i + 1).Value2 = $scoreHeaders[$i]
            $wsScore.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $wsScore.Rows.Item(1).AutoFilter() | Out-Null
        L "DB_평가 시트 복구 생성 완료"
    }
    try {
        $blankSheet = $wb.Worksheets.Item("Sheet2")
        if ($blankSheet.UsedRange.CountLarge -eq 1 -and [string]::IsNullOrWhiteSpace([string]$blankSheet.Range("A1").Value2)) {
            $excel.DisplayAlerts = $false
            $blankSheet.Delete()
            L "빈 기본 시트 Sheet2 제거 완료"
        }
    } catch { L "제거할 빈 기본 시트 Sheet2 없음" }

    # ---- VBA 모듈 추가 (재실행 대비: 동일 이름 기존 모듈 제거 후 추가) ----
    $vbproj = $wb.VBProject
    foreach ($nm in @("Module_기초자료","Module_출력","Module_검색_절차")) {
        try {
            $existing = $vbproj.VBComponents.Item($nm)
            $vbproj.VBComponents.Remove($existing)
            L "기존 모듈 $nm 제거 후 재생성"
        } catch { L "기존 모듈 $nm 없음" }
    }

    $modInput = $vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    $modInput.Name = "Module_기초자료"
    $codeInput = @'
Option Explicit

Sub 초기화()
    Call 초기화_실행(True)
End Sub

Public Sub 검증_초기화()
    Call 초기화_실행(False)
End Sub

Private Sub 초기화_실행(ByVal showMessage As Boolean)
    With Sheets("기초자료입력")
        .Range("C4").Value = ""
        .Range("C5").Value = ""
        .Range("C6").Value = ""
        .Range("C7").Value = ""
        .Range("C8").Value = ""
        .Range("C9").Value = ""
        .Range("C12").Value = ""
        .Range("C13").Value = ""
        .Range("C14").Value = ""
        .Range("C15").Value = ""
        .Range("C16").Value = ""
        .Range("C17").Value = ""
        .Range("C18").Value = ""
        .Range("C21").Value = ""
        .Range("C22").Value = ""
        .Range("C23").Value = ""
        .Range("C24").Value = ""
        .Range("C25").Value = ""
        .Range("B29:D38").ClearContents
        .Range("B44:B53").ClearContents
        .Range("C44:F53").ClearContents
        .Range("H44:L53").ClearContents
        .Range("N44:Q53").ClearContents
        .Range("B58:C67").ClearContents
        .Range("B72:D81").ClearContents
        .Range("H1").Value = ""
    End With
    If showMessage Then MsgBox "기초자료가 초기화되었습니다.", vbInformation
End Sub

Function 필수값검증() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("기초자료입력")
    필수값검증 = True
    If Trim(ws.Range("C4").Value & "") = "" Then
        MsgBox "학교명(C-01)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C5").Value & "") = "" Then
        MsgBox "학년도(C-02)를 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C12").Value & "") = "" Then
        MsgBox "구매명(B-01)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
    If Trim(ws.Range("C22").Value & "") = "" Then
        MsgBox "제목(D-02)을 입력하세요.", vbExclamation
        필수값검증 = False
        Exit Function
    End If
End Function

Public Function 검증_필수값검증() As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("기초자료입력")
    검증_필수값검증 = Trim(ws.Range("C4").Value & "") <> "" And _
        Trim(ws.Range("C5").Value & "") <> "" And _
        Trim(ws.Range("C12").Value & "") <> "" And _
        Trim(ws.Range("C22").Value & "") <> ""
End Function

Public Function F050원자료없음(ByRef reason As String) As Boolean
    Dim ws As Worksheet
    Set ws = Sheets("F-050_만족도조사설문지")
    If Application.WorksheetFunction.CountA(ws.Range("E14:I23")) > 0 Or _
       Application.WorksheetFunction.CountA(ws.Range("E24:I24")) > 0 Or _
       Application.WorksheetFunction.CountA(ws.Range("B32:I35")) > 0 Then
        reason = "F-050은 인쇄 전용 빈 설문지입니다. 문항 응답·소계·기타 의견은 저장하거나 PDF로 출력할 수 없습니다."
        Exit Function
    End If
    F050원자료없음 = True
End Function

Private Sub 필드복사_기초자료_DB(wsIn As Worksheet, wsDB As Worksheet, r As Long)
    wsDB.Cells(r, 2).Value = wsIn.Range("C4").Value
    wsDB.Cells(r, 3).Value = wsIn.Range("C5").Value
    wsDB.Cells(r, 4).Value = wsIn.Range("C6").Value
    wsDB.Cells(r, 5).Value = wsIn.Range("C7").Value
    wsDB.Cells(r, 6).Value = wsIn.Range("C8").Value
    wsDB.Cells(r, 7).Value = wsIn.Range("C9").Value
    wsDB.Cells(r, 8).Value = wsIn.Range("C12").Value
    wsDB.Cells(r, 9).Value = wsIn.Range("C13").Value
    wsDB.Cells(r, 10).Value = wsIn.Range("C14").Value
    wsDB.Cells(r, 11).Value = wsIn.Range("C15").Value
    wsDB.Cells(r, 12).Value = wsIn.Range("C16").Value
    wsDB.Cells(r, 13).Value = wsIn.Range("C17").Value
    wsDB.Cells(r, 14).Value = wsIn.Range("C18").Value
    wsDB.Cells(r, 15).Value = wsIn.Range("C21").Value
    wsDB.Cells(r, 16).Value = wsIn.Range("C22").Value
    wsDB.Cells(r, 17).Value = wsIn.Range("C23").Value
    wsDB.Cells(r, 18).Value = wsIn.Range("C24").Value
    wsDB.Cells(r, 19).Value = wsIn.Range("C25").Value
    wsDB.Cells(r, 20).Value = Now
    wsDB.Cells(r, 22).Value = wsIn.Range("C10").Value  ' C-08 담당자 성명 (2026-09-23 신설, V열 — U열은 불러오기표시라벨 전용)
    wsDB.Cells(r, 23).Value = wsIn.Range("F10").Value  ' C-07 교장 성명 (2026-09-23 신설, W열)
End Sub

Private Function 품목행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long
    Dim hasName As Boolean, hasQuantity As Boolean, hasPrice As Boolean
    Dim quantityNumeric As Boolean, priceNumeric As Boolean, rowInvalid As Boolean
    품목행검증 = True
    For sourceRow = 29 To 38
        hasName = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> ""
        hasQuantity = Trim(wsIn.Cells(sourceRow, 3).Value & "") <> ""
        hasPrice = Trim(wsIn.Cells(sourceRow, 4).Value & "") <> ""
        If hasName Or hasQuantity Or hasPrice Then
            ' VBA의 Or는 단축평가를 하지 않아, 하나의 Or 체인 안에서 IsNumeric 검사와
            ' CDbl 변환을 같이 두면 빈 칸에서도 CDbl이 실행돼 오류 13이 남(2026-09-21
            ' 실사용 중 발견). 숫자 여부를 먼저 확정한 뒤에만 CDbl 비교를 수행하도록 분리함.
            quantityNumeric = IsNumeric(wsIn.Cells(sourceRow, 3).Value)
            priceNumeric = IsNumeric(wsIn.Cells(sourceRow, 4).Value)
            rowInvalid = Not hasName Or Not hasQuantity Or Not hasPrice Or Not quantityNumeric Or Not priceNumeric
            If Not rowInvalid Then
                rowInvalid = CDbl(wsIn.Cells(sourceRow, 3).Value) <= 0 Or CDbl(wsIn.Cells(sourceRow, 4).Value) <= 0
            End If
            If rowInvalid Then
                If showMessage Then MsgBox "품목 " & (sourceRow - 28) & "행은 품목명·수량·단가를 모두 입력하고, 수량과 단가는 0보다 큰 숫자로 입력하세요.", vbExclamation
                품목행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_품목행검증() As Boolean
    검증_품목행검증 = 품목행검증(Sheets("기초자료입력"), False)
End Function

Private Function 숫자개수(ByVal textValue As String) As Long
    Dim i As Long, count As Long
    For i = 1 To Len(textValue)
        If Mid(textValue, i, 1) Like "#" Then count = count + 1
    Next i
    숫자개수 = count
End Function

Private Function 업체행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String
    업체행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            If Len(vendorName) < 2 Or Len(vendorName) > 150 Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행은 2~150자로 입력하세요.", vbExclamation
                업체행검증 = False
                Exit Function
            End If
            If InStr(vendorName, "대표자") > 0 Or InStr(vendorName, "대표") > 0 Or _
               InStr(vendorName, "연락처") > 0 Or InStr(vendorName, "전화") > 0 Or _
               InStr(vendorName, "휴대폰") > 0 Or InStr(vendorName, "사업자번호") > 0 Or _
               숫자개수(vendorName) >= 8 Or _
               vendorName Like "*[0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*0[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*01[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then
                If showMessage Then MsgBox "업체명만 입력하세요. 대표자·연락처·전화번호·사업자번호는 입력·저장하지 않습니다.", vbExclamation
                업체행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

' R-08(대표자 성명)·R-09(대표자 연락처) 검증(2026-09-23 신설). 업체명(B열)과 달리 이 두 칸은
' 원래부터 성명·연락처를 담는 자리이므로 업체행검증처럼 "성명·전화 패턴을 거부"하지 않고,
' 반대로 형식(성명 2~30자, 연락처는 숫자·하이픈·공백·괄호로만 구성된 7~20자)만 확인한다.
Private Function 전화번호형식허용(ByVal value As String) As Boolean
    Dim i As Long, ch As String
    전화번호형식허용 = False
    If Len(value) < 7 Or Len(value) > 20 Then Exit Function
    For i = 1 To Len(value)
        ch = Mid(value, i, 1)
        If Not (ch Like "#" Or ch = "-" Or ch = " " Or ch = "(" Or ch = ")") Then Exit Function
    Next i
    전화번호형식허용 = True
End Function

Private Function 업체대표자행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String, repName As String, repPhone As String
    업체대표자행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        repName = Trim(wsIn.Cells(sourceRow, 19).Value & "")
        repPhone = Trim(wsIn.Cells(sourceRow, 24).Value & "")
        If repName <> "" Or repPhone <> "" Then
            If vendorName = "" Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행은 업체명 없이 대표자 성명·연락처만 입력할 수 없습니다.", vbExclamation
                업체대표자행검증 = False
                Exit Function
            End If
            If repName <> "" And (Len(repName) < 2 Or Len(repName) > 30 Or InStr(repName, Chr(10)) > 0 Or InStr(repName, Chr(13)) > 0) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행의 대표자 성명은 2~30자로 입력하고 줄바꿈은 사용하지 마세요.", vbExclamation
                업체대표자행검증 = False
                Exit Function
            End If
            If repPhone <> "" And Not 전화번호형식허용(repPhone) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행의 대표자 연락처는 숫자·하이픈 형식(예: 02-000-0000)으로 입력하세요.", vbExclamation
                업체대표자행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_업체대표자행검증() As Boolean
    검증_업체대표자행검증 = 업체대표자행검증(Sheets("기초자료입력"), False)
End Function

Public Function 검증_업체행검증() As Boolean
    검증_업체행검증 = 업체행검증(Sheets("기초자료입력"), False)
End Function

Public Function 검증_업체중복경고() As Boolean
    Dim wsIn As Worksheet, sourceRow As Long, compareRow As Long
    Dim vendorName As String, compareName As String
    Set wsIn = Sheets("기초자료입력")
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            For compareRow = sourceRow + 1 To 53
                compareName = Trim(wsIn.Cells(compareRow, 2).Value & "")
                If compareName <> "" And StrComp(vendorName, compareName, vbTextCompare) = 0 Then
                    검증_업체중복경고 = True
                    Exit Function
                End If
            Next compareRow
        End If
    Next sourceRow
    검증_업체중복경고 = False
End Function

Private Function 정량평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    ' 업체명만 먼저 등록하고 정량평가 점수는 평가 이후 입력하는 것이 정상 흐름이므로,
    ' 점수 4칸이 모두 비어 있으면 통과시키고, 하나라도 있으면 4칸 모두·범위까지 완전해야 저장을 허용한다.
    ' (F-015 출력 가능 여부는 별도의 F015업체행완전한가에서 더 엄격하게 판단한다.)
    Dim sourceRow As Long, vendorName As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    정량평가행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 3).Value
        v2 = wsIn.Cells(sourceRow, 4).Value
        v3 = wsIn.Cells(sourceRow, 5).Value
        v4 = wsIn.Cells(sourceRow, 6).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 업체명 없이 정량평가 점수만 입력할 수 없습니다.", vbExclamation
            정량평가행검증 = False
            Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 수행경험·공인인증·거리적접근성·상한가격 점수를 모두 입력하거나 모두 비워 두세요.", vbExclamation
            정량평가행검증 = False
            Exit Function
        ElseIf hasAllScore Then
            If Not 정수범위(v1, 0, 10) Or Not 정수범위(v2, 0, 10) Or Not 정수범위(v3, 0, 15) Or Not 정수범위(v4, 0, 15) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-015)은 수행경험(0~10)·공인인증(0~10)·거리적접근성(0~15)·상한가격(0~15) 범위의 정수만 입력하세요.", vbExclamation
                정량평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_정량평가행검증() As Boolean
    검증_정량평가행검증 = 정량평가행검증(Sheets("기초자료입력"), False)
End Function

Public Function 정성등급유효(ByVal grade As Variant) As Boolean
    Select Case Trim(grade & "")
        Case "탁월", "우수", "보통", "미흡", "불량": 정성등급유효 = True
        Case Else: 정성등급유효 = False
    End Select
End Function

Private Function 정성등급점수(ByVal grade As Variant, ByVal maxScore As Long) As Long
    Select Case Trim(grade & "")
        Case "탁월": 정성등급점수 = maxScore
        Case "우수": 정성등급점수 = IIf(maxScore = 15, 12, 8)
        Case "보통": 정성등급점수 = IIf(maxScore = 15, 9, 6)
        Case "미흡": 정성등급점수 = IIf(maxScore = 15, 6, 4)
        Case "불량": 정성등급점수 = IIf(maxScore = 15, 3, 2)
        Case Else: 정성등급점수 = -1
    End Select
End Function

Private Function 정성점수유효(ByVal score As Variant, ByVal maxScore As Long) As Boolean
    Dim expectedScore As Variant
    For Each expectedScore In Array(maxScore, IIf(maxScore = 15, 12, 8), IIf(maxScore = 15, 9, 6), IIf(maxScore = 15, 6, 4), IIf(maxScore = 15, 3, 2))
        If score = expectedScore Then
            정성점수유효 = True
            Exit Function
        End If
    Next expectedScore
End Function

Private Function 정성점수등급(ByVal score As Variant, ByVal maxScore As Long) As String
    Select Case score
        Case maxScore: 정성점수등급 = "탁월"
        Case IIf(maxScore = 15, 12, 8): 정성점수등급 = "우수"
        Case IIf(maxScore = 15, 9, 6): 정성점수등급 = "보통"
        Case IIf(maxScore = 15, 6, 4): 정성점수등급 = "미흡"
        Case IIf(maxScore = 15, 3, 2): 정성점수등급 = "불량"
        Case Else: 정성점수등급 = ""
    End Select
End Function

Private Function 정성평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant, adjustment As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    정성평가행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 8).Value
        v2 = wsIn.Cells(sourceRow, 9).Value
        v3 = wsIn.Cells(sourceRow, 10).Value
        v4 = wsIn.Cells(sourceRow, 11).Value
        adjustment = wsIn.Cells(sourceRow, 12).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> "" Or Trim(adjustment & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> "" And Trim(adjustment & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 업체명 없이 정성평가 점수만 입력할 수 없습니다.", vbExclamation
            정성평가행검증 = False
            Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 재질·완성도·A/S·하자보상·가감점을 모두 입력하거나 모두 비워 두세요.", vbExclamation
            정성평가행검증 = False
            Exit Function
        ElseIf hasAllScore Then
            If Not 정성등급유효(v1) Or Not 정성등급유효(v2) Or Not 정성등급유효(v3) Or Not 정성등급유효(v4) Or Not 정수범위(adjustment, -15, 5) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-016)은 재질·완성도·A/S·하자보상에 탁월·우수·보통·미흡·불량 중 하나를 선택하고, 가감점은 -15~5 범위의 정수로 입력하세요.", vbExclamation
                정성평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_정성평가행검증() As Boolean
    검증_정성평가행검증 = 정성평가행검증(Sheets("기초자료입력"), False)
End Function

Private Function 자기평점행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, vendorName As String, v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant
    Dim hasAnyScore As Boolean, hasAllScore As Boolean
    자기평점행검증 = True
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        v1 = wsIn.Cells(sourceRow, 14).Value: v2 = wsIn.Cells(sourceRow, 15).Value: v3 = wsIn.Cells(sourceRow, 16).Value: v4 = wsIn.Cells(sourceRow, 17).Value
        hasAnyScore = Trim(v1 & "") <> "" Or Trim(v2 & "") <> "" Or Trim(v3 & "") <> "" Or Trim(v4 & "") <> ""
        hasAllScore = Trim(v1 & "") <> "" And Trim(v2 & "") <> "" And Trim(v3 & "") <> "" And Trim(v4 & "") <> ""
        If hasAnyScore And vendorName = "" Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 업체명 없이 자기평점만 입력할 수 없습니다.", vbExclamation
            자기평점행검증 = False: Exit Function
        ElseIf hasAnyScore And Not hasAllScore Then
            If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 수행경험·공인인증·거리적접근성·상한가격 자기평점을 모두 입력하거나 모두 비워 두세요.", vbExclamation
            자기평점행검증 = False: Exit Function
        ElseIf hasAllScore Then
            If Not 정수범위(v1, 0, 10) Or Not 정수범위(v2, 0, 10) Or Not 정수범위(v3, 0, 15) Or Not 정수범위(v4, 0, 15) Then
                If showMessage Then MsgBox "업체 " & (sourceRow - 43) & "행(F-018)은 수행경험(0~10)·공인인증(0~10)·거리적접근성(0~15)·상한가격(0~15) 범위의 정수만 입력하세요.", vbExclamation
                자기평점행검증 = False: Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_자기평점행검증() As Boolean
    검증_자기평점행검증 = 자기평점행검증(Sheets("기초자료입력"), False)
End Function

' 2026-09-23: 마스킹 식별표시 승인목록 방식을 폐지하고 위원 실명(2~30자, 줄바꿈 금지)
' 형식만 확인한다. 연락처·서명은 여전히 별도 칸이 없어 입력·저장할 수 없다.
Private Function 위원행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, roleName As String, memberName As String
    Dim hasRole As Boolean, hasName As Boolean
    위원행검증 = True
    For sourceRow = 58 To 67
        roleName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        memberName = Trim(wsIn.Cells(sourceRow, 3).Value & "")
        hasRole = roleName <> ""
        hasName = memberName <> ""
        If hasRole Or hasName Then
            If Not hasRole Or Not hasName Or Len(roleName) < 2 Or Len(roleName) > 50 Then
                If showMessage Then MsgBox "위원 " & (sourceRow - 57) & "행은 역할·직위(2~50자)와 위원 성명을 모두 입력하세요.", vbExclamation
                위원행검증 = False
                Exit Function
            End If
            If Len(memberName) < 2 Or Len(memberName) > 30 Or InStr(memberName, Chr(10)) > 0 Or InStr(memberName, Chr(13)) > 0 Then
                If showMessage Then MsgBox "위원 " & (sourceRow - 57) & "행의 위원 성명은 2~30자로 입력하고 줄바꿈은 사용하지 마세요.", vbExclamation
                위원행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 검증_위원행검증() As Boolean
    검증_위원행검증 = 위원행검증(Sheets("기초자료입력"), False)
End Function

Public Function 검증_위원중복경고() As Boolean
    Dim wsIn As Worksheet, sourceRow As Long, compareRow As Long
    Dim roleName As String, compareRole As String
    Set wsIn = Sheets("기초자료입력")
    For sourceRow = 58 To 67
        roleName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If roleName <> "" Then
            For compareRow = sourceRow + 1 To 67
                compareRole = Trim(wsIn.Cells(compareRow, 2).Value & "")
                If compareRole <> "" And StrComp(roleName, compareRole, vbTextCompare) = 0 Then
                    검증_위원중복경고 = True
                    Exit Function
                End If
            Next compareRow
        End If
    Next sourceRow
    검증_위원중복경고 = False
End Function

Private Function 평가행검증(ByVal wsIn As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourceRow As Long, itemName As String
    Dim hasItem As Boolean, hasAllocation As Boolean, hasScore As Boolean
    Dim allocationNumeric As Boolean, scoreNumeric As Boolean, rowInvalid As Boolean
    평가행검증 = True
    For sourceRow = 72 To 81
        itemName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        hasItem = itemName <> ""
        hasAllocation = Trim(wsIn.Cells(sourceRow, 3).Value & "") <> ""
        hasScore = Trim(wsIn.Cells(sourceRow, 4).Value & "") <> ""
        If hasItem Or hasAllocation Or hasScore Then
            ' 품목행검증과 같은 이유(Or 단축평가 없음)로 IsNumeric·CDbl을 분리함(2026-09-21).
            allocationNumeric = IsNumeric(wsIn.Cells(sourceRow, 3).Value)
            scoreNumeric = IsNumeric(wsIn.Cells(sourceRow, 4).Value)
            rowInvalid = Not hasItem Or Not hasAllocation Or Not hasScore Or Len(itemName) < 2 Or Len(itemName) > 100 Or _
                Not allocationNumeric Or Not scoreNumeric
            If Not rowInvalid Then
                rowInvalid = CDbl(wsIn.Cells(sourceRow, 3).Value) < 0 Or CDbl(wsIn.Cells(sourceRow, 4).Value) < 0 Or _
                    CDbl(wsIn.Cells(sourceRow, 4).Value) > CDbl(wsIn.Cells(sourceRow, 3).Value)
            End If
            If rowInvalid Then
                If showMessage Then MsgBox "평가 " & (sourceRow - 71) & "행은 평가항목(2~100자), 0 이상 배점, 0 이상이며 배점 이하인 점수를 모두 입력하세요.", vbExclamation
                평가행검증 = False
                Exit Function
            End If
        End If
    Next sourceRow
End Function

Public Function 정수범위(ByVal value As Variant, ByVal minimum As Long, ByVal maximum As Long) As Boolean
    If Not IsNumeric(value) Then Exit Function
    If CDbl(value) <> Fix(CDbl(value)) Then Exit Function
    정수범위 = CDbl(value) >= minimum And CDbl(value) <= maximum
End Function

Private Function 양수숫자(ByVal value As Variant) As Boolean
    양수숫자 = IsNumeric(value)
    If 양수숫자 Then 양수숫자 = CDbl(value) > 0
End Function

Private Function 음이아닌숫자(ByVal value As Variant) As Boolean
    음이아닌숫자 = IsNumeric(value)
    If 음이아닌숫자 Then 음이아닌숫자 = CDbl(value) >= 0
End Function

Private Function 품목금액일치(ByVal quantity As Variant, ByVal unitPrice As Variant, ByVal amount As Variant) As Boolean
    If Not 양수숫자(quantity) Or Not 양수숫자(unitPrice) Or Not IsNumeric(amount) Then Exit Function
    품목금액일치 = CDbl(amount) = CDbl(quantity) * CDbl(unitPrice)
End Function

Private Function 평가점수범위(ByVal allocation As Variant, ByVal score As Variant) As Boolean
    If Not 음이아닌숫자(allocation) Or Not 음이아닌숫자(score) Then Exit Function
    평가점수범위 = CDbl(score) <= CDbl(allocation)
End Function

Private Function 마지막사용행(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        마지막사용행 = 1
    Else
        마지막사용행 = lastCell.Row
    End If
End Function

Private Function DB레코드존재(ByVal wsDB As Worksheet, ByVal seq As Long) As Boolean
    Dim rowIndex As Long
    For rowIndex = 2 To 마지막사용행(wsDB)
        If wsDB.Cells(rowIndex, 1).Value = seq Then
            DB레코드존재 = True
            Exit Function
        End If
    Next rowIndex
End Function

' 2026-09-24: 개인정보 보호(F-050 원자료 비저장)는 "완성도 검증은 경고만" 결정과 별개로
' 계속 저장을 막아야 하는 유일한 치명적 사유라서 저장경계검증에서 분리함.
Public Function 저장경계_치명적검증(ByRef reason As String) As Boolean
    저장경계_치명적검증 = F050원자료없음(reason)
End Function

Private Function 저장경계검증(ByRef reason As String) As Boolean
    Dim wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Dim rowIndex As Long, lastRow As Long, seq As Long
    Dim vendorName As String, roleName As String, maskedLabel As String, itemName As String
    Dim repName As String, repPhone As String
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")

    lastRow = 마지막사용행(wsDB)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsDB.Range("A" & rowIndex & ":T" & rowIndex)) > 0 Then
            If Not 정수범위(wsDB.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Trim(wsDB.Cells(rowIndex, 2).Value & "") = "" Or Trim(wsDB.Cells(rowIndex, 3).Value & "") = "" Or _
               Trim(wsDB.Cells(rowIndex, 8).Value & "") = "" Or Trim(wsDB.Cells(rowIndex, 16).Value & "") = "" Then
                reason = "DB " & rowIndex & "행은 순번과 학교명·학년도·구매명·제목을 모두 갖춘 업무 레코드여야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsItems)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsItems.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsItems.Cells(rowIndex, 1).Value, 1, 2147483647) Then
                reason = "DB_품목 " & rowIndex & "행의 레코드순번이 올바르지 않습니다."
                Exit Function
            End If
            seq = CLng(wsItems.Cells(rowIndex, 1).Value)
            If Not DB레코드존재(wsDB, seq) Or Not 정수범위(wsItems.Cells(rowIndex, 2).Value, 1, 10) Or _
               Trim(wsItems.Cells(rowIndex, 3).Value & "") = "" Or _
               Not 품목금액일치(wsItems.Cells(rowIndex, 4).Value, wsItems.Cells(rowIndex, 5).Value, wsItems.Cells(rowIndex, 6).Value) Then
                reason = "DB_품목 " & rowIndex & "행은 연결된 레코드와 완전한 품목·수량·단가·금액을 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsVendors)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsVendors.Range("A" & rowIndex & ":E" & rowIndex)) > 0 Then
            vendorName = Trim(wsVendors.Cells(rowIndex, 3).Value & "")
            repName = Trim(wsVendors.Cells(rowIndex, 4).Value & "")
            repPhone = Trim(wsVendors.Cells(rowIndex, 5).Value & "")
            If Not 정수범위(wsVendors.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsVendors.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsVendors.Cells(rowIndex, 2).Value, 1, 10) Or Len(vendorName) < 2 Or Len(vendorName) > 150 Or _
               InStr(vendorName, "대표자") > 0 Or InStr(vendorName, "대표") > 0 Or InStr(vendorName, "연락처") > 0 Or _
               InStr(vendorName, "전화") > 0 Or InStr(vendorName, "휴대폰") > 0 Or InStr(vendorName, "사업자번호") > 0 Or _
               숫자개수(vendorName) >= 8 Or vendorName Like "*[0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*0[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Or _
               vendorName Like "*01[0-9]-[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then
                reason = "DB_업체 " & rowIndex & "행의 업체명 칸에는 상호명만 허용되며 대표자·연락처·전화번호·사업자번호는 별도 칸(대표자성명·대표자연락처)에만 저장할 수 있습니다."
                Exit Function
            End If
            If repName <> "" And (Len(repName) < 2 Or Len(repName) > 30) Then
                reason = "DB_업체 " & rowIndex & "행의 대표자 성명은 2~30자여야 합니다."
                Exit Function
            End If
            If repPhone <> "" And Not 전화번호형식허용(repPhone) Then
                reason = "DB_업체 " & rowIndex & "행의 대표자 연락처는 숫자·하이픈 형식이어야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsCommittee)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsCommittee.Range("A" & rowIndex & ":D" & rowIndex)) > 0 Then
            roleName = Trim(wsCommittee.Cells(rowIndex, 3).Value & "")
            maskedLabel = Trim(wsCommittee.Cells(rowIndex, 4).Value & "")
            If Not 정수범위(wsCommittee.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsCommittee.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsCommittee.Cells(rowIndex, 2).Value, 1, 10) Or Len(roleName) < 2 Or Len(roleName) > 50 Or _
               Len(maskedLabel) < 2 Or Len(maskedLabel) > 30 Then
                reason = "DB_위원 " & rowIndex & "행은 역할·직위(2~50자)와 위원 성명(2~30자)을 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsScore)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsScore.Range("A" & rowIndex & ":E" & rowIndex)) > 0 Then
            itemName = Trim(wsScore.Cells(rowIndex, 3).Value & "")
            If Not 정수범위(wsScore.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsScore.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsScore.Cells(rowIndex, 2).Value, 1, 10) Or Len(itemName) < 2 Or Len(itemName) > 100 Or _
               Not 평가점수범위(wsScore.Cells(rowIndex, 4).Value, wsScore.Cells(rowIndex, 5).Value) Then
                reason = "DB_평가 " & rowIndex & "행은 평가항목·배점·점수가 완전하고 점수가 배점 이하이어야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex

    lastRow = 마지막사용행(wsQuant)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsQuant.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsQuant.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsQuant.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 2).Value, 1, 10) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 3).Value, 0, 10) Or Not 정수범위(wsQuant.Cells(rowIndex, 4).Value, 0, 10) Or _
               Not 정수범위(wsQuant.Cells(rowIndex, 5).Value, 0, 15) Or Not 정수범위(wsQuant.Cells(rowIndex, 6).Value, 0, 15) Then
                reason = "DB_정량평가 " & rowIndex & "행은 연결된 레코드와 정상 범위의 수행경험·공인인증·거리적접근성·상한가격 점수를 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    lastRow = 마지막사용행(wsQual)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsQual.Range("A" & rowIndex & ":G" & rowIndex)) > 0 Then
            If Not 정수범위(wsQual.Cells(rowIndex, 1).Value, 1, 2147483647) Or _
               Not DB레코드존재(wsDB, CLng(Val(wsQual.Cells(rowIndex, 1).Value))) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 2).Value, 1, 10) Or _
                Not 정성점수유효(wsQual.Cells(rowIndex, 3).Value, 15) Or Not 정성점수유효(wsQual.Cells(rowIndex, 4).Value, 10) Or _
                Not 정성점수유효(wsQual.Cells(rowIndex, 5).Value, 15) Or Not 정성점수유효(wsQual.Cells(rowIndex, 6).Value, 10) Or _
               Not 정수범위(wsQual.Cells(rowIndex, 7).Value, -15, 5) Then
                reason = "DB_정성평가 " & rowIndex & "행은 연결된 레코드와 정상 범위의 재질·완성도·A/S·하자보상·가감점을 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    lastRow = 마지막사용행(wsSelf)
    For rowIndex = 2 To lastRow
        If Application.WorksheetFunction.CountA(wsSelf.Range("A" & rowIndex & ":F" & rowIndex)) > 0 Then
            If Not 정수범위(wsSelf.Cells(rowIndex, 1).Value, 1, 2147483647) Or Not DB레코드존재(wsDB, CLng(Val(wsSelf.Cells(rowIndex, 1).Value))) Or Not 정수범위(wsSelf.Cells(rowIndex, 2).Value, 1, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 3).Value, 0, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 4).Value, 0, 10) Or Not 정수범위(wsSelf.Cells(rowIndex, 5).Value, 0, 15) Or Not 정수범위(wsSelf.Cells(rowIndex, 6).Value, 0, 15) Then
                reason = "DB_자기평점 " & rowIndex & "행은 연결된 레코드와 정상 범위의 업체 자기평점을 모두 가져야 합니다."
                Exit Function
            End If
        End If
    Next rowIndex
    저장경계검증 = True
End Function

Public Function 검증_저장경계() As Boolean
    Dim reason As String
    검증_저장경계 = 저장경계검증(reason)
End Function

Public Function 검증_저장경계_메시지() As String
    Dim reason As String
    If Not 저장경계검증(reason) Then 검증_저장경계_메시지 = reason
End Function

Public Function 검증_평가행검증() As Boolean
    검증_평가행검증 = 평가행검증(Sheets("기초자료입력"), False)
End Function

Private Sub 품목복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsItems As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsItems.Cells(rowIndex, 1).Value = seq Then wsItems.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 29 To 38
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsItems.Cells(targetRow, 1).Value = seq
            wsItems.Cells(targetRow, 2).Value = sourceRow - 28
            wsItems.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsItems.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            wsItems.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 4).Value
            wsItems.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 5).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 업체복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsVendors As Worksheet, ByVal seq As Long)
    ' S열(19)=대표자 성명(R-08)·X열(24)=대표자 연락처(R-09), 2026-09-23 신설.
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsVendors.Cells(rowIndex, 1).Value = seq Then wsVendors.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" Then
            wsVendors.Cells(targetRow, 1).Value = seq
            wsVendors.Cells(targetRow, 2).Value = sourceRow - 43
            wsVendors.Cells(targetRow, 3).Value = vendorName
            wsVendors.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 19).Value
            wsVendors.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 24).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 정량평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsQuant As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsQuant.Cells(rowIndex, 1).Value = seq Then wsQuant.Rows(rowIndex).Delete
    Next rowIndex

    targetRow = wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" And Trim(wsIn.Cells(sourceRow, 3).Value & "") <> "" Then
            wsQuant.Cells(targetRow, 1).Value = seq
            wsQuant.Cells(targetRow, 2).Value = sourceRow - 43
            wsQuant.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 3).Value
            wsQuant.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 4).Value
            wsQuant.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 5).Value
            wsQuant.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 6).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 정성평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsQual As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long, vendorName As String
    For rowIndex = wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsQual.Cells(rowIndex, 1).Value = seq Then wsQual.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        vendorName = Trim(wsIn.Cells(sourceRow, 2).Value & "")
        If vendorName <> "" And Trim(wsIn.Cells(sourceRow, 8).Value & "") <> "" Then
            wsQual.Cells(targetRow, 1).Value = seq
            wsQual.Cells(targetRow, 2).Value = sourceRow - 43
            wsQual.Cells(targetRow, 3).Value = 정성등급점수(wsIn.Cells(sourceRow, 8).Value, 15)
            wsQual.Cells(targetRow, 4).Value = 정성등급점수(wsIn.Cells(sourceRow, 9).Value, 10)
            wsQual.Cells(targetRow, 5).Value = 정성등급점수(wsIn.Cells(sourceRow, 10).Value, 15)
            wsQual.Cells(targetRow, 6).Value = 정성등급점수(wsIn.Cells(sourceRow, 11).Value, 10)
            wsQual.Cells(targetRow, 7).Value = wsIn.Cells(sourceRow, 12).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 자기평점복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsSelf As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsSelf.Cells(rowIndex, 1).Value = seq Then wsSelf.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 44 To 53
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 14).Value & "") <> "" Then
            wsSelf.Cells(targetRow, 1).Value = seq: wsSelf.Cells(targetRow, 2).Value = sourceRow - 43
            wsSelf.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 14).Value: wsSelf.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 15).Value
            wsSelf.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 16).Value: wsSelf.Cells(targetRow, 6).Value = wsIn.Cells(sourceRow, 17).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 위원복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsCommittee As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsCommittee.Cells(rowIndex, 1).Value = seq Then wsCommittee.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 58 To 67
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsCommittee.Cells(targetRow, 1).Value = seq
            wsCommittee.Cells(targetRow, 2).Value = sourceRow - 57
            wsCommittee.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsCommittee.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Private Sub 평가복사_기초자료_DB(ByVal wsIn As Worksheet, ByVal wsScore As Worksheet, ByVal seq As Long)
    Dim rowIndex As Long, sourceRow As Long, targetRow As Long
    For rowIndex = wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsScore.Cells(rowIndex, 1).Value = seq Then wsScore.Rows(rowIndex).Delete
    Next rowIndex
    targetRow = wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2
    For sourceRow = 72 To 81
        If Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" Then
            wsScore.Cells(targetRow, 1).Value = seq
            wsScore.Cells(targetRow, 2).Value = sourceRow - 71
            wsScore.Cells(targetRow, 3).Value = wsIn.Cells(sourceRow, 2).Value
            wsScore.Cells(targetRow, 4).Value = wsIn.Cells(sourceRow, 3).Value
            wsScore.Cells(targetRow, 5).Value = wsIn.Cells(sourceRow, 4).Value
            targetRow = targetRow + 1
        End If
    Next sourceRow
End Sub

Sub 저장하기()
    Call 저장하기_실행(True)
End Sub

Public Sub 검증_저장하기()
    Call 저장하기_실행(False)
End Sub

Private Sub 저장하기_실행(ByVal showMessage As Boolean)
    ' 2026-09-24 사용자 결정: 필수값·행 단위 완성도 검증 실패는 더 이상 저장 자체를
    ' 막지 않는다(경고만 표시하고 입력된 그대로 저장을 계속 진행함). 개인정보 보호 목적의
    ' F-050 원자료 차단은 이 결정과 무관하게 Workbook_BeforeSave에서 계속 저장을 막는다.
    Call 필수값검증()
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    Call 품목행검증(wsIn, showMessage)
    Call 업체행검증(wsIn, showMessage)
    Call 업체대표자행검증(wsIn, showMessage)
    Call 정량평가행검증(wsIn, showMessage)
    Call 정성평가행검증(wsIn, showMessage)
    Call 자기평점행검증(wsIn, showMessage)
    Call 위원행검증(wsIn, showMessage)
    Call 평가행검증(wsIn, showMessage)
    If showMessage And 검증_업체중복경고() Then MsgBox "중복된 업체명이 있습니다. 실제 동일 업체인지 확인한 뒤 저장하세요.", vbExclamation
    If showMessage And 검증_위원중복경고() Then MsgBox "중복된 위원 역할·직위가 있습니다. 실제 구성과 역할을 확인하세요.", vbExclamation

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    Dim newRow As Long
    If lastRow <= 1 And wsDB.Cells(2, 1).Value = "" Then
        newRow = 2
    Else
        newRow = lastRow + 1
    End If
    Dim seq As Long
    seq = newRow - 1

    wsDB.Cells(newRow, 1).Value = seq
    Call 필드복사_기초자료_DB(wsIn, wsDB, newRow)
    Call 품목복사_기초자료_DB(wsIn, wsItems, seq)
    Call 업체복사_기초자료_DB(wsIn, wsVendors, seq)
    Call 정량평가복사_기초자료_DB(wsIn, wsQuant, seq)
    Call 정성평가복사_기초자료_DB(wsIn, wsQual, seq)
    Call 자기평점복사_기초자료_DB(wsIn, wsSelf, seq)
    Call 위원복사_기초자료_DB(wsIn, wsCommittee, seq)
    Call 평가복사_기초자료_DB(wsIn, wsScore, seq)

    wsIn.Range("H1").Value = seq
    ' 2026-09-23 실사용 버그 수정: 기존에는 DB 시트에만 기록하고 "저장되었습니다"라고 안내했으나
    ' 실제 파일(.xlsm)은 저장하지 않아, 사용자가 이 메시지를 믿고 파일을 닫으면(저장 여부를 묻는
    ' 창에서 "아니요"를 누르거나 강제 종료되면) 방금 입력한 데이터가 통째로 사라지는 것이 "저장
    ' 후 불러오기 안 됨" 보고의 실제 원인으로 추정됨. 버튼을 누르는 즉시 파일도 저장해 메시지와
    ' 실제 동작을 일치시킴. Workbook_BeforeSave가 저장 경계를 재검증하며 실패 시 Cancel=True로
    ' 저장을 취소하므로, 그 결과(ThisWorkbook.Saved)를 확인해 정확한 안내로 구분함.
    ThisWorkbook.Save
    If showMessage Then
        If ThisWorkbook.Saved Then
            MsgBox "저장되었습니다. (순번 " & seq & ")", vbInformation
        Else
            MsgBox "DB에는 기록됐으나 파일 저장이 취소되었습니다(저장 경계 재검증 실패). 데이터를 다시 확인한 뒤 저장하세요.", vbExclamation
        End If
    End If
End Sub

Sub 수정하기()
    Call 수정하기_실행(True)
End Sub

Public Sub 검증_수정하기()
    Call 수정하기_실행(False)
End Sub

Private Sub 수정하기_실행(ByVal showMessage As Boolean)
    ' 2026-09-24: 저장하기_실행과 동일한 사용자 결정(경고만, 저장은 항상 진행)을 적용함.
    Call 필수값검증()
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    Call 품목행검증(wsIn, showMessage)
    Call 업체행검증(wsIn, showMessage)
    Call 업체대표자행검증(wsIn, showMessage)
    Call 정량평가행검증(wsIn, showMessage)
    Call 정성평가행검증(wsIn, showMessage)
    Call 자기평점행검증(wsIn, showMessage)
    Call 위원행검증(wsIn, showMessage)
    Call 평가행검증(wsIn, showMessage)
    If showMessage And 검증_업체중복경고() Then MsgBox "중복된 업체명이 있습니다. 실제 동일 업체인지 확인한 뒤 수정하세요.", vbExclamation
    If showMessage And 검증_위원중복경고() Then MsgBox "중복된 위원 역할·직위가 있습니다. 실제 구성과 역할을 확인하세요.", vbExclamation

    Dim seq As Variant
    seq = wsIn.Range("H1").Value
    If seq = "" Then
        MsgBox "먼저 [불러오기]로 수정할 레코드를 선택하거나, 신규 레코드는 [저장하기]를 사용하세요.", vbExclamation
        Exit Sub
    End If
    Dim r As Long
    r = CLng(seq) + 1
    If wsDB.Cells(r, 1).Value <> CLng(seq) Then
        MsgBox "DB에서 해당 레코드를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    Call 필드복사_기초자료_DB(wsIn, wsDB, r)
    Call 품목복사_기초자료_DB(wsIn, wsItems, CLng(seq))
    Call 업체복사_기초자료_DB(wsIn, wsVendors, CLng(seq))
    Call 정량평가복사_기초자료_DB(wsIn, wsQuant, CLng(seq))
    Call 정성평가복사_기초자료_DB(wsIn, wsQual, CLng(seq))
    Call 자기평점복사_기초자료_DB(wsIn, wsSelf, CLng(seq))
    Call 위원복사_기초자료_DB(wsIn, wsCommittee, CLng(seq))
    Call 평가복사_기초자료_DB(wsIn, wsScore, CLng(seq))
    ' 2026-09-23: 저장하기_실행과 동일한 이유로 수정 직후 파일도 저장함(아래 참고).
    ThisWorkbook.Save
    If showMessage Then
        If ThisWorkbook.Saved Then
            MsgBox "수정되었습니다. (순번 " & seq & ")", vbInformation
        Else
            MsgBox "DB에는 기록됐으나 파일 저장이 취소되었습니다(저장 경계 재검증 실패). 데이터를 다시 확인한 뒤 저장하세요.", vbExclamation
        End If
    End If
End Sub

' 2026-09-21: InputBox 직접 입력 대신 기초자료입력!I2 드롭다운(라벨 "순번 N: 학교명 / ...")
' 선택값에서 순번만 추출함. 드롭다운 목록은 DB 시트 U열·이름정의 "DB불러오기목록" 참고.
' 파싱만 분리해 헤드리스 자동 검증에서 MsgBox 없이 형식 검증을 할 수 있게 함.
Private Function 불러오기_드롭다운파싱(ByVal selectedLabel As String) As Long
    불러오기_드롭다운파싱 = 0
    If Left(selectedLabel, 3) <> "순번 " Then Exit Function
    Dim afterPrefix As String, colonPos As Long, seqStr As String
    afterPrefix = Mid(selectedLabel, 4)
    colonPos = InStr(afterPrefix, ":")
    If colonPos = 0 Then Exit Function
    seqStr = Trim(Left(afterPrefix, colonPos - 1))
    If Not IsNumeric(seqStr) Then Exit Function
    불러오기_드롭다운파싱 = CLng(seqStr)
End Function

Public Function 검증_불러오기_드롭다운파싱(ByVal selectedLabel As String) As Long
    검증_불러오기_드롭다운파싱 = 불러오기_드롭다운파싱(selectedLabel)
End Function

' 2026-09-24 사용자 요청: 상단 드롭다운을 먼저 선택해야 하는 방식 대신, 버튼을 누르면
' 저장된 레코드 목록이 팝업(frmLoadPicker)으로 바로 떠서 그 자리에서 고를 수 있게 함.
Sub 불러오기()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmLoadPicker")
    Call frm.초기화
    frm.Show
End Sub

Public Sub 검증_첫레코드불러오기()
    Call 불러오기_순번(1, False)
End Sub

' frmLoadPicker 팝업(2026-09-24 사용자 요청)이 순번만 알고 있는 상태에서 이 Private Sub를
' 부를 수 있게 하는 공개 진입점. 실제 로직은 그대로 불러오기_순번 하나만 사용한다.
Public Sub 불러오기_레코드선택(ByVal seq As Long)
    Call 불러오기_순번(seq, True)
End Sub

Private Sub 불러오기_순번(ByVal seq As Long, ByVal showMessage As Boolean)
    Dim wsIn As Worksheet, wsDB As Worksheet, wsItems As Worksheet, wsVendors As Worksheet, wsCommittee As Worksheet, wsScore As Worksheet, wsQuant As Worksheet, wsQual As Worksheet, wsSelf As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsDB = Sheets("DB")
    Set wsItems = Sheets("DB_품목")
    Set wsVendors = Sheets("DB_업체")
    Set wsCommittee = Sheets("DB_위원")
    Set wsScore = Sheets("DB_평가")
    Set wsQuant = Sheets("DB_정량평가")
    Set wsQual = Sheets("DB_정성평가")
    Set wsSelf = Sheets("DB_자기평점")
    Dim r As Long
    r = seq + 1
    If wsDB.Cells(r, 1).Value <> seq Then
        If showMessage Then MsgBox "해당 순번의 레코드를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    wsIn.Range("C4").Value = wsDB.Cells(r, 2).Value
    wsIn.Range("C5").Value = wsDB.Cells(r, 3).Value
    wsIn.Range("C6").Value = wsDB.Cells(r, 4).Value
    wsIn.Range("C7").Value = wsDB.Cells(r, 5).Value
    wsIn.Range("C8").Value = wsDB.Cells(r, 6).Value
    wsIn.Range("C9").Value = wsDB.Cells(r, 7).Value
    wsIn.Range("C12").Value = wsDB.Cells(r, 8).Value
    wsIn.Range("C13").Value = wsDB.Cells(r, 9).Value
    wsIn.Range("C14").Value = wsDB.Cells(r, 10).Value
    wsIn.Range("C15").Value = wsDB.Cells(r, 11).Value
    wsIn.Range("C16").Value = wsDB.Cells(r, 12).Value
    wsIn.Range("C17").Value = wsDB.Cells(r, 13).Value
    wsIn.Range("C18").Value = wsDB.Cells(r, 14).Value
    wsIn.Range("C21").Value = wsDB.Cells(r, 15).Value
    wsIn.Range("C22").Value = wsDB.Cells(r, 16).Value
    wsIn.Range("C23").Value = wsDB.Cells(r, 17).Value
    wsIn.Range("C24").Value = wsDB.Cells(r, 18).Value
    wsIn.Range("C25").Value = wsDB.Cells(r, 19).Value
    wsIn.Range("C10").Value = wsDB.Cells(r, 22).Value  ' C-08 담당자 성명 (2026-09-23 신설, V열)
    wsIn.Range("F10").Value = wsDB.Cells(r, 23).Value  ' C-07 교장 성명 (2026-09-23 신설, W열)
    wsIn.Range("B29:D38").ClearContents
    Dim itemRow As Long, inputRow As Long
    For itemRow = 2 To wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row
        If wsItems.Cells(itemRow, 1).Value = seq Then
            inputRow = 28 + CLng(wsItems.Cells(itemRow, 2).Value)
            If inputRow >= 29 And inputRow <= 38 Then
                wsIn.Cells(inputRow, 2).Value = wsItems.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsItems.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 4).Value = wsItems.Cells(itemRow, 5).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B58:C67").ClearContents
    For itemRow = 2 To wsCommittee.Cells(wsCommittee.Rows.Count, 1).End(xlUp).Row
        If wsCommittee.Cells(itemRow, 1).Value = seq Then
            inputRow = 57 + CLng(wsCommittee.Cells(itemRow, 2).Value)
            If inputRow >= 58 And inputRow <= 67 Then
                wsIn.Cells(inputRow, 2).Value = wsCommittee.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsCommittee.Cells(itemRow, 4).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B72:D81").ClearContents
    For itemRow = 2 To wsScore.Cells(wsScore.Rows.Count, 1).End(xlUp).Row
        If wsScore.Cells(itemRow, 1).Value = seq Then
            inputRow = 71 + CLng(wsScore.Cells(itemRow, 2).Value)
            If inputRow >= 72 And inputRow <= 81 Then
                wsIn.Cells(inputRow, 2).Value = wsScore.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 3).Value = wsScore.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 4).Value = wsScore.Cells(itemRow, 5).Value
            End If
        End If
    Next itemRow
    wsIn.Range("B44:F53").ClearContents
    wsIn.Range("H44:L53").ClearContents
    wsIn.Range("N44:Q53").ClearContents
    wsIn.Range("S44:S53,X44:X53").ClearContents
    For itemRow = 2 To wsVendors.Cells(wsVendors.Rows.Count, 1).End(xlUp).Row
        If wsVendors.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsVendors.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 2).Value = wsVendors.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 19).Value = wsVendors.Cells(itemRow, 4).Value  ' R-08 대표자 성명 (2026-09-23 신설)
                wsIn.Cells(inputRow, 24).Value = wsVendors.Cells(itemRow, 5).Value  ' R-09 대표자 연락처 (2026-09-23 신설)
            End If
        End If
    Next itemRow
    For itemRow = 2 To wsSelf.Cells(wsSelf.Rows.Count, 1).End(xlUp).Row
        If wsSelf.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsSelf.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 14).Value = wsSelf.Cells(itemRow, 3).Value: wsIn.Cells(inputRow, 15).Value = wsSelf.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 16).Value = wsSelf.Cells(itemRow, 5).Value: wsIn.Cells(inputRow, 17).Value = wsSelf.Cells(itemRow, 6).Value
            End If
        End If
    Next itemRow
    For itemRow = 2 To wsQual.Cells(wsQual.Rows.Count, 1).End(xlUp).Row
        If wsQual.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsQual.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 8).Value = 정성점수등급(wsQual.Cells(itemRow, 3).Value, 15)
                wsIn.Cells(inputRow, 9).Value = 정성점수등급(wsQual.Cells(itemRow, 4).Value, 10)
                wsIn.Cells(inputRow, 10).Value = 정성점수등급(wsQual.Cells(itemRow, 5).Value, 15)
                wsIn.Cells(inputRow, 11).Value = 정성점수등급(wsQual.Cells(itemRow, 6).Value, 10)
                wsIn.Cells(inputRow, 12).Value = wsQual.Cells(itemRow, 7).Value
            End If
        End If
    Next itemRow
    For itemRow = 2 To wsQuant.Cells(wsQuant.Rows.Count, 1).End(xlUp).Row
        If wsQuant.Cells(itemRow, 1).Value = seq Then
            inputRow = 43 + CLng(wsQuant.Cells(itemRow, 2).Value)
            If inputRow >= 44 And inputRow <= 53 Then
                wsIn.Cells(inputRow, 3).Value = wsQuant.Cells(itemRow, 3).Value
                wsIn.Cells(inputRow, 4).Value = wsQuant.Cells(itemRow, 4).Value
                wsIn.Cells(inputRow, 5).Value = wsQuant.Cells(itemRow, 5).Value
                wsIn.Cells(inputRow, 6).Value = wsQuant.Cells(itemRow, 6).Value
            End If
        End If
    Next itemRow
    wsIn.Range("H1").Value = seq
    If showMessage Then MsgBox "불러왔습니다. (순번 " & seq & ")", vbInformation
End Sub
'@

    $codeInput = $codeInput -replace "`r`n", "`r" -replace "`n", "`r"
    $modInput.CodeModule.AddFromString($codeInput)
    L "Module_기초자료 추가 완료 (줄 수: $($modInput.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT1_SAVE_OK" } catch { L "DIAG_CHECKPOINT1_SAVE_FAIL: $($_.Exception.Message)" }

    $modOutput = $vbproj.VBComponents.Add(1)
    $modOutput.Name = "Module_출력"
    $codeOutput = @'
Option Explicit

' 2026-09-19 진단: 이 프로시저가 Form ID마다 개별 skippedInvalidFxxx 카운터 변수와
' 반복되는 "wsSel.Cells(r,2).Value = "F-XXX"" 비교, 그리고 그 변수들을 전부 나열한
' 초장문 skipDetail 연결식·OR 조건식을 갖고 있어, Form ID를 추가할 때마다 이 프로시저
' 하나의 컴파일 크기가 커지다가 VBA "프로시저가 너무 큼" 한계를 실제로 넘어 저장(컴파일)
' 자체가 실패하는 사고가 재현됨(이분 탐색으로 확정). 재발 방지를 위해 카운터를 Dictionary
' 하나로, "단순 단일 출력" 그룹(업체 반복 없이 검증 함수 1개만 통과하면 서식선택_출력의
' F열 시트를 그대로 출력하는 F-001~006·014·024·032~043)의 20개 ElseIf 문자열 비교를
' Select Case 조회 함수(단일출력검증함수명)로 옮겨 이 프로시저 자체의 크기를 줄인다.
' 업체 반복 출력 그룹(F-015~023·025)의 동작 로직은 전혀 바꾸지 않았다(캐시된 fid 변수
' 사용과 IncSkip 호출로만 치환).
Private Sub IncSkip(ByRef d As Object, ByVal key As String)
    If d.Exists(key) Then
        d(key) = d(key) + 1
    Else
        d.Add key, 1
    End If
End Sub

Private Function 단일출력검증함수명(ByVal fid As String) As String
    Select Case fid
        Case "F-001": 단일출력검증함수명 = "검증_F001출력가능"
        Case "F-002": 단일출력검증함수명 = "검증_F002출력가능"
        Case "F-003": 단일출력검증함수명 = "검증_F003출력가능"
        Case "F-004": 단일출력검증함수명 = "검증_F004출력가능"
        Case "F-005": 단일출력검증함수명 = "검증_F005출력가능"
        Case "F-006": 단일출력검증함수명 = "검증_F006출력가능"
        Case "F-014": 단일출력검증함수명 = "검증_F014출력가능"
        Case "F-009": 단일출력검증함수명 = "검증_F009출력가능"
        Case "F-010": 단일출력검증함수명 = "검증_F010출력가능"
        Case "F-011": 단일출력검증함수명 = "검증_F011출력가능"
        Case "F-012": 단일출력검증함수명 = "검증_F012출력가능"
        Case "F-008": 단일출력검증함수명 = "검증_F008출력가능"
        Case "F-024": 단일출력검증함수명 = "검증_F024출력가능"
        Case "F-032": 단일출력검증함수명 = "검증_F032출력가능"
        Case "F-033": 단일출력검증함수명 = "검증_F033출력가능"
        Case "F-034": 단일출력검증함수명 = "검증_F034출력가능"
        Case "F-035": 단일출력검증함수명 = "검증_F035출력가능"
        Case "F-036": 단일출력검증함수명 = "검증_F036출력가능"
        Case "F-037": 단일출력검증함수명 = "검증_F037출력가능"
        Case "F-038": 단일출력검증함수명 = "검증_F038출력가능"
        Case "F-039": 단일출력검증함수명 = "검증_F039출력가능"
        Case "F-040": 단일출력검증함수명 = "검증_F040출력가능"
        Case "F-041": 단일출력검증함수명 = "검증_F041출력가능"
        Case "F-042": 단일출력검증함수명 = "검증_F042출력가능"
        Case "F-043": 단일출력검증함수명 = "검증_F043출력가능"
        Case "F-044": 단일출력검증함수명 = "검증_F044출력가능"
        Case "F-045": 단일출력검증함수명 = "검증_F045출력가능"
        Case "F-046": 단일출력검증함수명 = "검증_F046출력가능"
        Case "F-047": 단일출력검증함수명 = "검증_F047출력가능"
        Case "F-048": 단일출력검증함수명 = "검증_F048출력가능"
        Case "F-049": 단일출력검증함수명 = "검증_F049출력가능"
        Case "F-050": 단일출력검증함수명 = "검증_F050출력가능"
        Case "F-027": 단일출력검증함수명 = "검증_F027출력가능"
        Case "F-029": 단일출력검증함수명 = "검증_F029출력가능"
        Case "F-030": 단일출력검증함수명 = "검증_F030출력가능"
        Case "F-031": 단일출력검증함수명 = "검증_F031출력가능"
        Case Else: 단일출력검증함수명 = ""
    End Select
End Function

' Application.Run의 문자열 기반 동적 디스패치는, 이미 외부 Automation Application.Run으로
' 실행 중인 프로시저 내부에서 다시 호출될 때 재진입 이슈로 이어질 수 있어(2026-09-20
' 확인), 문자열 이름 대신 각 함수를 직접 호출하는 명시적 분기로 대체한다.
Private Function 단일출력검증실행(ByVal fid As String) As Boolean
    Select Case fid
        Case "F-001": 단일출력검증실행 = 검증_F001출력가능()
        Case "F-002": 단일출력검증실행 = 검증_F002출력가능()
        Case "F-003": 단일출력검증실행 = 검증_F003출력가능()
        Case "F-004": 단일출력검증실행 = 검증_F004출력가능()
        Case "F-005": 단일출력검증실행 = 검증_F005출력가능()
        Case "F-006": 단일출력검증실행 = 검증_F006출력가능()
        Case "F-014": 단일출력검증실행 = 검증_F014출력가능()
        Case "F-009": 단일출력검증실행 = 검증_F009출력가능()
        Case "F-010": 단일출력검증실행 = 검증_F010출력가능()
        Case "F-011": 단일출력검증실행 = 검증_F011출력가능()
        Case "F-012": 단일출력검증실행 = 검증_F012출력가능()
        Case "F-008": 단일출력검증실행 = 검증_F008출력가능()
        Case "F-024": 단일출력검증실행 = 검증_F024출력가능()
        Case "F-032": 단일출력검증실행 = 검증_F032출력가능()
        Case "F-033": 단일출력검증실행 = 검증_F033출력가능()
        Case "F-034": 단일출력검증실행 = 검증_F034출력가능()
        Case "F-035": 단일출력검증실행 = 검증_F035출력가능()
        Case "F-036": 단일출력검증실행 = 검증_F036출력가능()
        Case "F-037": 단일출력검증실행 = 검증_F037출력가능()
        Case "F-038": 단일출력검증실행 = 검증_F038출력가능()
        Case "F-039": 단일출력검증실행 = 검증_F039출력가능()
        Case "F-040": 단일출력검증실행 = 검증_F040출력가능()
        Case "F-041": 단일출력검증실행 = 검증_F041출력가능()
        Case "F-042": 단일출력검증실행 = 검증_F042출력가능()
        Case "F-043": 단일출력검증실행 = 검증_F043출력가능()
        Case "F-044": 단일출력검증실행 = 검증_F044출력가능()
        Case "F-045": 단일출력검증실행 = 검증_F045출력가능()
        Case "F-046": 단일출력검증실행 = 검증_F046출력가능()
        Case "F-047": 단일출력검증실행 = 검증_F047출력가능()
        Case "F-048": 단일출력검증실행 = 검증_F048출력가능()
        Case "F-049": 단일출력검증실행 = 검증_F049출력가능()
        Case "F-050": 단일출력검증실행 = 검증_F050출력가능()
        Case "F-027": 단일출력검증실행 = 검증_F027출력가능()
        Case "F-029": 단일출력검증실행 = 검증_F029출력가능()
        Case "F-030": 단일출력검증실행 = 검증_F030출력가능()
        Case "F-031": 단일출력검증실행 = 검증_F031출력가능()
        Case Else: 단일출력검증실행 = False
    End Select
End Function

Private Function 선택된시트목록(ByRef outNames() As String, ByVal showMessage As Boolean) As Long
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row

    Dim tmp() As String
    ReDim tmp(1 To lastRow + 50)   ' 업체별 F-015~F-020 출력에 대비해 여유를 둠
    Dim cnt As Long
    cnt = 0
    Dim skipCounts As Object
    Set skipCounts = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 5 To lastRow
        If wsSel.Cells(r, 1).Value = True Then
            Dim status As String, fid As String
            status = wsSel.Cells(r, 5).Value
            fid = wsSel.Cells(r, 2).Value

            Dim simpleValidator As String
            simpleValidator = 단일출력검증함수명(fid)

            If fid = "F-050" Then
                Dim f050Reason As String
                If 검증_필수값검증() And F050원자료없음(f050Reason) Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    IncSkip skipCounts, fid
                End If
            ElseIf simpleValidator <> "" Then
                ' 미리보기·PDF는 기초자료가 비어 있어도 빈 양식을 확인할 수 있게 한다.
                ' 저장 검증과 F-050 개인정보 원자료 방지는 별도 경계로 그대로 유지한다.
                cnt = cnt + 1
                tmp(cnt) = wsSel.Cells(r, 6).Value
            ElseIf fid = "F-015" Then
                If Not 검증_F015출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim vendorRow As Long, addedAny As Boolean
                    addedAny = False
                    For vendorRow = 44 To 53
                        If F015업체행완전한가(vendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F015임시시트생성(vendorRow - 43)
                            addedAny = True
                        End If
                    Next vendorRow
                    If Not addedAny Then IncSkip skipCounts, fid
                End If
            ElseIf fid = "F-016" Then
                If Not 검증_F016출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim qualitativeVendorRow As Long, addedQualitative As Boolean
                    addedQualitative = False
                    For qualitativeVendorRow = 44 To 53
                        If F016업체행완전한가(qualitativeVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F016임시시트생성(qualitativeVendorRow - 43)
                            addedQualitative = True
                        End If
                    Next qualitativeVendorRow
                    If Not addedQualitative Then IncSkip skipCounts, fid
                End If
            ElseIf fid = "F-017" Then
                If Not 검증_F017출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim confirmationVendorRow As Long
                    For confirmationVendorRow = 44 To 53
                        If F017업체행완전한가(confirmationVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F017임시시트생성(confirmationVendorRow - 43)
                        End If
                    Next confirmationVendorRow
                End If
            ElseIf fid = "F-018" Then
                If Not 검증_F018출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim selfVendorRow As Long
                    For selfVendorRow = 44 To 53
                        If F018업체행완전한가(selfVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F018임시시트생성(selfVendorRow - 43)
                        End If
                    Next selfVendorRow
                End If
            ElseIf fid = "F-019" Then
                If Not 검증_F019출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim applicationVendorRow As Long
                    For applicationVendorRow = 44 To 53
                        If F019업체행완전한가(applicationVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F019임시시트생성(applicationVendorRow - 43)
                        End If
                    Next applicationVendorRow
                End If
            ElseIf fid = "F-020" Then
                If Not 검증_F020출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim reportVendorRow As Long
                    For reportVendorRow = 44 To 53
                        If F020업체행완전한가(reportVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F020임시시트생성(reportVendorRow - 43)
                        End If
                    Next reportVendorRow
                End If
            ElseIf fid = "F-021" Then
                If Not 검증_F021출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim proposalVendorRow As Long
                    For proposalVendorRow = 44 To 53
                        If F021업체행완전한가(proposalVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F021임시시트생성(proposalVendorRow - 43)
                        End If
                    Next proposalVendorRow
                End If
            ElseIf fid = "F-022" Then
                If Not 검증_F022출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim performanceVendorRow As Long
                    For performanceVendorRow = 44 To 53
                        If F022업체행완전한가(performanceVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F022임시시트생성(performanceVendorRow - 43)
                        End If
                    Next performanceVendorRow
                End If
            ElseIf fid = "F-023" Then
                If Not 검증_F023출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim specVendorRow As Long
                    For specVendorRow = 44 To 53
                        If F023업체행완전한가(specVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F023임시시트생성(specVendorRow - 43)
                        End If
                    Next specVendorRow
                End If
            ElseIf fid = "F-025" Then
                If Not 검증_F025출력가능() Then
                    cnt = cnt + 1
                    tmp(cnt) = wsSel.Cells(r, 6).Value
                Else
                    Dim asVendorRow As Long
                    For asVendorRow = 44 To 53
                        If F025업체행완전한가(asVendorRow) Then
                            cnt = cnt + 1
                            tmp(cnt) = F025임시시트생성(asVendorRow - 43)
                        End If
                    Next asVendorRow
                End If
            ElseIf status = "Y" Then
                cnt = cnt + 1
                tmp(cnt) = wsSel.Cells(r, 6).Value
            ElseIf status = "D" Then
                IncSkip skipCounts, "HWPX 우선순위 위임"
            Else
                IncSkip skipCounts, "미구현"
            End If
        End If
    Next r

    Dim skipDetail As String
    skipDetail = ""
    Dim skipKey As Variant
    For Each skipKey In skipCounts.Keys
        skipDetail = skipDetail & skipKey & " 제외: " & skipCounts(skipKey) & "건, "
    Next skipKey
    If Len(skipDetail) >= 2 Then skipDetail = Left(skipDetail, Len(skipDetail) - 2)

    If cnt = 0 Then
        If showMessage Then MsgBox "구현된 서식이 선택되지 않았습니다." & vbCrLf & skipDetail, vbExclamation
        선택된시트목록 = 0
        Exit Function
    End If

    If skipCounts.Count > 0 And showMessage Then
        MsgBox "일부 선택 서식은 아직 준비되지 않아 제외합니다." & vbCrLf & skipDetail, vbInformation
    End If

    ReDim outNames(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        outNames(i) = tmp(i)
    Next i
    선택된시트목록 = cnt
End Function

Public Function 검증_F024출력가능() As Boolean
    Dim itemCount As Long
    itemCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B29:B38"))
    검증_F024출력가능 = 검증_품목행검증() And itemCount >= 1 And itemCount <= 6
End Function

Public Function 검증_F014출력가능() As Boolean
    Dim scoreCount As Long
    scoreCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B72:B81"))
    검증_F014출력가능 = 검증_필수값검증() And 검증_평가행검증() And scoreCount >= 1 And scoreCount <= 10
End Function

Public Function 검증_F015출력가능() As Boolean
    Dim vendorRow As Long, hasCompleteVendor As Boolean
    If Not 검증_정량평가행검증() Then Exit Function
    hasCompleteVendor = False
    For vendorRow = 44 To 53
        If F015업체행완전한가(vendorRow) Then
            hasCompleteVendor = True
            Exit For
        End If
    Next vendorRow
    검증_F015출력가능 = 검증_필수값검증() And hasCompleteVendor
End Function

Public Function 검증_F016출력가능() As Boolean
    Dim vendorRow As Long, hasCompleteVendor As Boolean
    If Not 검증_정성평가행검증() Then Exit Function
    For vendorRow = 44 To 53
        If F016업체행완전한가(vendorRow) Then
            hasCompleteVendor = True
            Exit For
        End If
    Next vendorRow
    검증_F016출력가능 = 검증_필수값검증() And hasCompleteVendor
End Function

Public Function 검증_F017출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F017업체행완전한가(vendorRow) Then
            검증_F017출력가능 = True
            Exit Function
        End If
    Next vendorRow
End Function

Public Function 검증_F018출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Or Not 검증_자기평점행검증() Then Exit Function
    For vendorRow = 44 To 53
        If F018업체행완전한가(vendorRow) Then 검증_F018출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F019출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F019업체행완전한가(vendorRow) Then 검증_F019출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F020출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F020업체행완전한가(vendorRow) Then 검증_F020출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F021출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F021업체행완전한가(vendorRow) Then 검증_F021출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F022출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F022업체행완전한가(vendorRow) Then 검증_F022출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F023출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F023업체행완전한가(vendorRow) Then 검증_F023출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F025출력가능() As Boolean
    Dim vendorRow As Long
    If Not 검증_필수값검증() Then Exit Function
    For vendorRow = 44 To 53
        If F025업체행완전한가(vendorRow) Then 검증_F025출력가능 = True: Exit Function
    Next vendorRow
End Function

Public Function 검증_F001출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F001출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

Public Function 검증_F002출력가능() As Boolean
    검증_F002출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F003출력가능() As Boolean
    검증_F003출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F004출력가능() As Boolean
    검증_F004출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F005출력가능() As Boolean
    검증_F005출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F006출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F006출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

' F-041~F-043은 F-017~F-025처럼 업체마다 별도 문서를 만드는 것이 아니라, 낙찰자·계약상대자는
' 언제나 1개 업체이므로 각 시트의 I3 선택 순번이 가리키는 업체행 1곳만 완전한지(상호명이 비어있지
' 않은지) 확인한다. Module_출력에서 업체 반복 루프를 돌리지 않고 F-001/F-006과 같은 단일 출력으로
' 취급하는 이유가 이 검증 함수에도 그대로 반영된다.
Private Function 낙찰업체선택완전한가(ByVal sheetName As String) As Boolean
    Dim wsForm As Worksheet, vendorRow As Long
    Set wsForm = Sheets(sheetName)
    vendorRow = 43 + CLng(Val(wsForm.Range("I3").Value))
    If vendorRow < 44 Or vendorRow > 53 Then Exit Function
    낙찰업체선택완전한가 = Trim(Sheets("기초자료입력").Cells(vendorRow, 2).Value & "") <> ""
End Function

Public Function 검증_F041출력가능() As Boolean
    With Sheets("F-041_낙찰자결정")
        검증_F041출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-041_낙찰자결정") And _
            양수값(.Range("C17").Value) And 양수값(.Range("E21").Value) And 양수값(.Range("F21").Value)
    End With
End Function

Public Function 검증_F042출력가능() As Boolean
    With Sheets("F-042_낙찰자결정통보")
        검증_F042출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-042_낙찰자결정통보") And _
            양수값(.Range("D17").Value) And 양수값(.Range("E17").Value) And 양수값(.Range("F17").Value) And _
            계약금액일치(.Range("D17").Value, .Range("E17").Value, .Range("F17").Value) And _
            Trim(Sheets("기초자료입력").Range("C18").Value & "") <> ""
    End With
End Function

Public Function 검증_F043출력가능() As Boolean
    With Sheets("F-043_계약체결")
        검증_F043출력가능 = 검증_필수값검증() And 낙찰업체선택완전한가("F-043_계약체결") And _
            양수값(.Range("C16").Value) And 양수값(Sheets("기초자료입력").Range("C16").Value) And _
            Trim(Sheets("기초자료입력").Range("C17").Value & "") <> ""
    End With
End Function

Public Function 검증_F044출력가능() As Boolean
    검증_F044출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F045출력가능() As Boolean
    검증_F045출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F046출력가능() As Boolean
    검증_F046출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F048출력가능() As Boolean
    검증_F048출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F047출력가능() As Boolean
    검증_F047출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F049출력가능() As Boolean
    검증_F049출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F050출력가능() As Boolean
    검증_F050출력가능 = 검증_필수값검증()
End Function

' 계획.md 8.1절 후순위 보류 4종(F-027·F-029·F-030·F-031). 넷 다 업체 반복행 없이 단일 출력이며
' 학교 공통값만 필요하므로 검증_필수값검증()만 요구한다. F-031만 F-043의 낙찰단가(C16)가 숫자로
' 채워져 있어야 계약금액이 의미 있는 값을 갖게 되므로 추가 조건을 둔다(비어 있으면 0원으로
' 표시되는 것을 방지 — F-008에서 실제로 겪은 "빈 셀이 0으로 표시" 결함과 같은 유형).
Public Function 검증_F027출력가능() As Boolean
    검증_F027출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F029출력가능() As Boolean
    검증_F029출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F030출력가능() As Boolean
    검증_F030출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F031출력가능() As Boolean
    ' VBA의 IsNumeric(Empty)는 True를 반환하는 알려진 함정이라(완전히 빈 셀의 .Value는 Empty),
    ' 이것만으로는 "낙찰단가 미입력"을 걸러내지 못한다. IsEmpty를 먼저 확인해 실제로 값이 있는
    ' 숫자 셀만 허용한다.
    Dim unitPriceCell As Variant
    unitPriceCell = Sheets("F-043_계약체결").Range("C16").Value
    검증_F031출력가능 = 검증_필수값검증() And (Not IsEmpty(unitPriceCell)) And IsNumeric(unitPriceCell)
End Function

Public Function 검증_F009출력가능() As Boolean
    With Sheets("기초자료입력")
        검증_F009출력가능 = 검증_필수값검증() And 양수값(.Range("C15").Value) And _
            Trim(.Range("C17").Value & "") <> "" And Trim(.Range("C14").Value & "") <> ""
    End With
End Function

Public Function 검증_F010출력가능() As Boolean
    With Sheets("기초자료입력")
        검증_F010출력가능 = 검증_필수값검증() And 양수값(.Range("C15").Value) And _
            양수값(.Range("C16").Value) And Trim(.Range("C18").Value & "") <> ""
    End With
End Function

Public Function 검증_F011출력가능() As Boolean
    With Sheets("기초자료입력")
        검증_F011출력가능 = 검증_필수값검증() And 양수값(.Range("C15").Value) And _
            양수값(.Range("C16").Value) And Trim(.Range("C17").Value & "") <> "" And _
            Trim(.Range("C18").Value & "") <> ""
    End With
End Function

Public Function 검증_F012출력가능() As Boolean
    검증_F012출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F008출력가능() As Boolean
    Dim itemCount08 As Long
    itemCount08 = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B29:B34"))
    검증_F008출력가능 = 검증_필수값검증() And itemCount08 >= 1 And itemCount08 <= 6
End Function

Private Function 양수값(ByVal value As Variant) As Boolean
    ' VBA의 And는 단축평가를 하지 않아 IsNumeric(value)가 False여도 CDbl(value)가 그대로
    ' 실행됨 — 빈 칸(공백 문자열)처럼 숫자로 변환 불가능한 값에서 형식 불일치(오류 13)가
    ' 발생하는 것을 실사용 중 발견해(2026-09-21) If로 분리함.
    If Not IsNumeric(value) Then Exit Function
    양수값 = CDbl(value) > 0
End Function

Private Function 계약금액일치(ByVal unitPrice As Variant, ByVal quantity As Variant, ByVal contractAmount As Variant) As Boolean
    If Not 양수값(unitPrice) Or Not 양수값(quantity) Or Not 양수값(contractAmount) Then Exit Function
    계약금액일치 = Abs(CDbl(contractAmount) - WorksheetFunction.Round(CDbl(unitPrice) * CDbl(quantity), 0)) < 0.000001
End Function

Public Function 검증_F037출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F037출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

Public Function 검증_F038출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F038출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

' F-039는 F-041~F-043과 같은 I3 단일 업체 선택 패턴을 재사용한다(업체 반복 루프 없음).
' 위원별 원점수(옷감·완성도·A/S·하자)는 이 문서 전용 입력칸(E8:H15)이며 다른 서식과
' 공유하는 DB가 아니다. 위원 성명(D열)은 청탁방지·익명성 원칙에 따라 자동 반영하지 않고
' 자필 공란으로 유지하므로 검증 대상에서 제외한다. 총계·평균은 원문의 "위원별 평가 점수
' 중 최고점 및 최저점을 제외" 안내에 따라 8명 점수 중 최댓값·최솟값을 뺀 뒤 합산·평균한다.
Public Function 검증_F039출력가능() As Boolean
    With Sheets("F-039_업체별제안서평가표")
        Dim ok As Boolean
        ok = 검증_필수값검증() And 낙찰업체선택완전한가("F-039_업체별제안서평가표")
        If ok Then
            ok = Application.WorksheetFunction.Count(.Range("E8:H15")) = 32 And _
                Application.WorksheetFunction.Min(.Range("E8:E15")) >= 0 And Application.WorksheetFunction.Max(.Range("E8:E15")) <= 15 And _
                Application.WorksheetFunction.Min(.Range("F8:F15")) >= 0 And Application.WorksheetFunction.Max(.Range("F8:F15")) <= 10 And _
                Application.WorksheetFunction.Min(.Range("G8:G15")) >= 0 And Application.WorksheetFunction.Max(.Range("G8:G15")) <= 15 And _
                Application.WorksheetFunction.Min(.Range("H8:H15")) >= 0 And Application.WorksheetFunction.Max(.Range("H8:H15")) <= 10
        End If
        검증_F039출력가능 = ok
    End With
End Function

Public Function 검증_F040출력가능() As Boolean
    검증_F040출력가능 = 검증_필수값검증()
End Function

Public Function 검증_F032출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F032출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F033출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F033출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F034출력가능() As Boolean
    Dim committeeCount As Long
    committeeCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B58:B67"))
    검증_F034출력가능 = 검증_필수값검증() And 검증_위원행검증() And committeeCount >= 1
End Function

' F-035·F-036은 F-032·F-033과 같이 업체 반복행(기초자료입력!B44:B53) 전체를 한 표에
' 보여주는 집계 보고서이며, 점수는 F-015/F-016이 이미 쓰는 학교 평가 점수 열(C:F 정량,
' H:L 정성)을 그대로 재조회한다. 업체 존재만 확인하고 점수 입력 여부는 강제하지 않는다
' (F-032/F-033과 동일 수준 — 평가 진행 중에도 접수 현황·중간 보고서를 출력할 수 있어야 함).
Public Function 검증_F035출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F035출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Public Function 검증_F036출력가능() As Boolean
    Dim vendorCount As Long
    vendorCount = Application.WorksheetFunction.CountA(Sheets("기초자료입력").Range("B44:B53"))
    검증_F036출력가능 = 검증_필수값검증() And 검증_업체행검증() And vendorCount >= 1
End Function

Private Function F015업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F015업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 3).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 4).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 5).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 6).Value & "") <> "" And _
        정수범위(wsIn.Cells(sourceRow, 3).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 4).Value, 0, 10) And _
        정수범위(wsIn.Cells(sourceRow, 5).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 6).Value, 0, 15)
End Function

Private Function F016업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F016업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 8).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 9).Value & "") <> "" And _
        Trim(wsIn.Cells(sourceRow, 10).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 11).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 12).Value & "") <> "" And _
        정성등급유효(wsIn.Cells(sourceRow, 8).Value) And 정성등급유효(wsIn.Cells(sourceRow, 9).Value) And _
        정성등급유효(wsIn.Cells(sourceRow, 10).Value) And 정성등급유효(wsIn.Cells(sourceRow, 11).Value) And _
        정수범위(wsIn.Cells(sourceRow, 12).Value, -15, 5)
End Function

Private Function F017업체행완전한가(ByVal sourceRow As Long) As Boolean
    F017업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F018업체행완전한가(ByVal sourceRow As Long) As Boolean
    Dim wsIn As Worksheet
    Set wsIn = Sheets("기초자료입력")
    F018업체행완전한가 = Trim(wsIn.Cells(sourceRow, 2).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 14).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 15).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 16).Value & "") <> "" And Trim(wsIn.Cells(sourceRow, 17).Value & "") <> "" And 정수범위(wsIn.Cells(sourceRow, 14).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 15).Value, 0, 10) And 정수범위(wsIn.Cells(sourceRow, 16).Value, 0, 15) And 정수범위(wsIn.Cells(sourceRow, 17).Value, 0, 15)
End Function

Private Function F019업체행완전한가(ByVal sourceRow As Long) As Boolean
    F019업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F020업체행완전한가(ByVal sourceRow As Long) As Boolean
    F020업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F021업체행완전한가(ByVal sourceRow As Long) As Boolean
    F021업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F022업체행완전한가(ByVal sourceRow As Long) As Boolean
    F022업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F023업체행완전한가(ByVal sourceRow As Long) As Boolean
    F023업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F025업체행완전한가(ByVal sourceRow As Long) As Boolean
    F025업체행완전한가 = Trim(Sheets("기초자료입력").Cells(sourceRow, 2).Value & "") <> ""
End Function

Private Function F015임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF15 As Worksheet, wsTemp As Worksheet
    Set wsF15 = Sheets("F-015_정량적평가")
    wsF15.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F015_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F015임시시트생성 = wsTemp.Name
End Function

Private Function F016임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF16 As Worksheet, wsTemp As Worksheet
    Set wsF16 = Sheets("F-016_정성적평가")
    wsF16.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F016_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F016임시시트생성 = wsTemp.Name
End Function

Private Function F017임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF17 As Worksheet, wsTemp As Worksheet
    Set wsF17 = Sheets("F-017_제출서류자기확인서")
    wsF17.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F017_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F017임시시트생성 = wsTemp.Name
End Function

Private Function F018임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF18 As Worksheet, wsTemp As Worksheet
    Set wsF18 = Sheets("F-018_정량적평가자기평점표")
    wsF18.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F018_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F018임시시트생성 = wsTemp.Name
End Function

Private Function F019임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF19 As Worksheet, wsTemp As Worksheet
    Set wsF19 = Sheets("F-019_입찰참가신청서")
    wsF19.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F019_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("N3").Value = vendorSlot
    F019임시시트생성 = wsTemp.Name
End Function

Private Function F020임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF20 As Worksheet, wsTemp As Worksheet
    Set wsF20 = Sheets("F-020_입찰참가신고서")
    wsF20.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F020_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F020임시시트생성 = wsTemp.Name
End Function

Private Function F021임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF21 As Worksheet, wsTemp As Worksheet
    Set wsF21 = Sheets("F-021_교복납품제안서")
    wsF21.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F021_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F021임시시트생성 = wsTemp.Name
End Function

Private Function F022임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF22 As Worksheet, wsTemp As Worksheet
    Set wsF22 = Sheets("F-022_교복납품실적표")
    wsF22.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F022_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("I3").Value = vendorSlot
    F022임시시트생성 = wsTemp.Name
End Function

Private Function F023임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF23 As Worksheet, wsTemp As Worksheet
    Set wsF23 = Sheets("F-023_교복제조사양서")
    wsF23.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F023_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F023임시시트생성 = wsTemp.Name
End Function

Private Function F025임시시트생성(ByVal vendorSlot As Long) As String
    Dim wsF25 As Worksheet, wsTemp As Worksheet
    Set wsF25 = Sheets("F-025_교복AS계획서")
    wsF25.Copy After:=Sheets(Sheets.Count)
    Set wsTemp = ActiveSheet  ' .Copy 직후 Sheets(Sheets.Count) 즉시 참조가 타이밍에 따라 불안정함(2026-09-21 F-015 Subscript/1004 오류로 재현). ActiveSheet는 .Copy가 항상 그 자리에서 활성화하므로 더 안정적임
    wsTemp.Name = "F025_임시_" & vendorSlot & "_" & Format(Now, "hhnnss")
    wsTemp.Range("G3").Value = vendorSlot
    F025임시시트생성 = wsTemp.Name
End Function

Private Sub F015임시시트정리(ByRef names() As String)
    Dim i As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    For i = LBound(names) To UBound(names)
        If Left(names(i), 5) = "F015_" Or Left(names(i), 5) = "F016_" Or Left(names(i), 5) = "F017_" Or Left(names(i), 5) = "F018_" Or Left(names(i), 5) = "F019_" Or Left(names(i), 5) = "F020_" Or Left(names(i), 5) = "F021_" Or Left(names(i), 5) = "F022_" Or Left(names(i), 5) = "F023_" Or Left(names(i), 5) = "F025_" Then Sheets(names(i)).Delete
    Next i
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Sub 선택서식_인쇄미리보기()
    Dim names() As String
    Dim n As Long
    Dim callerSheet As Worksheet
    Set callerSheet = ActiveSheet
    n = 선택된시트목록(names, True)
    If n = 0 Then Exit Sub
    ThisWorkbook.Sheets(names).Select
    ActiveWindow.SelectedSheets.PrintPreview
    Call F015임시시트정리(names)
    ' 2026-09-24: 임시 업체별 시트 삭제 후 Excel이 다른 탭으로 튀는 문제(폼ID별_실행과 동일
    ' 원인)를 막기 위해 이 버튼을 누른 원래 화면으로 되돌린다.
    On Error Resume Next
    callerSheet.Activate
    On Error GoTo 0
End Sub

Sub 선택서식_PDF저장()
    Call 선택서식_PDF저장_실행(True)
End Sub

' 2026-09-24 사용자 요청: 기존 Shell()은 비동기라 "시작했습니다" 안내가 뜨자마자 콘솔
' 창이 닫혀버려 실제로 끝났는지·어디에 저장됐는지 확인할 수 없었다. WshShell.Run을
' waitOnReturn:=True로 불러 완료까지 기다린다. 한 건이면 결과 HWPX 파일을 직접 열고,
' 여러 건이면 결과 폴더를 열어 불필요하게 한글 창이 대량으로 뜨지 않게 한다.
Private Function HWPX_프로젝트루트() As String
    ' 배포본을 artifacts\excel 밖으로 복사해 열거나 Excel이 다른 작업 폴더에서 실행돼도
    ' 스크립트와 템플릿을 찾을 수 있도록, 현재 통합문서 위치부터 상위 폴더를 실제 파일 존재로 찾는다.
    Dim candidate As String, i As Long
    candidate = ThisWorkbook.Path
    For i = 0 To 4
        If Dir(candidate & "\scripts\export_hwpx_p204_from_excel.ps1") <> "" Then
            HWPX_프로젝트루트 = candidate
            Exit Function
        End If
        candidate = CreateObject("Scripting.FileSystemObject").GetParentFolderName(candidate)
        If candidate = "" Then Exit For
    Next i
End Function

Private Sub HWPX작성_실행(ByVal formIdFilter As String, ByVal requestedCount As Long)
    Dim rootPath As String, scriptPath As String, templatePath As String, outputPath As String
    rootPath = HWPX_프로젝트루트()
    If rootPath = "" Then
        MsgBox "HWPX 작성 도구가 있는 프로젝트 폴더를 찾지 못했습니다." & vbCrLf & "통합문서는 프로젝트의 artifacts\excel 폴더에서 실행하세요.", vbExclamation
        Exit Sub
    End If
    scriptPath = rootPath & "\scripts\export_hwpx_p204_from_excel.ps1"
    templatePath = rootPath & "\artifacts\hwpx\P2-04"
    outputPath = rootPath & "\artifacts\hwpx\P2-04_생성"
    If Dir(scriptPath) = "" Or Dir(templatePath, vbDirectory) = "" Then
        MsgBox "HWPX 작성 도구를 찾지 못했습니다. 프로젝트 artifacts\excel 폴더의 배포본에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    ' 2026-09-24 재조사(종료 코드 1 재현): export_hwpx_p204_from_excel.ps1은 화면에 보이는
    ' 값이 아니라 디스크에 저장된 XLSM 파일을 COM으로 다시 열어 기초자료입력!C4를 읽는다.
    ' 사용자가 기초자료입력을 입력만 하고 아직 저장하지 않은 채 바로 HWPX 작성 버튼을
    ' 누르면, 디스크의 이전 파일(학교명이 비어 있는 초기 상태일 수 있음)을 읽어
    ' "필수 허용값이 비어 있습니다: 학교명(C4)" 예외로 항상 종료 코드 1이 났다(실측 재현).
    ' 저장하기()와 동일하게 먼저 저장을 시도하고, Workbook_BeforeSave가 치명적 결함으로
    ' 저장을 취소한 경우(F-050 원자료 등)에는 구체적 사유를 안내하고 외부 스크립트를
    ' 아예 실행하지 않는다.
    If Not ThisWorkbook.Saved Then
        ThisWorkbook.Save
        If Not ThisWorkbook.Saved Then
            MsgBox "저장이 취소되어 HWPX 작성을 진행할 수 없습니다." & vbCrLf & "먼저 [저장하기]로 원인을 확인한 뒤 다시 시도하세요.", vbExclamation
            Exit Sub
        End If
    End If

    ' 2026-09-24 버그 수정: powershell.exe -File 호출에서는 명령줄 인자가 PowerShell
    ' 코드로 재해석되지 않고 그대로 문자열로 전달된다. "-FormIds @('F-007')"를 넘기면
    ' export_hwpx_p204_from_excel.ps1의 [string[]]$FormIds가 리터럴 문자열
    ' "@('F-007')" 한 개짜리 배열로 바인딩되어 실제 Form ID와 절대 일치하지 않고, 그
    ' 결과 "처리할 템플릿을 찾지 못했습니다" 오류로 항상 실패했다(사용자 실사용 재현).
    ' Form ID는 공백·따옴표가 필요 없는 단순 토큰이라 그대로 넘기면 배열 파라미터에
    ' 정상적으로 한 개짜리 배열로 바인딩된다.
    Dim formIdArg As String
    formIdArg = ""
    If formIdFilter <> "" Then formIdArg = " -FormIds " & formIdFilter

    Dim commandLine As String
    ' TemplateDir을 절대 경로로 넘긴다. WshShell.Run은 Excel 실행 위치를 현재 작업 폴더로
    ' 물려받으므로, 생략하면 artifacts\hwpx\P2-04 상대 경로를 못 찾아 모든 HWPX 작성이 실패할 수 있다.
    commandLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -ExcelPath """ & ThisWorkbook.FullName & """ -TemplateDir """ & templatePath & """ -OutputDir """ & outputPath & """" & formIdArg

    ' 2026-09-24 재조사(종료 코드 1 재현, 계속): 학교명 저장 문제를 고친 뒤에도 F-007처럼
    ' 문서번호·제출기한 같은 선택 셀이 비어 있으면 스크립트가 정상 설계대로 그 서식만
    ' 건너뛰고(SKIP), 요청한 서식이 전부 건너뛰어지면 "성공한 서식이 하나도 없습니다"로
    ' 종료 코드 1을 낸다. 기존 wsh.Run은 표준출력을 버려 사용자에게 "종료 코드 1"이라는
    ' 정보 없는 문구만 보였다(실측 재현 — 진짜 이유는 콘솔에만 찍히고 사라짐).
    ' WshShell.Exec으로 바꿔 표준출력·표준오류를 읽어 실패 사유(SKIP/FAIL 목록)를
    ' 메시지박스에 그대로 보여준다. 파이프가 가득 차 자식 프로세스가 멈추지 않도록
    ' 실행 중에도 계속 읽는다(널리 쓰이는 WshExec 폴링 패턴).
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim proc As Object
    Set proc = wsh.Exec(commandLine)
    Dim stdOutText As String, stdErrText As String
    Do While proc.Status = 0
        If Not proc.StdOut.AtEndOfStream Then stdOutText = stdOutText & proc.StdOut.ReadAll()
        If Not proc.StdErr.AtEndOfStream Then stdErrText = stdErrText & proc.StdErr.ReadAll()
        DoEvents
    Loop
    If Not proc.StdOut.AtEndOfStream Then stdOutText = stdOutText & proc.StdOut.ReadAll()
    If Not proc.StdErr.AtEndOfStream Then stdErrText = stdErrText & proc.StdErr.ReadAll()
    Dim exitCode As Long
    exitCode = proc.ExitCode

    If exitCode = 0 Then
        If requestedCount = 1 And formIdFilter <> "" Then
            Dim generatedFile As String
            generatedFile = Dir(outputPath & "\" & Trim(formIdFilter) & "_*_생성본.hwpx")
            If generatedFile <> "" Then
                MsgBox "HWPX 작성이 완료되어 결과 파일을 엽니다." & vbCrLf & outputPath & "\" & generatedFile, vbInformation
                ThisWorkbook.FollowHyperlink outputPath & "\" & generatedFile
            Else
                MsgBox "HWPX 작성은 완료됐지만 결과 파일을 찾지 못했습니다." & vbCrLf & "저장 위치: " & outputPath, vbExclamation
                Shell "explorer.exe """ & outputPath & """", vbNormalFocus
            End If
        Else
            MsgBox requestedCount & "개 HWPX 작성을 완료했습니다. 결과 폴더를 엽니다.", vbInformation
            Shell "explorer.exe """ & outputPath & """", vbNormalFocus
        End If
    Else
        Dim detail As String
        Dim summaryPos As Long
        summaryPos = InStr(stdOutText, "===== 결과 요약 =====")
        If summaryPos > 0 Then
            detail = Mid(stdOutText, summaryPos)
        ElseIf Len(Trim(stdErrText)) > 0 Then
            detail = stdErrText
        Else
            detail = "(상세 로그 없음)"
        End If
        If Len(detail) > 1200 Then detail = Left(detail, 1200) & vbCrLf & "...(이하 생략)"
        MsgBox "HWPX 작성 중 오류가 발생했습니다(종료 코드 " & exitCode & ")." & vbCrLf & vbCrLf & _
            detail & vbCrLf & vbCrLf & "저장 위치: " & outputPath, vbExclamation
    End If
End Sub

Sub 선택서식_HWPX작성()
    Dim wsSel As Worksheet, r As Long, lastRow As Long, formIds As String, selectedCount As Long
    Set wsSel = Sheets("서식선택_출력")
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastRow
        If Trim(wsSel.Cells(r, 5).Value & "") = "Y" Then
            If wsSel.Cells(r, 1).Value = True Or wsSel.Cells(r, 1).Value = 1 Then
                formIds = formIds & " " & Trim(wsSel.Cells(r, 2).Value & "")
                selectedCount = selectedCount + 1
            End If
        End If
    Next r
    If selectedCount = 0 Then
        MsgBox "먼저 [서식선택_출력] 또는 [기초자료입력]에서 작성할 서식을 체크하세요.", vbExclamation
        Exit Sub
    End If
    Call HWPX작성_실행(Trim(formIds), selectedCount)
End Sub

' 서식선택_출력의 행별 HWPX 버튼(2026-09-24 신규): 미리보기·PDF와 같은 방식으로
' Application.Caller의 버튼 위치에서 행 번호를 읽어 해당 서식 하나만 생성한다.
Sub 행별_HWPX작성()
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsSel.Shapes(Application.Caller)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    Dim r As Long
    r = shp.TopLeftCell.Row
    Call 폼ID별_HWPX실행(Trim(wsSel.Cells(r, 2).Value & ""))
End Sub

' 서식선택_출력의 행별 "해당시트이동" 버튼(2026-09-24 사용자 요청, 12번 항목): 미리보기·
' PDF·HWPX와 같은 방식으로 Application.Caller의 버튼 위치에서 행 번호를 읽어 그 행의
' 연결시트(F열)로 바로 이동한다. 출력을 거치지 않고 서식 화면 자체를 보고 싶을 때 쓴다.
Sub 행별_시트이동()
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsSel.Shapes(Application.Caller)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    Dim r As Long
    r = shp.TopLeftCell.Row
    Dim targetSheetName As String
    targetSheetName = Trim(wsSel.Cells(r, 6).Value & "")
    If targetSheetName = "" Then Exit Sub
    On Error Resume Next
    Sheets(targetSheetName).Activate
    On Error GoTo 0
End Sub

Public Sub 폼ID별_HWPX실행(ByVal formId As String)
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long, r As Long, targetRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    targetRow = 0
    For r = 5 To lastRow
        If wsSel.Cells(r, 2).Value = formId Then
            targetRow = r
            Exit For
        End If
    Next r
    If targetRow = 0 Then Exit Sub
    If Trim(wsSel.Cells(targetRow, 5).Value & "") <> "Y" Then
        MsgBox "이 서식은 아직 구현되지 않았습니다.", vbExclamation
        Exit Sub
    End If
    Call HWPX작성_실행(formId, 1)
End Sub

Public Function 검증_선택서식_PDF저장() As Boolean
    검증_선택서식_PDF저장 = 선택서식_PDF저장_실행(False)
End Function

' Err.Raise가 Application.Run 경계를 넘으면 헤드리스 자동화가 멈추거나 예측할 수 없는 COM
' 예외로 나타나는 것을 반복 확인함(2026-09-21, 검증_행별출력에서 먼저 고친 것과 같은 부류의
' 문제가 여기서도 재발함 — 검증_행별출력이 내부적으로 이 함수를 Call로 호출하기 때문에 실패가
' 그대로 상위로 전파됨). Boolean 반환 + 파일 로그(pdf_error_log.txt)로 대체함.
Private Function 선택서식_PDF저장_실행(ByVal showMessage As Boolean) As Boolean
    선택서식_PDF저장_실행 = False
    On Error GoTo ErrorHandler
    Dim callerSheet As Worksheet
    Set callerSheet = ActiveSheet
    Dim names() As String
    Dim n As Long
    n = 선택된시트목록(names, showMessage)
    If n = 0 Then Exit Function

    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\output\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    ' 파일명이 초 단위 타임스탬프뿐이면(예: 행별 버튼을 빠르게 연달아 누를 때) 같은 초에
    ' 저장된 이전 PDF를 덮어쓸 수 있음(2026-09-21 검증 중 재현). 단일 서식 출력이면 시트명을,
    ' 여러 서식 묶음이면 기존 접두어를 쓰고, 1/100초 단위 보정을 더해 충돌을 피한다.
    Dim fileLabel As String
    If n = 1 Then
        fileLabel = names(1)
    Else
        fileLabel = "선택서식"
    End If
    Dim centis As Long
    centis = CLng((Timer - Int(Timer)) * 100) Mod 100
    Dim fileName As String
    fileName = folderPath & fileLabel & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Format(centis, "00") & ".pdf"

    ThisWorkbook.Sheets(names).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Call F015임시시트정리(names)
    On Error Resume Next
    callerSheet.Activate
    On Error GoTo ErrorHandler

    If showMessage Then
        MsgBox "PDF로 저장되었습니다:" & vbCrLf & fileName, vbInformation
        ' showMessage=False는 검증_선택서식_PDF저장()의 헤드리스 안전 빌드 경로이므로
        ' 여기서만 실제 사용자 버튼 클릭(선택서식_PDF저장)에 한해 뷰어를 자동으로 연다.
        ThisWorkbook.FollowHyperlink fileName
    End If
    선택서식_PDF저장_실행 = True
    Exit Function
ErrorHandler:
    Dim errNumber As Long, errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    Call F015임시시트정리(names)
    callerSheet.Activate
    Dim logFile As Integer
    logFile = FreeFile
    Open ThisWorkbook.Path & "\pdf_error_log.txt" For Append As #logFile
    Print #logFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | 오류 " & errNumber & ": " & errDescription
    Close #logFile
    On Error GoTo 0
    If showMessage Then MsgBox "PDF 저장 중 오류가 발생했습니다: " & errDescription, vbCritical
End Function

Sub 행별_미리보기()
    Call 행별_실행(False)
End Sub

Sub 행별_PDF저장()
    Call 행별_실행(True)
End Sub

Private Sub 행별_실행(ByVal asPdf As Boolean)
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")

    Dim shp As Shape
    On Error Resume Next
    Set shp = wsSel.Shapes(Application.Caller)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    Dim r As Long
    r = shp.TopLeftCell.Row
    Dim formId As String
    formId = Trim(wsSel.Cells(r, 2).Value & "")
    Call 폼ID별_실행(formId, asPdf)
End Sub

' 기초자료입력의 "빠른 출력 선택" 패널(2026-09-21): 보기 버튼을 서식선택_출력과 동일한
' 행 번호(5~56)로 나란히 배치했으므로, 캡션 시트가 다를 뿐 같은 행 매칭 방식을 재사용한다.
Public Sub 기초자료_보기_실행()
    Dim wsIn As Worksheet, wsSel As Worksheet
    Set wsIn = Sheets("기초자료입력")
    Set wsSel = Sheets("서식선택_출력")

    Dim shp As Shape
    On Error Resume Next
    Set shp = wsIn.Shapes(Application.Caller)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub

    Dim r As Long
    r = shp.TopLeftCell.Row
    Dim formId As String
    formId = Trim(wsSel.Cells(r, 2).Value & "")
    Call 폼ID별_실행(formId, False)
End Sub

' 계약절차안내 팝업(frmContractGuide)의 미리보기/PDF 버튼도 이 공용 로직을 재사용한다(2026-09-21).
Public Sub 계약절차_서식_미리보기(ByVal formId As String)
    Call 폼ID별_실행(formId, False)
End Sub

Public Sub 계약절차_서식_PDF(ByVal formId As String)
    Call 폼ID별_실행(formId, True)
End Sub

' 계약절차안내 팝업의 HWPX 작성 버튼(2026-09-24 사용자 요청)도 행별 HWPX 로직을 재사용한다.
Public Sub 계약절차_서식_HWPX(ByVal formId As String)
    Call 폼ID별_HWPX실행(formId)
End Sub

' 출력 게이트 실패의 가장 흔한 원인 2가지(공통 필수값 누락, 위원 반복행 마스킹 형식 위반)를
' 순서대로 점검해 사람이 이해할 수 있는 구체적 이유를 돌려준다. 두 원인 모두 아니면 서식마다
' 다른 반복행(업체명·품목·평가점수 등) 요구사항을 안내하는 일반 문구로 대체한다. 2026-09-21
' 실사용자가 위원 반복행에 실명을 입력해(마스킹 형식 아님) F-001이 막힌 것을 "필요한 항목을
' 입력하라"는 일반 문구만으로는 이해하지 못한 것을 확인해 추가함.
Private Function 출력불가사유() As String
    If Not 검증_필수값검증() Then
        출력불가사유 = "기초자료입력에서 학교명·학년도·구매명·제목 등 필수 공통값을 입력하세요."
        Exit Function
    End If
    If Not 검증_위원행검증() Then
        출력불가사유 = "위원 반복행의 식별표시는 실제 성명이 아니라 '위원 ○○', '위원 **' 같은 마스킹 형식으로 입력하세요(연락처·서명은 입력하지 않습니다)."
        Exit Function
    End If
    출력불가사유 = "이 서식에 필요한 업체명·품목·위원·평가점수 등 반복행 데이터가 비어 있거나 형식에 맞지 않습니다. 기초자료입력의 해당 반복행을 확인하세요."
End Function

Private Sub 폼ID별_실행(ByVal formId As String, ByVal asPdf As Boolean)
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")

    Dim lastRow As Long, r As Long, targetRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    targetRow = 0
    For r = 5 To lastRow
        If wsSel.Cells(r, 2).Value = formId Then
            targetRow = r
            Exit For
        End If
    Next r
    If targetRow = 0 Then Exit Sub

    If Trim(wsSel.Cells(targetRow, 5).Value & "") <> "Y" Then
        MsgBox "이 서식은 아직 구현되지 않았습니다.", vbExclamation
        Exit Sub
    End If

    ' 다른 행의 선택(TRUE/FALSE) 상태를 보존한 뒤 이 행만 단독 선택하고, 기존 검증된
    ' 일괄 출력 로직(업체·위원 등 반복행 전개 포함)을 그대로 재사용한다.
    Dim i As Long
    Dim savedState() As Variant
    ReDim savedState(5 To lastRow)
    For i = 5 To lastRow
        savedState(i) = wsSel.Cells(i, 1).Value
        wsSel.Cells(i, 1).Value = (i = targetRow)
    Next i

    ' 행별_실행/폼ID별_실행은 구현상태='Y'만 확인하고, 그 서식이 실제로 출력 가능한지(학교명·
    ' 업체명 등 필수값이 채워졌는지)는 선택된시트목록의 개별 출력가능() 게이트가 따로 판정함.
    ' 예전에는 이 실패를 선택된시트목록 내부의 일괄 선택용 메시지("구현된 서식이 선택되지
    ' 않았습니다")에 그대로 맡겼는데, 행 하나만 다루는 이 경로에서는 그 문구가 "왜 안 되는지"를
    ' 설명하지 못해 혼란을 준다는 사용자 지적(2026-09-21)에 따라 명확한 안내로 분리함.
    Dim names() As String
    Dim n As Long
    On Error Resume Next
    n = 선택된시트목록(names, False)
    On Error GoTo 0

    If n = 0 Then
        For i = 5 To lastRow
            wsSel.Cells(i, 1).Value = savedState(i)
        Next i
        MsgBox "이 서식은 지금 출력할 수 없습니다." & vbCrLf & 출력불가사유(), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    If asPdf Then
        Dim folderPath As String
        folderPath = ThisWorkbook.Path & "\output\"
        If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
        Dim fileLabel As String
        If n = 1 Then fileLabel = names(1) Else fileLabel = "선택서식"
        Dim centis As Long
        centis = CLng((Timer - Int(Timer)) * 100) Mod 100
        Dim fileName As String
        fileName = folderPath & fileLabel & "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Format(centis, "00") & ".pdf"
        ThisWorkbook.Sheets(names).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        MsgBox "PDF로 저장되었습니다:" & vbCrLf & fileName, vbInformation
        ThisWorkbook.FollowHyperlink fileName
    Else
        ThisWorkbook.Sheets(names).Select
        ActiveWindow.SelectedSheets.PrintPreview
    End If
    Call F015임시시트정리(names)
    ' 2026-09-24 사용자 실사용 확인: 위 정리가 임시 업체별 시트(탭 순서 맨 끝)를 지우면
    ' Excel이 자동으로 그 앞 탭(F-050)을 활성화해, 미리보기를 닫으면 엉뚱한 서식으로 이동한
    ' 것처럼 보였음. 원래 호출 화면(서식선택_출력)으로 명시적으로 되돌린다.
    wsSel.Activate
    On Error GoTo 0

    For i = 5 To lastRow
        wsSel.Cells(i, 1).Value = savedState(i)
    Next i
End Sub

' Application.Run 경계를 넘어 Err.Raise를 던지면 헤드리스 자동화 세션에서
' VBA 오류 대화상자가 뜬 채 무응답으로 멈추는 것을 2026-09-21에 직접 재현·확인함.
' 그래서 실패를 예외가 아닌 Boolean 반환값으로 알린다(Form ID 없음/미구현 서식은 False).
Public Function 검증_행별출력(ByVal formId As String, ByVal asPdf As Boolean) As Boolean
    검증_행별출력 = False
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long, r As Long, targetRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    targetRow = 0
    For r = 5 To lastRow
        If wsSel.Cells(r, 2).Value = formId Then
            targetRow = r
            Exit For
        End If
    Next r
    If targetRow = 0 Then Exit Function
    If Trim(wsSel.Cells(targetRow, 5).Value & "") <> "Y" Then Exit Function

    Dim i As Long
    Dim savedState() As Variant
    ReDim savedState(5 To lastRow)
    For i = 5 To lastRow
        savedState(i) = wsSel.Cells(i, 1).Value
        wsSel.Cells(i, 1).Value = (i = targetRow)
    Next i

    If asPdf Then
        Call 검증_선택서식_PDF저장()
    Else
        Dim names() As String
        Dim n As Long
        n = 선택된시트목록(names, False)
        If n > 0 Then Call F015임시시트정리(names)
    End If

    For i = 5 To lastRow
        wsSel.Cells(i, 1).Value = savedState(i)
    Next i
    검증_행별출력 = True
End Function
'@
    $codeOutput = $codeOutput -replace "`r`n", "`r" -replace "`n", "`r"
    $modOutput.CodeModule.AddFromString($codeOutput)
    L "Module_출력 추가 완료 (줄 수: $($modOutput.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT2_SAVE_OK" } catch { L "DIAG_CHECKPOINT2_SAVE_FAIL: $($_.Exception.Message)" }

    $modNavigation = $vbproj.VBComponents.Add(1)
    $modNavigation.Name = "Module_검색_절차"
    $codeNavigation = @'
Option Explicit

' 계약 절차 안내 팝업(2026-09-21 사용자 요청): 금액 기준 자동판정은 하지 않고, 계약방법
' 3종 버튼을 항상 노출한 뒤 클릭 시 frmContractGuide 폼으로 단계 흐름·설명·관련 서식을 안내한다.
Sub 계약절차안내_1인견적()
    Call 계약절차안내_열기("1인견적 수의계약")
End Sub

Sub 계약절차안내_2인견적()
    Call 계약절차안내_열기("2인견적 수의계약")
End Sub

Sub 계약절차안내_입찰계약()
    Call 계약절차안내_열기("입찰계약")
End Sub

Public Sub 계약절차안내_열기(ByVal contractType As String)
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmContractGuide")
    Call frm.초기화(contractType)
    ' 모달(vbModal, 기본값)로 띄운 채 미리보기(PrintPreview, 앱 전체를 차지하는 백스테이지
    ' 인쇄 화면)를 부르면 두 UI가 서로 제어권을 다퉈 화면이 멈추는 것을 실사용 중 확인함
    ' (2026-09-21). 메인 창과 공존 가능한 모들리스(vbModeless=0)로 띄워 해결함.
    frm.Show 0
End Sub

' 헤드리스 자동 검증용(2026-09-21): .Show는 모달 팝업이라 자동화가 멈추므로 절대 호출하지
' 않고, 초기화()가 채운 단계 목록·첫 단계 설명·관련서식 목록만 조회한 뒤 Unload한다.
Public Function 검증_계약절차폼(ByVal contractType As String) As String
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmContractGuide")
    Call frm.초기화(contractType)
    Dim result As String
    result = frm.lstSteps.ListCount & "|" & frm.txtDesc.Value & "|" & frm.lstForms.ListCount
    Unload frm
    검증_계약절차폼 = result
End Function

' frmLoadPicker(2026-09-24)를 .Show 없이(모달 대기 없이) 목록 채움만 검증하는 헤드리스 진입점.
Public Function 검증_불러오기팝업() As String
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmLoadPicker")
    Call frm.초기화
    Dim result As String
    result = frm.lstRecords.ListCount
    Unload frm
    검증_불러오기팝업 = result
End Function

Sub 계약방법안내_기초자료반영()
    Call 계약방법안내_기초자료반영_내부(True)
End Sub

Public Sub 검증_계약방법안내_기초자료반영()
    Call 계약방법안내_기초자료반영_내부(False)
End Sub

Private Sub 계약방법안내_기초자료반영_내부(ByVal showMessage As Boolean)
    Dim wsMethod As Worksheet, selectedMethod As String
    Set wsMethod = Sheets("계약방법안내")
    If Trim(wsMethod.Range("B5").Value & "") = "" Then
        If showMessage Then MsgBox "계약방식을 선택하세요.", vbExclamation
        Exit Sub
    End If
    selectedMethod = Trim(wsMethod.Range("B7").Value & "")
    If selectedMethod = "" Then
        If showMessage Then MsgBox "반영할 계약방법 안내값이 없습니다.", vbExclamation
        Exit Sub
    End If
    Sheets("기초자료입력").Range("C14").Value = selectedMethod
    If showMessage Then MsgBox "계약방식(B-03)에 매뉴얼 예시를 반영했습니다. 출력 전 최신 법령·지침과 사실관계를 최종 확인하세요.", vbInformation
End Sub

Public Sub 계약방식선택_1인견적(): Call 계약방식즉시반영("1인견적 수의계약", True): End Sub
Public Sub 계약방식선택_2인견적(): Call 계약방식즉시반영("2인견적 수의계약", True): End Sub
Public Sub 계약방식선택_2단계입찰(): Call 계약방식즉시반영("2단계 입찰(규격·가격 동시)", True): End Sub
Public Sub 계약방식선택_일반입찰(): Call 계약방식즉시반영("일반경쟁입찰", True): End Sub
Public Sub 검증_계약방식선택(ByVal methodName As String)
    Call 계약방식즉시반영(methodName, False)
End Sub
Private Sub 계약방식즉시반영(ByVal methodName As String, ByVal showGuide As Boolean)
    Sheets("계약방법안내").Range("B5").Value = methodName
    Sheets("기초자료입력").Range("C14").Value = methodName
    Application.Calculate
    If showGuide Then
        If methodName = "1인견적 수의계약" Or methodName = "2인견적 수의계약" Then
            Call 계약절차안내_열기(methodName)
        Else
            Call 계약절차안내_열기("입찰계약")
        End If
    End If
End Sub

Sub 학교검색_실행()
    Call 학교검색_실행_내부(True)
End Sub

Public Sub 검증_학교검색_실행()
    Call 학교검색_실행_내부(False)
End Sub

Private Sub 학교검색_실행_내부(ByVal showMessage As Boolean)
    Dim wsSearch As Worksheet, wsSchool As Worksheet
    Dim schoolName As String, regionName As String, schoolLevel As String
    Dim lastRow As Long, sourceRow As Long, resultRow As Long, count As Long

    Set wsSearch = Sheets("학교검색")
    Set wsSchool = Sheets("학교정보")
    schoolName = Trim(wsSearch.Range("B3").Value & "")
    regionName = Trim(wsSearch.Range("D3").Value & "")
    schoolLevel = Trim(wsSearch.Range("F3").Value & "")
    wsSearch.Range("A7:F796").ClearContents

    lastRow = wsSchool.Cells(wsSchool.Rows.Count, 5).End(xlUp).Row
    resultRow = 7
    count = 0
    For sourceRow = 2 To lastRow
        If (schoolName = "" Or InStr(1, wsSchool.Cells(sourceRow, 5).Value & "", schoolName, vbTextCompare) > 0) And _
           (regionName = "" Or InStr(1, wsSchool.Cells(sourceRow, 2).Value & "", regionName, vbTextCompare) > 0) And _
           (schoolLevel = "" Or InStr(1, wsSchool.Cells(sourceRow, 3).Value & "", schoolLevel, vbTextCompare) > 0) Then
            wsSearch.Cells(resultRow, 1).Value = wsSchool.Cells(sourceRow, 1).Value
            wsSearch.Cells(resultRow, 2).Value = wsSchool.Cells(sourceRow, 2).Value
            wsSearch.Cells(resultRow, 3).Value = wsSchool.Cells(sourceRow, 3).Value
            wsSearch.Cells(resultRow, 4).Value = wsSchool.Cells(sourceRow, 5).Value
            wsSearch.Cells(resultRow, 5).Value = wsSchool.Cells(sourceRow, 6).Value
            wsSearch.Cells(resultRow, 6).Value = wsSchool.Cells(sourceRow, 7).Value
            resultRow = resultRow + 1
            count = count + 1
        End If
    Next sourceRow
    If showMessage Then MsgBox count & "건을 찾았습니다. 결과 행을 선택한 뒤 [선택 학교를 기초자료에 반영]을 누르세요.", vbInformation
End Sub

Sub 선택학교_기초자료반영()
    Call 선택학교_기초자료반영_내부(True)
End Sub

Public Sub 검증_선택학교_기초자료반영()
    Call 선택학교_기초자료반영_내부(False)
End Sub

Private Sub 선택학교_기초자료반영_내부(ByVal showMessage As Boolean)
    Dim wsSearch As Worksheet
    Dim resultRow As Long, selectedSchool As String
    Set wsSearch = Sheets("학교검색")
    If ActiveSheet.Name <> wsSearch.Name Then
        If showMessage Then MsgBox "학교검색 시트의 검색 결과 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    resultRow = ActiveCell.Row
    If resultRow < 7 Or resultRow > 796 Then
        If showMessage Then MsgBox "검색 결과의 학교 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    selectedSchool = Trim(wsSearch.Cells(resultRow, 4).Value & "")
    If selectedSchool = "" Then
        If showMessage Then MsgBox "선택한 행에 학교명이 없습니다. 먼저 검색하세요.", vbExclamation
        Exit Sub
    End If
    Sheets("기초자료입력").Range("C4").Value = selectedSchool
    If showMessage Then MsgBox "학교명(C-01)에 '" & selectedSchool & "'을(를) 반영했습니다.", vbInformation
End Sub

' 서식이 너무 많아 각 F-XXX 시트에서 목록으로 돌아가기 어렵다는 사용자 요청(2026-09-21)에
' 따라, 각 서식 시트 우측 상단(인쇄영역 밖)에 이동 버튼 2개를 배치하며 공유하는 매크로.
Sub 이동_서식선택출력()
    Sheets("서식선택_출력").Activate
End Sub

Sub 이동_계약방법안내()
    Sheets("계약방법안내").Activate
End Sub

Sub 이동_기초자료입력()
    Sheets("기초자료입력").Activate
End Sub

' DB 시트 접근 개선(요청 4, 2026-09-23): DB·DB_품목·DB_업체·DB_위원·DB_평가 5개 시트는 학교·업체·
' 위원 등 실제 업무·개인정보를 담고 있어 계속 숨김(DB=Hidden, 나머지=VeryHidden) 상태로 유지하되
' (사용자가 실수로 값을 직접 편집하지 못하도록), 관리자가 점검할 때만 버튼으로 열람하고 다시
' 숨기도록 왕복 매크로를 제공한다. 시트를 항상 보이게 바꾸는 것은 "이동"이라는 요청 취지를 넘는
' 과도한 노출이라 판단해 선택하지 않았다(로그.md 2026-09-23 "요청 4" 항목 근거 기록 참고).
' showMessage=False 경로(검증_ 접두)는 자동 검증 스크립트가 MsgBox 없이 헤드리스로 토글을
' 확인할 수 있도록 분리함(관련FormID로이동_내부와 동일한 기존 패턴). 버튼에는 항상 사용자
' 확인 메시지가 있는 공개 Sub만 연결한다.
Private Sub DB시트토글_내부(ByVal openIt As Boolean, ByVal showMessage As Boolean)
    If openIt Then
        If showMessage Then
            Dim answer As VbMsgBoxResult
            answer = MsgBox("DB 시트에는 학교·업체·위원 등 실제 업무·개인정보가 저장되어 있습니다." & vbCrLf & _
                "값을 직접 고치지 말고 확인 목적으로만 사용하세요. 계속하시겠습니까?", vbYesNo + vbExclamation, "DB 시트 열기")
            If answer <> vbYes Then Exit Sub
        End If
        Sheets("DB").Visible = xlSheetVisible
        Sheets("DB_품목").Visible = xlSheetVisible
        Sheets("DB_업체").Visible = xlSheetVisible
        Sheets("DB_위원").Visible = xlSheetVisible
        Sheets("DB_평가").Visible = xlSheetVisible
        If showMessage Then
            Sheets("DB").Activate
            MsgBox "확인이 끝나면 DB 시트의 [DB 닫기(다시 숨김)] 버튼을 눌러 원래대로 숨겨주세요.", vbInformation
        End If
    Else
        Sheets("DB").Visible = xlSheetHidden
        Sheets("DB_품목").Visible = xlSheetVeryHidden
        Sheets("DB_업체").Visible = xlSheetVeryHidden
        Sheets("DB_위원").Visible = xlSheetVeryHidden
        Sheets("DB_평가").Visible = xlSheetVeryHidden
        If showMessage Then Sheets("기초자료입력").Activate
    End If
End Sub

Sub DB시트열기()
    Call DB시트토글_내부(True, True)
End Sub

Sub DB시트닫기()
    Call DB시트토글_내부(False, True)
End Sub

Public Sub 검증_DB시트열기()
    Call DB시트토글_내부(True, False)
End Sub

Public Sub 검증_DB시트닫기()
    Call DB시트토글_내부(False, False)
End Sub

Sub 관련FormID로이동()
    Call 관련FormID로이동_내부(True)
End Sub

Public Sub 검증_관련FormID로이동()
    Call 관련FormID로이동_내부(False)
End Sub

Private Sub 관련FormID로이동_내부(ByVal showMessage As Boolean)
    Dim wsFlow As Worksheet, wsSelect As Worksheet
    Dim flowRow As Long, formId As String, lastRow As Long, r As Long
    Set wsFlow = Sheets("입찰계약절차안내")
    Set wsSelect = Sheets("서식선택_출력")
    If ActiveSheet.Name <> wsFlow.Name Then
        If showMessage Then MsgBox "입찰계약절차안내 시트의 단계 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    flowRow = ActiveCell.Row
    If flowRow < 5 Or flowRow > 13 Then
        If showMessage Then MsgBox "1~9단계의 행을 선택하세요.", vbExclamation
        Exit Sub
    End If
    formId = Trim(wsFlow.Cells(flowRow, 4).Value & "")
    lastRow = wsSelect.Cells(wsSelect.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastRow
        If wsSelect.Cells(r, 2).Value = formId Then
            wsSelect.Activate
            wsSelect.Cells(r, 2).Select
            Exit Sub
        End If
    Next r
    If showMessage Then MsgBox "이동할 Form ID를 찾지 못했습니다: " & formId, vbExclamation
End Sub
'@
    $codeNavigation = $codeNavigation -replace "`r`n", "`r" -replace "`n", "`r"
    $modNavigation.CodeModule.AddFromString($codeNavigation)
    L "Module_검색_절차 추가 완료 (줄 수: $($modNavigation.CodeModule.CountOfLines))"
    try { $wb.Save(); L "DIAG_CHECKPOINT3_SAVE_OK" } catch { L "DIAG_CHECKPOINT3_SAVE_FAIL: $($_.Exception.Message)" }

    # ---- UserForm: frmContractGuide (계약 절차 안내 팝업, 2026-09-21 사용자 요청) ----
    # 원본 XLSM UserForm2("계약 절차 안내")의 좌측 단계 목록 + 우측 설명 + 관련서류 열기 구조를
    # 참고하되, 원본의 금액 기준 자동판정은 재현하지 않음(2026-09-20 결정 유지). 관련 서식은
    # 원문(이미 구현된 F-XXX)만 연결하고, 클릭 시 Module_출력의 행별 미리보기/PDF 로직을 재사용함.
    $existingFrm = $null
    foreach ($c in $vbproj.VBComponents) {
        if ($c.Name -eq "frmContractGuide") { $existingFrm = $c; break }
    }
    if ($existingFrm) {
        # 검증본을 기반으로 재주입할 때는 이미 검증된 폼을 그대로 쓴다. Excel COM은
        # UserForm 삭제 뒤 같은 이름의 폼을 즉시 추가하면 CTL_E_PATHFILEACCESSERROR로
        # VBE 창을 남길 수 있어, 폼 내부를 바꾸지 않는 이번 UX 변경에서는 재생성하지 않는다.
        L "기존 frmContractGuide UserForm 유지"
    } else {
    $frmComp = $vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    $frmComp.Name = "frmContractGuide"
    $frmDesigner = $frmComp.Designer
    # 외부 COM(PowerShell)에서는 Designer 객체가 Width/Height/Caption을 직접 노출하지 않고
    # "해당 개체는 'Width' 속성을 지원하지 않습니다" 오류가 남(2026-09-21 실측). VBComponent의
    # Properties 컬렉션을 통해 이름으로 설정하면 우회됨. Controls.Add는 Designer에서 그대로 동작함.
    # 2026-09-24 사용자 실사용 확인(스크린샷 5.JPG): Height=470에 버튼 Top=432를 두면
    # 창 아래쪽 여백이 실제 렌더링 시 캡션바·테두리에 밀려 버튼이 창 밖으로 잘려 안 보임
    # (기존 검증됐던 배치는 Height=430/버튼Top=390, 여백 16px였는데 재디자인 때 여백을
    # 10px로 좁힌 것이 원인). 목록-버튼 사이 빈 공간도 없애고 여백을 24px로 넉넉히 둔다.
    # 2026-09-24 재조사(항목 7, "버튼 아래 여백이 없어 답답함"): 잘림 방지를 위해 Height=400/
    # 버튼Top=346으로 고정했던 이전 배치는 잘리지는 않지만 목록-버튼 간격(18px)과 버튼-창
    # 하단 여백(26px)이 빠듯해 답답해 보인다는 지적을 받음. 폼 높이를 20px 늘리고 버튼을
    # 16px 아래로 내려 위·아래 여백을 모두 넉넉하게 확보한다(목록-버튼 34px, 버튼-하단 30px).
    # 2026-09-24 재조사(4차, 실제 실행 화면 확인): Height=420에서도 버튼 밑 여백이 여전히
    # 부족해 보인다는 지적을 받음. 버튼 위치(Top=362)는 그대로 두고 폼 높이만 32px 더 늘려
    # 버튼-하단 여백을 30px→62px로 넉넉히 확보한다(버튼을 아래로 밀지 않으므로 이전에
    # Height=470/Top=432 조합에서 났던 잘림 문제와는 무관함).
    $frmComp.Properties("Width").Value = 560
    $frmComp.Properties("Height").Value = 452
    $frmComp.Properties("Caption").Value = "계약 절차 안내"

    # 2026-09-24 재디자인(사용자 제공 시안 시안-5-*.png 참고): 상단 색상 배너 + 좌측 단계
    # 목록 + 우측 설명·서식 2열 표 + 하단 강조 버튼 행. MSForms 컨트롤은 원본 시안의
    # 둥근 모서리·아이콘까지는 재현할 수 없어, 배색·타이포·여백·2열 표로 최대한 근접시킨다.
    $fraHeader = $frmDesigner.Controls.Add("Forms.Frame.1")
    $fraHeader.Name = "fraHeader"
    $fraHeader.Left = 0; $fraHeader.Top = 0; $fraHeader.Width = 560; $fraHeader.Height = 40
    $fraHeader.SpecialEffect = 0
    $fraHeader.Caption = ""

    $lblTitle = $frmDesigner.Controls.Add("Forms.Label.1")
    $lblTitle.Name = "lblTitle"
    $lblTitle.Left = 12; $lblTitle.Top = 10; $lblTitle.Width = 500; $lblTitle.Height = 20
    # 외부 COM에서 MSForms 컨트롤의 .Font 접근은 "개체 변수가 설정되지 않았습니다" 오류로
    # 실패함(2026-09-21 실측, 워크시트 Range.Font와는 다른 경로). 색상·굵기·크기는 VBA
    # 런타임 코드(초기화 Sub, 아래 $codeFrm)에서 설정한다 — VBA 자체 실행은 이 제약이 없음.

    $lblDescTitle = $frmDesigner.Controls.Add("Forms.Label.1")
    $lblDescTitle.Name = "lblDescTitle"
    $lblDescTitle.Left = 156; $lblDescTitle.Top = 50; $lblDescTitle.Width = 390; $lblDescTitle.Height = 18

    $lstSteps = $frmDesigner.Controls.Add("Forms.ListBox.1")
    $lstSteps.Name = "lstSteps"
    $lstSteps.Left = 12; $lstSteps.Top = 50; $lstSteps.Width = 136; $lstSteps.Height = 296

    $txtDesc = $frmDesigner.Controls.Add("Forms.TextBox.1")
    $txtDesc.Name = "txtDesc"
    $txtDesc.Left = 156; $txtDesc.Top = 70; $txtDesc.Width = 390; $txtDesc.Height = 90
    $txtDesc.MultiLine = $true
    $txtDesc.ScrollBars = 2  # fmScrollBarsVertical
    $txtDesc.Locked = $true
    $txtDesc.SpecialEffect = 0
    $txtDesc.BorderStyle = 0

    $lblFormsTitle = $frmDesigner.Controls.Add("Forms.Label.1")
    $lblFormsTitle.Name = "lblFormsTitle"
    $lblFormsTitle.Left = 156; $lblFormsTitle.Top = 172; $lblFormsTitle.Width = 390; $lblFormsTitle.Height = 16
    $lblFormsTitle.Caption = "관련 서식  (선택 후 미리보기 / PDF 저장 / HWPX 작성)"

    $lblColCode = $frmDesigner.Controls.Add("Forms.Label.1")
    $lblColCode.Name = "lblColCode"
    $lblColCode.Left = 160; $lblColCode.Top = 192; $lblColCode.Width = 60; $lblColCode.Height = 14
    $lblColCode.Caption = "코드"

    $lblColName = $frmDesigner.Controls.Add("Forms.Label.1")
    $lblColName.Name = "lblColName"
    $lblColName.Left = 224; $lblColName.Top = 192; $lblColName.Width = 200; $lblColName.Height = 14
    $lblColName.Caption = "서식명"

    $lstForms = $frmDesigner.Controls.Add("Forms.ListBox.1")
    $lstForms.Name = "lstForms"
    $lstForms.Left = 156; $lstForms.Top = 208; $lstForms.Width = 390; $lstForms.Height = 120
    $lstForms.ColumnCount = 2
    $lstForms.ColumnWidths = "60pt;280pt"

    $btnPreview = $frmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnPreview.Name = "btnPreview"
    $btnPreview.Left = 12; $btnPreview.Top = 362; $btnPreview.Width = 128; $btnPreview.Height = 28
    $btnPreview.Caption = "미리보기"

    $btnPDF = $frmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnPDF.Name = "btnPDF"
    $btnPDF.Left = 148; $btnPDF.Top = 362; $btnPDF.Width = 128; $btnPDF.Height = 28
    $btnPDF.Caption = "PDF 저장"

    # HWPX 작성 버튼(2026-09-24 사용자 요청): 관련서식 강조색과 맞춰 계약유형별 테마색을 채움(주 동작)
    $btnHWPX = $frmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnHWPX.Name = "btnHWPX"
    $btnHWPX.Left = 284; $btnHWPX.Top = 362; $btnHWPX.Width = 148; $btnHWPX.Height = 28
    $btnHWPX.Caption = "HWPX 작성"

    $btnClose = $frmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnClose.Name = "btnClose"
    $btnClose.Left = 440; $btnClose.Top = 362; $btnClose.Width = 108; $btnClose.Height = 28
    $btnClose.Caption = "닫기"

    $codeFrm = @'
Option Explicit

Private stepNames() As String
Private stepDescriptions() As String
Private stepFormIDs() As String
Private stepFormLabels() As String
Private stepCount As Long

Private Sub ApplyTheme(ByVal contractType As String)
    ' 2026-09-24 재디자인: 계약유형별 강조색(시안-5-*.png 참고 — 입찰계약 청록, 1인견적 파랑,
    ' 2인견적 진파랑, 2단계 입찰 남보라)을 상단 배너·안내 제목·주 동작 버튼에 일관 적용한다.
    Dim themeColor As Long
    Select Case contractType
        Case "입찰계약": themeColor = RGB(13, 148, 136)
        Case "1인견적 수의계약": themeColor = RGB(37, 99, 235)
        Case "2인견적 수의계약": themeColor = RGB(29, 63, 189)
        Case Else: themeColor = RGB(79, 70, 229)
    End Select

    Me.fraHeader.BackColor = themeColor
    Me.lblTitle.Caption = "계약 절차 안내 - " & contractType
    Me.lblTitle.BackStyle = 0
    Me.lblTitle.ForeColor = RGB(255, 255, 255)
    Me.lblTitle.Font.Bold = True
    Me.lblTitle.Font.Size = 13

    Me.lblDescTitle.Caption = contractType & " 안내"
    Me.lblDescTitle.ForeColor = themeColor
    Me.lblDescTitle.Font.Bold = True
    Me.lblDescTitle.Font.Size = 12

    Me.lblFormsTitle.Font.Bold = True
    Me.lblColCode.Font.Bold = True
    Me.lblColName.Font.Bold = True

    Me.btnHWPX.BackColor = themeColor
    Me.btnHWPX.ForeColor = RGB(255, 255, 255)
    Me.btnHWPX.Font.Bold = True
    Me.btnPreview.BackColor = RGB(240, 253, 250)
    Me.btnPDF.BackColor = RGB(254, 242, 242)
    Me.btnClose.BackColor = RGB(238, 238, 238)
End Sub

Public Sub 초기화(ByVal contractType As String)
    Me.Caption = "계약 절차 안내 - " & contractType
    Call ApplyTheme(contractType)

    If contractType = "입찰계약" Then
        stepCount = 11
        ReDim stepNames(1 To stepCount): ReDim stepDescriptions(1 To stepCount)
        ReDim stepFormIDs(1 To stepCount): ReDim stepFormLabels(1 To stepCount)

        stepNames(1) = "1. 입찰공고문 작성"
        stepDescriptions(1) = "교복 사양서와 함께 입찰공고문을 작성해 학교 누리집 및 국가종합전자조달시스템(나라장터, G2B)에 게재합니다. 2단계 입찰(규격·가격 동시)이 원칙이며, 규격제안서는 학교에 직접 제출, 가격입찰서는 G2B로 제출합니다."
        stepFormIDs(1) = "F-009,F-010,F-011,F-008,F-012"
        stepFormLabels(1) = "F-009 기초금액계약방법결정,F-010 사전규격공개기안문,F-011 입찰공고,F-008 교복사양서,F-012 계약특수조건"

        stepNames(2) = "2. 지정정보처리장치 공고"
        stepDescriptions(2) = "국가종합전자조달시스템(나라장터, G2B)에 입찰공고를 게재합니다. 이용자 등록을 마친 업체만 참여할 수 있으며, 전자입찰특별유의서 제7조에 따른 신원확인 입찰이 적용됩니다."
        stepFormIDs(2) = "F-011"
        stepFormLabels(2) = "F-011 입찰공고"

        stepNames(3) = "3. 입찰(제안서 제출)"
        stepDescriptions(3) = "업체가 제출서류 자기확인서, 입찰참가신청서, 납품 제안서, 실적표, 제조 사양서, 단가 비율표, A/S 계획서 등을 제출합니다. 규격제안서는 블라인드 심사를 위해 업체명·표시를 제거합니다."
        stepFormIDs(3) = "F-017,F-018,F-019,F-020,F-021,F-022,F-023,F-024,F-025,F-027,F-029,F-030"
        stepFormLabels(3) = "F-017 제출서류자기확인서,F-018 정량적평가자기평점표,F-019 입찰참가신청서,F-020 입찰참가신고서,F-021 교복납품제안서,F-022 교복납품실적표,F-023 교복제조사양서,F-024 단가비율표,F-025 교복AS계획서,F-027 위임장,F-029 개인정보제공동의서,F-030 청렴계약이행서약서"

        ' 2026-09-23 버그 수정: 기존에는 "4.개찰(가격 먼저 개봉) → 5.적격심사(규격 평가)" 순서였으나,
        ' 이는 F-011 입찰공고 자체의 8번 항목("규격평가 적격업체에 한해 가격을 개찰하고, 예정가격
        ' 이하 최저가격 입찰자를 낙찰자로 선정") 및 2단계 입찰(규격·가격 동시 제출, 규격심사 통과
        ' 업체에 한해서만 가격입찰서를 개봉하는 방식)의 정의와 정반대 순서였음. F-011에 이미 반영된
        ' 문구를 그대로 근거로 순서를 "규격(적격)심사 먼저 → 가격 개찰"로 바로잡음(새 법령·절차 해석을
        ' 추가하지 않고 같은 코드베이스에 이미 있던 문구와 일치시킴).
        stepNames(4) = "4. 규격(적격)심사"
        stepDescriptions(4) = "교복선정위원회가 제안서 평가항목·배점기준에 따라 정량·정성 평가를 실시합니다(2단계 입찰의 1단계: 규격제안서만 먼저 평가). 종합평점이 기준(적격) 이상인 업체만 다음 단계(가격 개찰) 대상이 됩니다."
        stepFormIDs(4) = "F-014,F-015,F-016,F-032,F-033,F-034,F-035,F-036,F-037,F-038,F-039,F-040"
        stepFormLabels(4) = "F-014 평가항목배점기준,F-015 정량적평가,F-016 정성적평가,F-032 제안서접수결과,F-033 제안서접수대장,F-034 평가위원회개최,F-035 정량평가결과,F-036 제안서평가결과,F-037 평가위원회참석등록부,F-038 업체참가등록부,F-039 업체별제안서평가표,F-040 위원청렴보안서약서"

        stepNames(5) = "5. 가격 개찰"
        stepDescriptions(5) = "규격(적격)심사를 통과한 업체에 한해서만 G2B에 제출된 가격입찰서를 개찰합니다(2단계 입찰의 2단계: 가격은 규격 통과 업체만 개봉). 예정가격 이하 최저가 입찰자를 낙찰자로 선정합니다."
        stepFormIDs(5) = "": stepFormLabels(5) = ""

        stepNames(6) = "6. 낙찰자 결정"
        stepDescriptions(6) = "적격심사를 통과한 최저가 입찰자를 낙찰자로 결정합니다. 상호·법인 명칭, 대표자 등록 사항 등 업체정보를 확인해 결격 여부를 판단합니다."
        stepFormIDs(6) = "F-041": stepFormLabels(6) = "F-041 낙찰자결정"

        stepNames(7) = "7. 계약 체결"
        stepDescriptions(7) = "낙찰자에게 결과를 통보하고 계약을 체결합니다. 계약금액은 낙찰단가x예정수량으로 계산하며, 계약 특수조건(F-012)을 첨부합니다."
        stepFormIDs(7) = "F-042,F-043,F-031"
        stepFormLabels(7) = "F-042 낙찰자결정통보,F-043 계약체결,F-031 교복품목별금액표"

        stepNames(8) = "8. 계약 이행"
        stepDescriptions(8) = "공급자가 학생 신체 치수를 측정하고 교복을 제작·납품합니다(계약특수조건 제1~2조). 신입생·재학생·전입생의 추가 구매 요청에도 계약 단가로 대응해야 합니다."
        stepFormIDs(8) = "": stepFormLabels(8) = ""

        stepNames(9) = "9. 검사·검수"
        stepDescriptions(9) = "학교가 납품된 교복이 규격서와 일치하는지 검사·검수합니다(계약특수조건 제6조). 최종납품확인서 제출과 학교의 검사·검수 합격이 있어야 납품이 완료됩니다."
        stepFormIDs(9) = "": stepFormLabels(9) = ""

        stepNames(10) = "10. 대가 지급"
        stepDescriptions(10) = "검사·검수 완료 후 공급자의 청구에 따라 5일 이내(토·공휴일 제외) 대금을 지급합니다(계약특수조건 제12조)."
        stepFormIDs(10) = "": stepFormLabels(10) = ""

        stepNames(11) = "11. 변경계약"
        stepDescriptions(11) = "원단 시장가격 변동 등을 사유로 계약금액이나 단가를 임의로 변경할 수 없습니다(계약특수조건 제13조). 부득이한 변경이 필요하면 학교와 공급자가 협의해 별도 절차를 거칩니다."
        stepFormIDs(11) = "": stepFormLabels(11) = ""

    ElseIf contractType = "2인견적 수의계약" Then
        stepCount = 5
        ReDim stepNames(1 To stepCount): ReDim stepDescriptions(1 To stepCount)
        ReDim stepFormIDs(1 To stepCount): ReDim stepFormLabels(1 To stepCount)

        stepNames(1) = "1. 견적 제출 요청"
        stepDescriptions(1) = "추정가격 기준에 맞는 2개 이상 업체에 견적서 제출을 요청합니다. 교복 사양서를 함께 제공해 동일한 조건으로 견적을 받습니다."
        stepFormIDs(1) = "F-008,F-012": stepFormLabels(1) = "F-008 교복사양서,F-012 계약특수조건"

        stepNames(2) = "2. 견적서 접수·비교"
        stepDescriptions(2) = "제출된 견적서를 비교해 가장 유리한 조건의 업체를 선정합니다. 견적 비교 근거는 관련 서류로 남겨 둡니다."
        stepFormIDs(2) = "": stepFormLabels(2) = ""

        stepNames(3) = "3. 계약 체결"
        stepDescriptions(3) = "선정된 업체와 계약을 체결합니다. 계약 특수조건(F-012)을 첨부하고 계약금액은 견적가x예정수량으로 산정합니다."
        stepFormIDs(3) = "F-043,F-031": stepFormLabels(3) = "F-043 계약체결,F-031 교복품목별금액표"

        stepNames(4) = "4. 계약 이행·검사검수"
        stepDescriptions(4) = "공급자가 치수를 측정하고 교복을 제작·납품하면, 학교가 규격서 일치 여부를 검사·검수합니다."
        stepFormIDs(4) = "": stepFormLabels(4) = ""

        stepNames(5) = "5. 대가 지급"
        stepDescriptions(5) = "검사·검수 완료 후 대금을 지급합니다."
        stepFormIDs(5) = "": stepFormLabels(5) = ""

    Else
        stepCount = 4
        ReDim stepNames(1 To stepCount): ReDim stepDescriptions(1 To stepCount)
        ReDim stepFormIDs(1 To stepCount): ReDim stepFormLabels(1 To stepCount)

        stepNames(1) = "1. 견적 제출 요청"
        stepDescriptions(1) = "추정가격이 소액인 경우 1개 업체에 견적서 제출을 요청할 수 있습니다. 교복 사양서를 함께 제공합니다."
        stepFormIDs(1) = "F-008,F-012": stepFormLabels(1) = "F-008 교복사양서,F-012 계약특수조건"

        stepNames(2) = "2. 계약 체결"
        stepDescriptions(2) = "제출된 견적을 기준으로 계약을 체결합니다. 계약 특수조건(F-012)을 첨부합니다."
        stepFormIDs(2) = "F-043,F-031": stepFormLabels(2) = "F-043 계약체결,F-031 교복품목별금액표"

        stepNames(3) = "3. 계약 이행·검사검수"
        stepDescriptions(3) = "공급자가 치수를 측정하고 교복을 제작·납품하면, 학교가 규격서 일치 여부를 검사·검수합니다."
        stepFormIDs(3) = "": stepFormLabels(3) = ""

        stepNames(4) = "4. 대가 지급"
        stepDescriptions(4) = "검사·검수 완료 후 대금을 지급합니다."
        stepFormIDs(4) = "": stepFormLabels(4) = ""
    End If

    Me.lstSteps.Clear
    Dim i As Long
    For i = 1 To stepCount
        Me.lstSteps.AddItem stepNames(i)
    Next i
    If stepCount > 0 Then Me.lstSteps.ListIndex = 0
End Sub

Private Sub lstSteps_Click()
    Call 단계표시(Me.lstSteps.ListIndex + 1)
End Sub

Private Sub 단계표시(ByVal idx As Long)
    If idx < 1 Or idx > stepCount Then Exit Sub
    Me.txtDesc.Value = stepDescriptions(idx)
    Me.lstForms.Clear
    If Trim(stepFormLabels(idx)) <> "" Then
        Dim labels() As String
        labels = Split(stepFormLabels(idx), ",")
        Dim j As Long, spacePos As Long
        For j = 0 To UBound(labels)
            ' 2026-09-24 재디자인: "F-009 기초금액계약방법결정" 형태를 코드·서식명 2열로
            ' 나눠 시안의 표 형태(코드/서식명 컬럼)에 근접시킨다.
            spacePos = InStr(1, labels(j), " ")
            If spacePos > 0 Then
                Me.lstForms.AddItem Trim(Left(labels(j), spacePos - 1))
                Me.lstForms.List(Me.lstForms.ListCount - 1, 1) = Trim(Mid(labels(j), spacePos + 1))
            Else
                Me.lstForms.AddItem labels(j)
            End If
        Next j
        ' 목록에서 항목을 직접 클릭해 선택하지 않으면 미리보기/PDF 버튼이 조용히 안내
        ' 메시지만 띄우고 아무 반응이 없는 것처럼 보임(2026-09-21 실사용 중 확인). 첫
        ' 서식을 자동 선택해 바로 미리보기/PDF를 누를 수 있게 함.
        If Me.lstForms.ListCount > 0 Then Me.lstForms.ListIndex = 0
    End If
End Sub

Private Function 선택서식ID() As String
    선택서식ID = ""
    If Me.lstForms.ListIndex < 0 Then Exit Function
    Dim idx As Long
    idx = Me.lstSteps.ListIndex + 1
    If idx < 1 Or idx > stepCount Then Exit Function
    If Trim(stepFormIDs(idx)) = "" Then Exit Function
    Dim ids() As String
    ids = Split(stepFormIDs(idx), ",")
    If Me.lstForms.ListIndex > UBound(ids) Then Exit Function
    선택서식ID = Trim(ids(Me.lstForms.ListIndex))
End Function

Private Sub btnPreview_Click()
    Dim fid As String
    fid = 선택서식ID()
    If fid = "" Then
        MsgBox "먼저 오른쪽 목록에서 관련 서식을 선택하세요.", vbExclamation
        Exit Sub
    End If
    ' 인쇄 미리보기(PrintPreview)는 Excel 화면 전체를 차지하는 백스테이지 보기로 전환됨.
    ' 팝업을 띄운 채로 부르면 두 화면이 서로 앞에 나서려다 멈추는 것을 확인해(2026-09-21),
    ' 미리보기 직전에 팝업을 숨기고 돌아온 뒤 다시 보여줌.
    Me.Hide
    Call 계약절차_서식_미리보기(fid)
    Me.Show 0
End Sub

Private Sub btnPDF_Click()
    Dim fid As String
    fid = 선택서식ID()
    If fid = "" Then
        MsgBox "먼저 오른쪽 목록에서 관련 서식을 선택하세요.", vbExclamation
        Exit Sub
    End If
    Me.Hide
    Call 계약절차_서식_PDF(fid)
    Me.Show 0
End Sub

Private Sub btnHWPX_Click()
    Dim fid As String
    fid = 선택서식ID()
    If fid = "" Then
        MsgBox "먼저 오른쪽 목록에서 관련 서식을 선택하세요.", vbExclamation
        Exit Sub
    End If
    Me.Hide
    Call 계약절차_서식_HWPX(fid)
    Me.Show 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'@
    $codeFrm = $codeFrm -replace "`r`n", "`r" -replace "`n", "`r"
    $frmComp.CodeModule.AddFromString($codeFrm)
    L "frmContractGuide UserForm 추가 완료 (컨트롤 13개, 계약유형별 테마색 재디자인 2026-09-24, 줄 수: $($frmComp.CodeModule.CountOfLines))"
    }

    # ---- UserForm: frmLoadPicker (불러오기 레코드 선택 팝업, 2026-09-24 사용자 요청) ----
    # 기존에는 상단 I2 드롭다운을 먼저 선택한 뒤 [불러오기] 버튼을 눌러야 했는데, 버튼을
    # 누르는 즉시 저장된 레코드 목록이 팝업으로 뜨는 방식을 원한다는 요청에 따라 신설함.
    # I2 드롭다운·불러오기_드롭다운파싱은 그대로 유지하고(회귀 없음), 이 폼은 같은 라벨
    # 목록(DB!U열)을 읽어 같은 파싱·불러오기 로직을 재사용한다.
    $existingLoadFrm = $null
    foreach ($c in $vbproj.VBComponents) {
        if ($c.Name -eq "frmLoadPicker") { $existingLoadFrm = $c; break }
    }
    if ($existingLoadFrm) {
        L "기존 frmLoadPicker UserForm 유지"
    } else {
    $loadFrmComp = $vbproj.VBComponents.Add(3)  # vbext_ct_MSForm
    $loadFrmComp.Name = "frmLoadPicker"
    $loadFrmDesigner = $loadFrmComp.Designer
    $loadFrmComp.Properties("Width").Value = 420
    $loadFrmComp.Properties("Height").Value = 340
    $loadFrmComp.Properties("Caption").Value = "저장된 레코드 불러오기"

    $lblLoadTitle = $loadFrmDesigner.Controls.Add("Forms.Label.1")
    $lblLoadTitle.Name = "lblLoadTitle"
    $lblLoadTitle.Left = 10; $lblLoadTitle.Top = 8; $lblLoadTitle.Width = 390; $lblLoadTitle.Height = 16
    $lblLoadTitle.Caption = "불러올 레코드를 선택하세요(더블클릭도 가능)."

    $lstRecords = $loadFrmDesigner.Controls.Add("Forms.ListBox.1")
    $lstRecords.Name = "lstRecords"
    $lstRecords.Left = 10; $lstRecords.Top = 28; $lstRecords.Width = 400; $lstRecords.Height = 230

    $btnLoad = $loadFrmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnLoad.Name = "btnLoad"
    $btnLoad.Left = 210; $btnLoad.Top = 268; $btnLoad.Width = 90; $btnLoad.Height = 26
    $btnLoad.Caption = "불러오기"

    $btnCancelLoad = $loadFrmDesigner.Controls.Add("Forms.CommandButton.1")
    $btnCancelLoad.Name = "btnCancelLoad"
    $btnCancelLoad.Left = 310; $btnCancelLoad.Top = 268; $btnCancelLoad.Width = 90; $btnCancelLoad.Height = 26
    $btnCancelLoad.Caption = "취소"

    $codeLoadFrm = @'
Option Explicit

Public Sub 초기화()
    Dim wsDB As Worksheet, r As Long, lastRow As Long
    Set wsDB = Sheets("DB")
    Me.lstRecords.Clear
    lastRow = wsDB.Cells(wsDB.Rows.Count, 21).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(wsDB.Cells(r, 21).Value & "") <> "" Then
            Me.lstRecords.AddItem wsDB.Cells(r, 21).Value
        End If
    Next r
    If Me.lstRecords.ListCount = 0 Then
        MsgBox "저장된 레코드가 없습니다. 먼저 [저장하기]로 레코드를 만드세요.", vbInformation
    Else
        Me.lstRecords.ListIndex = Me.lstRecords.ListCount - 1
    End If
End Sub

Private Sub 확정()
    If Me.lstRecords.ListIndex < 0 Then
        MsgBox "불러올 레코드를 목록에서 선택하세요.", vbExclamation
        Exit Sub
    End If
    Dim seq As Long
    seq = 검증_불러오기_드롭다운파싱(Me.lstRecords.Value)
    If seq = 0 Then
        MsgBox "레코드 형식을 확인할 수 없습니다.", vbExclamation
        Exit Sub
    End If
    Unload Me
    Call 불러오기_레코드선택(seq)
End Sub

Private Sub btnLoad_Click()
    Call 확정
End Sub

Private Sub lstRecords_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call 확정
End Sub

Private Sub btnCancelLoad_Click()
    Unload Me
End Sub
'@
    $codeLoadFrm = $codeLoadFrm -replace "`r`n", "`r" -replace "`n", "`r"
    $loadFrmComp.CodeModule.AddFromString($codeLoadFrm)
    L "frmLoadPicker UserForm 추가 완료 (컨트롤 4개, 줄 수: $($loadFrmComp.CodeModule.CountOfLines))"
    }

    try { $wb.Save(); L "DIAG_CHECKPOINT3B_SAVE_OK" } catch { L "DIAG_CHECKPOINT3B_SAVE_FAIL: $($_.Exception.Message)" }

    # ---- ThisWorkbook: Workbook_Open (서식선택 체크박스 초기화) ----
    # 참고: 한글 Office에서는 ThisWorkbook 문서모듈의 기본 컴포넌트 이름이 "ThisWorkbook"이 아니라
    # 로컬라이즈된 이름(예: 이_통합_문서)으로 생성됨. Type=100(문서모듈)이면서 워크시트 코드네임
    # 패턴(Sheet숫자)이 아닌 컴포넌트를 찾아 식별함.
    $thisWb = $null
    foreach ($c in $vbproj.VBComponents) {
        if ($c.Type -eq 100 -and $c.Name -notmatch '^Sheet\d+$') {
            $thisWb = $c
            break
        }
    }
    if (-not $thisWb) { throw "ThisWorkbook 문서모듈을 찾지 못함" }
    L "ThisWorkbook 문서모듈 식별: $($thisWb.Name)"

    # 재실행 대비: 기존 통합문서 이벤트를 모두 지운 뒤 한 번씩만 재추가한다.
    # Workbook_BeforeSave를 남긴 채 추가하면 동일 이벤트 이름이 중복되어 VBA 컴파일이
    # 중단되는 실제 결함이 있었으므로, 각 이벤트를 찾을 수 없을 때까지 반복 제거한다.
    $cm = $thisWb.CodeModule
    foreach ($procName in @("Workbook_Open", "Workbook_BeforeSave")) {
        $removedCount = 0
        while ($true) {
            try {
                $startLine = $cm.ProcStartLine($procName, 0)
                $procCount = $cm.ProcCountLines($procName, 0)
                $cm.DeleteLines($startLine, $procCount)
                $removedCount++
            } catch { break }
        }
        L "$procName 기존 프로시저 $removedCount 개 제거 후 재생성"
    }

    $codeThisWb = @'

Private Sub Workbook_Open()
    Dim wsSel As Worksheet
    Set wsSel = Sheets("서식선택_출력")
    Dim lastRow As Long
    lastRow = wsSel.Cells(wsSel.Rows.Count, 2).End(xlUp).Row
    Dim r As Long
    For r = 5 To lastRow
        wsSel.Cells(r, 1).Value = False
    Next r

    ' UserInterfaceOnly 보호 설정은 파일을 다시 열면 유지되지 않으므로, 매크로 허용 후 다시 적용한다.
    ' 비밀번호는 사용하지 않으며 일반 사용자의 직접 편집만 제한한다.
    Dim internalName As Variant
    For Each internalName In Array("DB", "DB_품목", "DB_업체", "DB_위원", "DB_평가", "DB_정량평가", "DB_정성평가", "DB_자기평점")
        Sheets(CStr(internalName)).Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Next internalName
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim reason As String
    ' 2026-09-24 사용자 결정: 완성도(품목·업체·위원·평가 등) 문제는 더 이상 저장을 막지
    ' 않고 경고만 표시함. 개인정보 보호(F-050 원자료 비저장)만은 예외로 계속 차단함.
    If Not 저장경계_치명적검증(reason) Then
        Cancel = True
        If Application.Visible And Application.UserControl Then
            MsgBox "저장이 취소되었습니다. " & reason, vbExclamation
        End If
        Exit Sub
    End If
    reason = 검증_저장경계_메시지()
    If reason <> "" And Application.Visible And Application.UserControl Then
        MsgBox "일부 데이터가 불완전하지만 저장은 계속 진행됩니다." & vbCrLf & reason, vbExclamation
    End If
End Sub
'@
    $codeThisWb = $codeThisWb -replace "`r`n", "`r" -replace "`n", "`r"
    $thisWb.CodeModule.AddFromString($codeThisWb)
    L "ThisWorkbook 이벤트(Workbook_Open, Workbook_BeforeSave) 각 1개 추가 완료"
    try { $wb.Save(); L "DIAG_CHECKPOINT4_SAVE_OK" } catch { L "DIAG_CHECKPOINT4_SAVE_FAIL: $($_.Exception.Message)" }

    # ---- Form 컨트롤 버튼 배치 (재실행 대비: 동일 이름 기존 버튼 제거) ----
    $wsIn = $wb.Worksheets.Item("기초자료입력")
    foreach ($nm in @("btn초기화","btn저장하기","btn수정하기","btn불러오기")) {
        try { $wsIn.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $btnDefs = @(
        @{name="btn초기화"; caption="초기화"; macro="초기화"; left=520; top=10},
        @{name="btn저장하기"; caption="저장하기"; macro="저장하기"; left=610; top=10},
        @{name="btn수정하기"; caption="수정하기"; macro="수정하기"; left=700; top=10},
        @{name="btn불러오기"; caption="불러오기"; macro="불러오기"; left=790; top=10}
    )
    foreach ($b in $btnDefs) {
        $btn = $wsIn.Buttons().Add($b.left, $b.top, 80, 24)
        $btn.Caption = $b.caption
        $btn.OnAction = $b.macro
        $btn.Name = $b.name
    }
    L "기초자료입력 버튼 4개 배치 완료"

    # 사용설명서 시작 화면: 사용자가 제공한 시안 이미지를 통합문서에 내장하고, 이미지의
    # 세 단계 카드 전체를 클릭 영역으로 쓴다. 별도 버튼을 얹어 시안을 가리지 않는다.
    $wsGuide = $wb.Worksheets.Item("사용설명서")
    foreach ($nm in @("btnGuideMethod", "btnGuideInput", "btnGuideOutput", "imgGuideBanner", "hotGuideMethod", "hotGuideInput", "hotGuideOutput")) {
        try { $wsGuide.Shapes($nm).Delete() } catch { L "기존 사용설명서 요소 $nm 없음" }
    }
    $guideImagePath = Join-Path $root '첫화면\시안-4-사용설명서.png'
    if (Test-Path $guideImagePath) {
        $guidePic = $wsGuide.Shapes.AddPicture((Resolve-Path $guideImagePath).Path, $false, $true, 5, 5, 850, 477)
        $guidePic.Name = "imgGuideBanner"
        # 2026-09-24 사용자 실사용 확인: 이미지가 아예 안 보이는 버그가 있었음. 원인은
        # Placement 기본값(1=이동+크기조정, 셀에 종속)이라 A1:H14 Clear() 이후 행 높이 등
        # 셀 치수에 따라 이미지가 사실상 보이지 않게 축소됨(계약방법안내 배너는 이미
        # Placement=3으로 이 문제를 피해 갔음, PDF 내보내기로 빈 화면임을 실측 확인).
        # 자유배치(3=xlFreeFloating)로 고정해 셀 변경과 무관하게 항상 표시되게 한다.
        $guidePic.Placement = 3
        $guideHotspots = @(
            @{name="hotGuideMethod"; macro="이동_계약방법안내"; left=31; top=127; width=257; height=305},
            @{name="hotGuideInput"; macro="이동_기초자료입력"; left=299; top=127; width=269; height=305},
            @{name="hotGuideOutput"; macro="이동_서식선택출력"; left=567; top=127; width=260; height=305}
        )
        foreach ($hot in $guideHotspots) {
            $shape = $wsGuide.Shapes.AddShape(1, $hot.left, $hot.top, $hot.width, $hot.height)
            $shape.Name = $hot.name; $shape.OnAction = $hot.macro
            $shape.Fill.Transparency = 1; $shape.Line.Visible = 0
            $shape.Placement = 3
        }
        # 2026-09-24 PDF 육안 확인 중 발견(사용자 미보고, 항목 6 검증 과정에서 자체 발견):
        # 시안 이미지가 가로 850pt(약 30cm)인데 시트에 인쇄 설정이 전혀 없어 기본값(A4
        # 세로)로 내보내면 STEP 03 카드 전체와 STEP 02 카드 절반이 잘려 나간다(PDF 실측
        # 확인). 가로 방향 + 한 페이지 맞춤으로 이미지 전체가 한 장에 보이게 한다.
        $wsGuide.PageSetup.Orientation = 2  # xlLandscape
        $wsGuide.PageSetup.Zoom = $false
        $wsGuide.PageSetup.FitToPagesWide = 1
        $wsGuide.PageSetup.FitToPagesTall = 1
        $wsGuide.PageSetup.PrintArea = "`$A`$1:`$T`$28"
        L "사용설명서 시안 이미지 및 3개 단계 클릭 영역 배치 완료 (가로 인쇄·한 페이지 맞춤 포함)"
    } else {
        throw "사용설명서 시안 이미지를 찾지 못했습니다: $guideImagePath"
    }

    $wsSel = $wb.Worksheets.Item("서식선택_출력")
    # 서식선택_출력은 행별 작업 화면으로 정리했다. 일괄 출력은 기초자료입력의
    # 빠른 출력 선택 패널에서만 제공하므로, 이 시트의 중복 상단 버튼은 제거한다.
    foreach ($nm in @("btn인쇄미리보기","btnPDF저장","btnHWPX작성")) {
        try { $wsSel.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    L "서식선택_출력의 중복 상단 일괄 출력 버튼 제거 완료"

    # 행별 미리보기/PDF/HWPX 버튼: 구현상태='Y'인 행에만 배치(재실행 대비 기존 버튼 우선 제거)
    # HWPX 열(2026-09-24 사용자 요청)은 미리보기·PDF와 같은 방식(Application.Caller로 행
    # 추론)을 그대로 재사용한다.
    for ($si = $wsSel.Shapes.Count; $si -ge 1; $si--) {
        $shpName = $wsSel.Shapes.Item($si).Name
        if ($shpName -like "btnPrev_*" -or $shpName -like "btnPdf_*" -or $shpName -like "btnHwpx_*" -or $shpName -like "btnMove_*") {
            $wsSel.Shapes.Item($si).Delete()
        }
    }
    $selLastRow = $wsSel.Cells($wsSel.Rows.Count, 2).End(-4162).Row
    $gLeft = $wsSel.Columns.Item(7).Left
    $hLeft = $wsSel.Columns.Item(8).Left
    $iLeft = $wsSel.Columns.Item(9).Left
    $jLeft = $wsSel.Columns.Item(10).Left
    $rowBtnCount = 0
    for ($r = 5; $r -le $selLastRow; $r++) {
        if ($wsSel.Cells($r, 5).Value2 -eq "Y") {
            $formId = $wsSel.Cells($r, 2).Value2
            $rowTop = $wsSel.Rows.Item($r).Top
            $rowHeight = $wsSel.Rows.Item($r).Height
            $btnTop = $rowTop + (($rowHeight - 15) / 2)

            $bp = $wsSel.Buttons().Add($gLeft + 1, $btnTop, 62, 15)
            $bp.Caption = "미리보기"
            $bp.OnAction = "행별_미리보기"
            $bp.Name = "btnPrev_$formId"

            $bd = $wsSel.Buttons().Add($hLeft + 1, $btnTop, 62, 15)
            $bd.Caption = "PDF"
            $bd.OnAction = "행별_PDF저장"
            $bd.Name = "btnPdf_$formId"

            $bh = $wsSel.Buttons().Add($iLeft + 1, $btnTop, 62, 15)
            $bh.Caption = "HWPX"
            $bh.OnAction = "행별_HWPX작성"
            $bh.Name = "btnHwpx_$formId"

            # "해당시트이동" 버튼(2026-09-24 사용자 요청, 12번 항목): 미리보기/PDF/HWPX
            # 버튼 바로 다음(J열)에 배치해 출력을 거치지 않고 서식 화면으로 바로 이동한다.
            $bm = $wsSel.Buttons().Add($jLeft + 1, $btnTop, 62, 15)
            $bm.Caption = "이동"
            $bm.OnAction = "행별_시트이동"
            $bm.Name = "btnMove_$formId"
            $rowBtnCount++

            # 해당 서식 시트 자체에도 목록으로 돌아가는 버튼 2개를 배치함(2026-09-21 사용자
            # 요청). 인쇄영역과 절대 겹치지 않도록 PrintObject=False로 인쇄 제외 처리함.
            $connectedSheetName = $wsSel.Cells($r, 6).Value2
            if ($connectedSheetName -ne "") {
                try {
                    $wsForm = $wb.Worksheets.Item($connectedSheetName)
                    # 2026-09-24 재조사(F-049는 버튼이 있는데 F-050만 없음): F-050은 "인쇄 전용
                    # 빈 설문지 보호"를 위해 구조 빌드 단계에서 이미
                    # Protect("", DrawingObjects:=True, ...)로 보호돼 있다(2983행). 이 VBA 빌드
                    # 단계는 보호가 걸린 뒤에 실행되는데, DrawingObjects=True인 시트는
                    # Buttons().Add()가 COM 예외를 던지고, 그 예외가 바로 아래 catch에서
                    # 조용히 로그만 남기고 삼켜져 F-050에 이동 버튼 2개가 실제로 빠졌다(COM
                    # 실측 확인). 보호돼 있으면 버튼을 넣는 동안만 잠깐 해제했다가 원래
                    # 보호 설정(F-050과 동일한 인자)으로 되돌린다.
                    $wasProtected = $wsForm.ProtectContents -or $wsForm.ProtectDrawingObjects
                    if ($wasProtected) { $wsForm.Unprotect("") }
                    foreach ($nm in @("btn목록_서식선택", "btn목록_계약방법")) {
                        try { $wsForm.Buttons($nm).Delete() } catch { $null = $_ }
                    }
                    $navTop = $wsForm.Rows.Item(1).Top
                    $navLeft = $wsForm.Columns.Item(14).Left  # N열, 모든 서식 인쇄영역(최대 L열) 밖
                    $navBtn1 = $wsForm.Buttons().Add($navLeft, $navTop, 100, 18)
                    $navBtn1.Caption = "서식목록으로"
                    $navBtn1.OnAction = "이동_서식선택출력"
                    $navBtn1.Name = "btn목록_서식선택"
                    $navBtn1.PrintObject = $false

                    $navBtn2 = $wsForm.Buttons().Add($navLeft + 105, $navTop, 100, 18)
                    $navBtn2.Caption = "계약방법안내로"
                    $navBtn2.OnAction = "이동_계약방법안내"
                    $navBtn2.Name = "btn목록_계약방법"
                    $navBtn2.PrintObject = $false
                    if ($wasProtected) { $wsForm.Protect("", $true, $true, $true, $true) }
                } catch {
                    L "경고: $connectedSheetName 시트를 찾지 못해 이동 버튼을 배치하지 못함"
                }
            }
        }
    }
    L "서식선택_출력 행별 버튼 배치 완료 (대상 행: $rowBtnCount, 각 서식 시트 이동 버튼 포함)"

    # 기초자료입력 "빠른 출력 선택" 패널(2026-09-21): 체크박스는 서식선택_출력!A열에 LinkedCell로
    # 연결해 두 시트가 같은 값을 실시간 공유하고(양방향), 보기 버튼은 구현상태='Y'인 행에만 배치함.
    foreach ($nm in @("btn빠른미리보기","btn빠른PDF","btn빠른HWPX")) {
        try { $wsIn.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    for ($si = $wsIn.Shapes.Count; $si -ge 1; $si--) {
        $shpName = $wsIn.Shapes.Item($si).Name
        if ($shpName -like "chkQuick_*" -or $shpName -like "btnQuickView_*") {
            $wsIn.Shapes.Item($si).Delete()
        }
    }
    $tLeft = $wsIn.Columns.Item(20).Left   # T열
    $wLeft = $wsIn.Columns.Item(23).Left   # W열

    # 기존 초기화/저장하기/수정하기/불러오기 버튼(좌표 520~870, top=10)과 완전히 겹쳐서
    # 그 버튼들을 가려버렸던 배치 실수를 발견함(2026-09-21 사용자 보고). 패널 영역(T열)
    # 기준 위치로 옮겨 더 이상 겹치지 않게 함.
    $quickBtnTop = $wsIn.Rows.Item(2).Top
    $quickBtn1 = $wsIn.Buttons().Add($tLeft, $quickBtnTop, 160, 24)
    $quickBtn1.Caption = "선택 서식 인쇄 미리보기"
    $quickBtn1.OnAction = "선택서식_인쇄미리보기"
    $quickBtn1.Name = "btn빠른미리보기"

    $quickBtn2 = $wsIn.Buttons().Add($tLeft + 170, $quickBtnTop, 160, 24)
    $quickBtn2.Caption = "선택 서식 PDF 저장"
    $quickBtn2.OnAction = "선택서식_PDF저장"
    $quickBtn2.Name = "btn빠른PDF"

    $quickBtn3 = $wsIn.Buttons().Add($tLeft + 340, $quickBtnTop, 150, 24)
    $quickBtn3.Caption = "선택 서식 HWPX 작성"
    $quickBtn3.OnAction = "선택서식_HWPX작성"
    $quickBtn3.Name = "btn빠른HWPX"

    # 기초자료입력에서 계약방법안내로 바로 이동하는 버튼(2026-09-21 사용자 요청)
    try { $wsIn.Buttons("btn이동_계약방법안내").Delete() } catch { L "기존 버튼 btn이동_계약방법안내 없음" }
    $btnGoMethod = $wsIn.Buttons().Add($tLeft + 500, $quickBtnTop, 150, 24)
    $btnGoMethod.Caption = "계약방법안내로"
    $btnGoMethod.OnAction = "이동_계약방법안내"
    $btnGoMethod.Name = "btn이동_계약방법안내"

    # DB 시트 관리자용 열람 버튼(요청 4, 2026-09-23). 같은 행 오른쪽에 배치하며, 실수로 누르지
    # 않도록 캡션에 "관리자용"임을 명시하고 클릭 시 매크로 안에서 다시 확인 메시지를 띄운다.
    try { $wsIn.Buttons("btn이동_DB열기").Delete() } catch { L "기존 버튼 btn이동_DB열기 없음" }
    $btnGoDb = $wsIn.Buttons().Add($tLeft + 660, $quickBtnTop, 150, 24)
    $btnGoDb.Caption = "DB 열기(관리자용)"
    $btnGoDb.OnAction = "DB시트열기"
    $btnGoDb.Name = "btn이동_DB열기"

    $quickCount = 0
    for ($r = 5; $r -le $selLastRow; $r++) {
        $qFormId = $wsSel.Cells($r, 2).Value2
        $qRowTop = $wsIn.Rows.Item($r).Top
        $qRowHeight = $wsIn.Rows.Item($r).Height

        $chk = $wsIn.CheckBoxes().Add($tLeft + 2, $qRowTop + 1, 14, 14)
        $chk.Caption = ""
        $chk.LinkedCell = "'서식선택_출력'!`$A`$$r"
        $chk.Name = "chkQuick_$qFormId"

        if ($wsSel.Cells($r, 5).Value2 -eq "Y") {
            $qBtnTop = $qRowTop + (($qRowHeight - 15) / 2)
            $btnView = $wsIn.Buttons().Add($wLeft + 1, $qBtnTop, 62, 15)
            $btnView.Caption = "보기"
            $btnView.OnAction = "기초자료_보기_실행"
            $btnView.Name = "btnQuickView_$qFormId"
        }
        $quickCount++
    }
    L "기초자료입력 빠른 출력 선택 패널 배치 완료 (체크박스 $quickCount 개, 보기 버튼은 구현된 서식만)"

    $wsSearch = $wb.Worksheets.Item("학교검색")
    foreach ($nm in @("btn학교검색", "btn선택학교반영")) {
        try { $wsSearch.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $btnSearch = $wsSearch.Buttons().Add(520, 35, 80, 24)
    $btnSearch.Caption = "검색"
    $btnSearch.OnAction = "학교검색_실행"
    $btnSearch.Name = "btn학교검색"
    $btnApplySchool = $wsSearch.Buttons().Add(610, 35, 190, 24)
    $btnApplySchool.Caption = "선택 학교를 기초자료에 반영"
    $btnApplySchool.OnAction = "선택학교_기초자료반영"
    $btnApplySchool.Name = "btn선택학교반영"
    L "학교검색 버튼 2개 배치 완료"

    $wsFlow = $wb.Worksheets.Item("입찰계약절차안내")
    try { $wsFlow.Buttons("btn관련FormID이동").Delete() } catch { L "기존 버튼 btn관련FormID이동 없음" }
    $btnFlow = $wsFlow.Buttons().Add(650, 35, 150, 24)
    $btnFlow.Caption = "관련 Form ID로 이동"
    $btnFlow.OnAction = "관련FormID로이동"
    $btnFlow.Name = "btn관련FormID이동"
    L "절차안내 Form ID 이동 버튼 배치 완료"

    $wsMethod = $wb.Worksheets.Item("계약방법안내")

    # 2026-09-24 사용자 요청(2차): 색을 흉내낸 도형 카드 대신, 제공받은 시안 이미지
    # (첫화면/시안1-1-수정.svg를 미리 PNG로 렌더링한 것)를 1행 컨테이너에 그대로 삽입하고,
    # 그 위에 카드 4개와 정확히 겹치는 "투명 클릭 영역"만 얹는다. 클릭 영역은 폼 컨트롤
    # Buttons가 아니라 일반 도형(AddShape)으로 만들어야 Fill/Line을 완전히 투명하게 지울 수
    # 있다(폼 컨트롤은 OS가 그리는 위젯이라 완전한 투명 처리가 보장되지 않음). Name·OnAction은
    # 기존과 동일하게 유지해 다른 매크로(계약방식선택_1인견적 등)와의 연결은 그대로 재사용한다.
    try { $wsMethod.Shapes("imgContractBanner").Delete() } catch { L "기존 계약방법안내 배너 이미지 없음" }
    $bannerImagePath = Join-Path $root 'artifacts\excel\_assets\contract_banner.png'
    # picLeft/picTop/cardScale은 배너 유무와 무관하게 버튼 3개 배치(아래)에서도 그대로
    # 재사용하므로 if(Test-Path) 밖에서 먼저 정의한다.
    $picLeft = 8; $picTop = 5; $picWidth = 900; $picHeight = 540
    $cardScale = $picWidth / 1200.0
    if (Test-Path $bannerImagePath) {
        $bannerImagePath = (Resolve-Path $bannerImagePath).Path
        $pic = $wsMethod.Shapes.AddPicture($bannerImagePath, $false, $true, $picLeft, $picTop, $picWidth, $picHeight)
        $pic.Name = "imgContractBanner"
        $pic.Placement = 3  # xlFreeFloating

        foreach ($nm in @("btn계약절차_1인견적","btn계약절차_2인견적","btn계약절차_2단계입찰","btn계약절차_일반입찰")) {
            try { $wsMethod.Shapes($nm).Delete() } catch { L "기존 카드 클릭 영역 $nm 없음" }
        }
        # SVG 원본 viewBox(1200x720) 좌표를 그림 표시 크기(900x540, 배율 0.75)로 환산해
        # 카드 4개(각 252x160, y=372)의 위치와 정확히 겹치는 투명 사각형을 올린다.
        $cardTop = $picTop + (372 * $cardScale)
        $cardWidth = 252 * $cardScale
        $cardHeight = 160 * $cardScale
        $cardDefs = @(
            @{name="btn계약절차_1인견적"; macro="계약방식선택_1인견적"; x=60},
            @{name="btn계약절차_2인견적"; macro="계약방식선택_2인견적"; x=336},
            @{name="btn계약절차_2단계입찰"; macro="계약방식선택_2단계입찰"; x=612},
            @{name="btn계약절차_일반입찰"; macro="계약방식선택_일반입찰"; x=888}
        )
        foreach ($def in $cardDefs) {
            $cardLeft = $picLeft + ($def.x * $cardScale)
            $hotspot = $wsMethod.Shapes.AddShape(1, $cardLeft, $cardTop, $cardWidth, $cardHeight)  # 1 = msoShapeRectangle
            $hotspot.Name = $def.name
            $hotspot.Fill.Visible = 0
            $hotspot.Line.Visible = 0
            $hotspot.OnAction = $def.macro
            $hotspot.Placement = 3
        }
        L "계약방법안내 시안 이미지 삽입 및 카드 4개 투명 클릭 영역 배치 완료"
    } else {
        L "경고: 계약방법안내 배너 이미지($bannerImagePath)를 찾지 못해 이미지 삽입을 건너뜀"
    }
    # 제공 시안 전체(4개 카드)가 가로로 한눈에 보이도록 계약방법안내는 화면·PDF 모두
    # 가로 1쪽 기준으로 맞춘다. 아래의 확인 표는 같은 폭 안에서 계속 읽을 수 있다.
    $methodPs = $wsMethod.PageSetup
    $methodPs.Orientation = 2; $methodPs.Zoom = $false; $methodPs.FitToPagesWide = 1; $methodPs.FitToPagesTall = $false
    $methodPs.PrintArea = "`$A`$1:`$C`$10"

    # 2026-09-24 사용자 요청(3차): 반영 버튼(구 3행)·이동 버튼 2개(구 16행)가 배너 이미지와
    # 멀리 떨어진 하단에 흩어져 있어 찾기 어렵다는 지적을 받아, 배너 상단 제목 "교복구매
    # 계약서식 자동화" 오른쪽·로고 왼쪽의 빈 공간(배너 PNG 실측: SVG 1200x720 기준 제목은
    # x~420에서 끝나고 로고는 x~912에서 시작)으로 버튼 3개를 모두 옮긴다. 카드 4개(y=372~)
    # 보다 위쪽이라 겹치지 않는다.
    # 2026-09-24 사용자 요청(4차): "계약방법 반영" 버튼을 삭제함 — 카드 4개(계약방식선택_*)를
    # 클릭하면 계약방식즉시반영()이 B-03(기초자료입력!C14)에 이미 즉시 반영하므로(2026-09-21
    # 카드 UX 개편 이후 이 버튼은 과거 B5/B6/B7 2단계 확인 설계의 잔재로 기능이 중복됨), 남은
    # 이동 버튼 2개만 배치하고 옛 반영 버튼이 있던 자리까지 당겨 채운다. 연결된 VBA 매크로
    # (계약방법안내_기초자료반영 등)는 검증_계약방법안내_기초자료반영이 계속 사용하므로 유지한다.
    foreach ($nm in @("btn계약방법반영","btn이동_기초자료입력","btn이동_서식선택출력")) {
        try { $wsMethod.Buttons($nm).Delete() } catch { L "기존 버튼 $nm 없음" }
    }
    $headerBtnTop = $picTop + (36 * $cardScale)
    $headerBtnHeight = 22
    $headerBtnGap = 8
    $headerBtnLeft1 = $picLeft + (440 * $cardScale)

    $btnGoIn = $wsMethod.Buttons().Add($headerBtnLeft1, $headerBtnTop, 110, $headerBtnHeight)
    $btnGoIn.Caption = "기초자료입력으로"
    $btnGoIn.OnAction = "이동_기초자료입력"
    $btnGoIn.Name = "btn이동_기초자료입력"

    $headerBtnLeft2 = $headerBtnLeft1 + 110 + $headerBtnGap
    $btnGoSel = $wsMethod.Buttons().Add($headerBtnLeft2, $headerBtnTop, 110, $headerBtnHeight)
    $btnGoSel.Caption = "서식선택_출력으로"
    $btnGoSel.OnAction = "이동_서식선택출력"
    $btnGoSel.Name = "btn이동_서식선택출력"
    L "계약방법안내 헤더 버튼 2개(기초자료입력으로·서식선택_출력으로) 배치 완료 (상단 제목 옆, 계약방법 반영 버튼은 카드 즉시반영과 기능 중복이라 삭제)"

    # DB 시트 자체에 "닫기(다시 숨김)" 버튼 배치(요청 4, 2026-09-23). 1차 시도(보이게+활성화만)도
    # 같은 "Buttons 클래스의 Add 속성을 가져올 수 없습니다" COM 예외로 재차 실패함을 실측으로
    # 확인함(2026-09-23). 재조사 결과 진짜 원인은 ActiveSheet 여부가 아니라 build_excel_v1_structure.ps1이
    # DB 시트를 `Protect($password, $DrawingObjects=$true, ...)`로 보호해 두어(내부 DB 보호 절 참고),
    # DrawingObjects=True 보호가 Shape/Buttons 계열 개체 삽입 자체를 차단하기 때문임. 버튼 배치 직전에만
    # Unprotect하고, 배치가 끝나면 같은 비밀번호 없음·DrawingObjects=True 조합으로 즉시 재보호한다.
    # 헤더(1~23열)와 겹치지 않도록 25열(Y) 오른쪽에 배치한다.
    $wsDbHome = $wb.Worksheets.Item("DB")
    $wsDbHome.Visible = -1  # xlSheetVisible (버튼 배치를 위해 일시적으로만 보이게 함)
    $wsDbHome.Unprotect("")
    $wsDbHome.Activate()
    try { $wsDbHome.Buttons("btnDB닫기").Delete() } catch { L "기존 버튼 btnDB닫기 없음" }
    $btnDbClose = $wsDbHome.Buttons().Add($wsDbHome.Columns.Item(25).Left, $wsDbHome.Rows.Item(1).Top, 170, 24)
    $btnDbClose.Caption = "DB 닫기(다시 숨김)"
    $btnDbClose.OnAction = "DB시트닫기"
    $btnDbClose.Name = "btnDB닫기"
    $wsDbHome.Protect("", $true, $true, $true, $true)
    $wsDbHome.Visible = 0  # xlSheetHidden (배치 완료 후 원래 보호 상태로 복원)
    $wsIn.Activate()
    L "DB 시트 닫기 버튼 배치 완료(재보호+Hidden 상태로 복원)"

    # ---- 저장 ----
    # 2026-09-19 진단: Save()가 일반 COMException으로 실패해 SaveAs(같은 경로, 52)로
    # 바꿔봤으나, 이미 열려 있는 파일을 자기 자신에 SaveAs로 덮어쓰는 시도가 오히려 더
    # 구체적인 "저장할 수 없습니다" 잠금류 오류로 실패함(자기 자신을 여는 새 쓰기 핸들과
    # 기존 열림 핸들이 충돌하는 것으로 추정). Save()는 이 코드베이스 전체 이력에서 수백 회
    # 성공한 방식이므로 되돌리되, 이 프로젝트에서 반복 관찰된 일시적 COM/훅 레이스 현상에
    # 대비해 짧은 재시도(지수 백오프 + GC)를 추가한다.
    $saveAttempt = 0
    $saveLastError = $null
    while ($saveAttempt -lt 3 -and -not $saveSucceeded) {
        $saveAttempt++
        try {
            $wb.Save()
            $saveSucceeded = $true
        } catch {
            $saveLastError = $_
            L "저장 시도 $saveAttempt/3 실패: $($_.Exception.Message)"
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds ([Math]::Pow(2, $saveAttempt))
        }
    }
    if (-not $saveSucceeded) { throw $saveLastError }
    L "VBA 매크로 및 버튼 추가 후 저장 완료(SaveAs): $targetPath"

} finally {
    # 정리 단계 자체의 COM 오류(예: 앞선 오류로 Excel 프로세스가 이미 비정상 상태가 된 경우의
    # Close() 실패)가 try 블록의 원래 예외를 가리거나 로그 flush·프로세스 강제 종료를 막지
    # 않도록 각 정리 호출을 개별적으로 감싼다.
    if ($wb) {
        try { $wb.Close($saveSucceeded) } catch { L "Workbook Close 실패(정리 계속 진행): $($_.Exception.Message)" }
    }
    try { $excel.Quit() } catch { L "Excel Quit 실패(정리 계속 진행): $($_.Exception.Message)" }
    if ($wb) {
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { L "Workbook ReleaseComObject 실패(정리 계속 진행): $($_.Exception.Message)" }
    }
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { L "Excel ReleaseComObject 실패(정리 계속 진행): $($_.Exception.Message)" }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($excelProcessId) {
        Start-Sleep -Milliseconds 500
        if (Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue) {
            Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
            L "PID $excelProcessId Excel 프로세스가 Quit() 이후에도 남아 있어 강제 종료함"
        }
    }
    # try 블록에서 예외가 발생했더라도(원본 예외는 finally 종료 후 그대로 재전파됨) 여기까지
    # 수집된 로그는 항상 디스크에 남겨 원인 진단이 가능하게 한다.
    [System.IO.File]::WriteAllText($logPath, $log.ToString(), [System.Text.UTF8Encoding]::new($false))
}
