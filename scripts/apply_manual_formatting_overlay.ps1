<#
사용자가 Excel에서 직접 조정한 서식(셀 병합·행높이·열너비·세로 중앙정렬 등)을
매 안전 빌드마다 자동으로 재현하기 위한 서식 오버레이 스크립트.

기준본(SourcePath, 기본값 _workspace/03_excel/manual_formatting_source_v14.xlsm — 사용자가
여러 서식에 걸쳐 셀병합·행높이·정렬을 수동 편집해 저장한 v14.xlsm의 영구 보존 사본)의
시트별 병합 범위·행높이·열너비·정렬·줄바꿈을 읽어, 대상 파일(TargetPath — 새로 빌드된
스테이징 xlsm)에 그대로 적용한다. 셀 값·수식·VBA는 전혀 건드리지 않는다.

run_build_excel_v1_structure_utf8.ps1 파이프라인에서 build_excel_v1_vba.ps1 실행 직후,
verify_excel_v1.ps1 실행 직전에 호출된다.
#>
param(
    [string]$SourcePath = '_workspace/03_excel/manual_formatting_source_v14.xlsm',
    [Parameter(Mandatory)][string]$TargetPath,
    [string[]]$SkipSheets = @()
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

$wbSrc = $null
$wbTgt = $null
try {
    $srcPath = (Resolve-Path $SourcePath).Path
    $tgtPath = (Resolve-Path $TargetPath).Path

    $wbSrc = $excel.Workbooks.Open($srcPath, 0, $true)
    $wbTgt = $excel.Workbooks.Open($tgtPath, 0, $false)
    # 파일을 막 연 직후에는 COM 링크가 일시적으로 불안정해 속성 설정이 간헐적으로
    # InvalidCastException을 던지는 경우가 있었음(진단 결과, 동일 파일을 재오픈해
    # 단독 실행하면 100% 성공하는 것으로 확인) - 안정화 대기 + 워밍업 접근으로 회피.
    Start-Sleep -Milliseconds 800
    [void]$wbTgt.Worksheets.Item(1).Cells.Item(1, 1).Value2

    $summary = New-Object System.Collections.Generic.List[string]
    $failCount = 0

    foreach ($wsSrc in $wbSrc.Worksheets) {
        $name = $wsSrc.Name
        if ($SkipSheets -contains $name) { continue }

        $wsTgt = $null
        foreach ($s in $wbTgt.Worksheets) { if ($s.Name -eq $name) { $wsTgt = $s; break } }
        if (-not $wsTgt) { $summary.Add("[스킵] ${name} : 대상 파일에 동일 시트 없음"); continue }

        $wasProtected = $wsTgt.ProtectContents
        $origDrawing = $wsTgt.ProtectDrawingObjects
        $origScenarios = $wsTgt.ProtectScenarios
        if ($wasProtected) {
            try { [void]$wsTgt.Unprotect('') } catch {
                $summary.Add("[스킵] ${name} : 시트 보호 해제 실패(비밀번호 있음?) - $($_.Exception.Message)")
                continue
            }
        }

        $used = $wsSrc.UsedRange
        $startRow = $used.Row
        $startCol = $used.Column
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count

        for ($i = 0; $i -lt $rows; $i++) {
            $r = $startRow + $i
            $h = $wsSrc.Rows.Item($r).RowHeight
            $ok = $false
            for ($attempt = 1; $attempt -le 3 -and -not $ok; $attempt++) {
                try {
                    if ($wsTgt.Rows.Item($r).RowHeight -ne $h) {
                        $wsTgt.Rows.Item($r).RowHeight = $h
                    }
                    $ok = $true
                } catch {
                    if ($attempt -eq 3) {
                        $failCount++
                        $summary.Add("  [경고] ${name} 행${r} 높이 적용 실패(3회 재시도): $($_.Exception.Message)")
                    } else {
                        Start-Sleep -Milliseconds 200
                    }
                }
            }
        }

        for ($j = 0; $j -lt $cols; $j++) {
            $c = $startCol + $j
            $w = $wsSrc.Columns.Item($c).ColumnWidth
            $ok = $false
            for ($attempt = 1; $attempt -le 3 -and -not $ok; $attempt++) {
                try {
                    if ($wsTgt.Columns.Item($c).ColumnWidth -ne $w) {
                        $wsTgt.Columns.Item($c).ColumnWidth = $w
                    }
                    $ok = $true
                } catch {
                    if ($attempt -eq 3) {
                        $failCount++
                        $summary.Add("  [경고] ${name} 열${c} 너비 적용 실패(3회 재시도): $($_.Exception.Message)")
                    } else {
                        Start-Sleep -Milliseconds 200
                    }
                }
            }
        }

        $seenMerge = New-Object System.Collections.Generic.HashSet[string]
        $mergeApplied = 0

        for ($i = 0; $i -lt $rows; $i++) {
            for ($j = 0; $j -lt $cols; $j++) {
                $r = $startRow + $i
                $c = $startCol + $j
                $cellSrc = $wsSrc.Cells.Item($r, $c)

                if ($cellSrc.MergeCells) {
                    $area = $cellSrc.MergeArea
                    $addr = $area.Address($false, $false)
                    if ($seenMerge.Contains($addr)) { continue }
                    [void]$seenMerge.Add($addr)

                    $tgtRange = $wsTgt.Range($addr)
                    try {
                        if (-not $tgtRange.MergeCells) {
                            $tgtRange.Merge() | Out-Null
                            $mergeApplied++
                        }
                        $tgtRange.HorizontalAlignment = $area.HorizontalAlignment
                        $tgtRange.VerticalAlignment = $area.VerticalAlignment
                        $tgtRange.WrapText = $area.WrapText
                    } catch {
                        $failCount++
                        $summary.Add("  [경고] ${name} ${addr} 병합/정렬 적용 실패: $($_.Exception.Message)")
                    }
                } else {
                    try {
                        $tgtCell = $wsTgt.Cells.Item($r, $c)
                        if ($tgtCell.MergeCells) { $tgtCell.UnMerge() }
                        $tgtCell.HorizontalAlignment = $cellSrc.HorizontalAlignment
                        $tgtCell.VerticalAlignment = $cellSrc.VerticalAlignment
                        $tgtCell.WrapText = $cellSrc.WrapText
                    } catch {
                        $failCount++
                        $summary.Add("  [경고] ${name} R${r}C${c} 정렬 적용 실패: $($_.Exception.Message)")
                    }
                }
            }
        }

        if ($wasProtected) {
            try {
                $wsTgt.Protect([Type]::Missing, $origDrawing, $wasProtected, $origScenarios)
            } catch {
                $failCount++
                $summary.Add("  [경고] ${name} 재보호 실패: $($_.Exception.Message)")
            }
        }

        $summary.Add("[$name] 행높이 ${rows}개 · 열너비 ${cols}개 검사, 병합 ${mergeApplied}건 신규 적용")
    }

    $wbTgt.Save()

    foreach ($line in $summary) { Write-Host $line }
    Write-Host "FORMATTING_OVERLAY_DONE: $TargetPath (경고 ${failCount}건)"
} finally {
    if ($wbTgt) { $wbTgt.Close($false) }
    if ($wbSrc) { $wbSrc.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
