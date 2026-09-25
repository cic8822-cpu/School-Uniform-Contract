<#
v14.xlsm(사용자가 여러 서식에서 셀병합·행높이·세로중앙정렬을 수동 편집해 저장한 파일)의
서식(병합 범위·행높이·열너비·정렬·줄바꿈)을 시트별로 읽어, 대상 파일(TargetPath, 최신
자동 빌드본의 사본)에 그대로 적용한다. 셀 값·수식·VBA는 전혀 건드리지 않는다.
Excel COM을 사용하며, 이미 대상과 값이 같은 병합/서식은 재적용을 건너뛴다.
#>
param(
    [Parameter(Mandatory)][string]$SourcePath,
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

    $summary = New-Object System.Collections.Generic.List[string]

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
            try { $wsTgt.Unprotect('') } catch {
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
            if ($wsTgt.Rows.Item($r).RowHeight -ne $h) {
                $wsTgt.Rows.Item($r).RowHeight = $h
            }
        }

        for ($j = 0; $j -lt $cols; $j++) {
            $c = $startCol + $j
            $w = $wsSrc.Columns.Item($c).ColumnWidth
            if ($wsTgt.Columns.Item($c).ColumnWidth -ne $w) {
                $wsTgt.Columns.Item($c).ColumnWidth = $w
            }
        }

        $seenMerge = New-Object System.Collections.Generic.HashSet[string]
        $mergeApplied = 0
        $alignApplied = 0

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
                        $alignApplied++
                    } catch {
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
                        $summary.Add("  [경고] ${name} R${r}C${c} 정렬 적용 실패: $($_.Exception.Message)")
                    }
                }
            }
        }

        if ($wasProtected) {
            $wsTgt.Protect($null, $origDrawing, $wasProtected, $origScenarios)
        }

        $summary.Add("[$name] 행높이 ${rows}개 · 열너비 ${cols}개 검사, 병합 ${mergeApplied}건 신규 적용")
    }

    $wbTgt.Save()

    foreach ($line in $summary) { Write-Host $line }
    Write-Host "OVERLAY_DONE: $TargetPath"
} finally {
    if ($wbTgt) { $wbTgt.Close($false) }
    if ($wbSrc) { $wbSrc.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
