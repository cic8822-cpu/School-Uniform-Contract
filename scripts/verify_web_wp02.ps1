param()

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$dataDirectory = Join-Path $root '_workspace\05_web\data'
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-WebCondition {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        $script:failures.Add($Message)
    }
}

function Read-JsonFile {
    param([string]$Name)

    $path = Join-Path $dataDirectory $Name
    Assert-WebCondition (Test-Path -LiteralPath $path -PathType Leaf) "WP-02 데이터 파일 누락: $Name"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        $script:failures.Add("WP-02 JSON 파싱 실패($Name): $($_.Exception.Message)")
        return $null
    }
}

$contractMethods = Read-JsonFile 'contract-methods.json'
$workflows = Read-JsonFile 'workflows.json'
$forms = Read-JsonFile 'forms.json'
$fields = Read-JsonFile 'fields.json'
$templateMappingPath = Join-Path $dataDirectory 'template-mapping.md'
Assert-WebCondition (Test-Path -LiteralPath $templateMappingPath -PathType Leaf) 'WP-02 Template Mapping 누락: template-mapping.md'

if ($null -ne $contractMethods) {
    $methods = @($contractMethods.contractMethods)
    Assert-WebCondition ($methods.Count -eq 4) "계약방법 Master 수 불일치: 기대 4, 실제 $($methods.Count)"
    $expectedMethods = @('CM-01', 'CM-02', 'CM-03', 'CM-04')
    Assert-WebCondition (((@($methods.id) | Sort-Object) -join ',') -eq (($expectedMethods | Sort-Object) -join ',')) '계약방법 Master ID 집합 불일치'
    Assert-WebCondition (@($methods | Where-Object { $_.autoDecision -ne $false }).Count -eq 0) '계약방법 자동판정 설정이 발견됨'
    foreach ($method in $methods) {
        Assert-WebCondition (-not [string]::IsNullOrWhiteSpace([string]$method.workflowId)) "$($method.id) workflowId 누락"
        Assert-WebCondition ($method.stepCount -in @(4, 5, 11)) "$($method.id) 단계 수가 허용 범위와 다름: $($method.stepCount)"
    }
}

if ($null -ne $workflows) {
    $workflowItems = @($workflows.workflows)
    Assert-WebCondition ($workflowItems.Count -eq 3) "Workflow Master 수 불일치: 기대 3, 실제 $($workflowItems.Count)"
    $expectedStepCounts = @{ 'WF-QUOTE1' = 4; 'WF-QUOTE2' = 5; 'WF-BID' = 11 }
    foreach ($workflow in $workflowItems) {
        Assert-WebCondition ($expectedStepCounts.ContainsKey([string]$workflow.workflowId)) "알 수 없는 Workflow ID: $($workflow.workflowId)"
        if ($expectedStepCounts.ContainsKey([string]$workflow.workflowId)) {
            Assert-WebCondition (@($workflow.steps).Count -eq $expectedStepCounts[$workflow.workflowId]) "$($workflow.workflowId) 단계 수 불일치: 기대 $($expectedStepCounts[$workflow.workflowId]), 실제 $(@($workflow.steps).Count)"
        }
        Assert-WebCondition (@($workflow.steps | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.stepName) }).Count -eq 0) "$($workflow.workflowId)에 이름이 없는 단계가 있음"
        $missingStepEvidence = @($workflow.steps | Where-Object {
            @($_.checklistItems).Count -eq 0 -or
            @($_.cautions).Count -eq 0 -or
            @($_.manualReferences).Count -eq 0
        })
        Assert-WebCondition ($missingStepEvidence.Count -eq 0) "$($workflow.workflowId)에 체크리스트·주의사항·근거가 없는 단계가 있음"
        $invalidSourceRefs = @($workflow.steps | ForEach-Object {
            $step = $_
            @($step.checklistItems + $step.cautions | ForEach-Object { @($_.sourceRefs) } | Where-Object {
                [string]::IsNullOrWhiteSpace([string]$_.document) -or
                [string]::IsNullOrWhiteSpace([string]$_.locator) -or
                [string]::IsNullOrWhiteSpace([string]$_.status)
            } | ForEach-Object { "$($step.stepNo):$($_.document)" })
        })
        Assert-WebCondition ($invalidSourceRefs.Count -eq 0) "$($workflow.workflowId)에 위치·상태가 없는 단계 근거가 있음"
        Assert-WebCondition ($null -ne $workflow.evidence -and -not [string]::IsNullOrWhiteSpace([string]$workflow.evidence.evidenceStatus)) "$($workflow.workflowId)의 절차 근거 상태가 누락됨"
        if ($workflow.evidence.evidenceStatus -eq 'excel-guide-only') {
            Assert-WebCondition ($workflow.evidence.displayPolicy -match '매뉴얼 확인 필요' -and @($workflow.steps | Where-Object { @($_.manualReferences | Where-Object { $_.status -eq 'manual-confirmation-required' }).Count -eq 0 }).Count -eq 0) "$($workflow.workflowId)의 매뉴얼 확인 필요 표시 계약이 불완전함"
        }
    }
    $bidWorkflow = @($workflowItems | Where-Object { $_.workflowId -eq 'WF-BID' })[0]
    Assert-WebCondition ($null -ne $bidWorkflow.contractMethodScope -and @($bidWorkflow.contractMethodScope.confirmedContractMethodIds) -contains 'CM-03' -and @($bidWorkflow.contractMethodScope.referenceOnlyContractMethodIds) -contains 'CM-04') '입찰 Workflow의 2단계계약 확정·일반경쟁 참고 범위가 누락됨'
}

if ($null -ne $contractMethods -and $null -ne $workflows) {
    $workflowIds = @($workflows.workflows.workflowId)
    Assert-WebCondition (@($contractMethods.contractMethods | Where-Object { $_.workflowId -notin $workflowIds }).Count -eq 0) '계약방법이 존재하지 않는 Workflow를 참조함'
    $generalBid = @($contractMethods.contractMethods | Where-Object { $_.id -eq 'CM-04' })[0]
    Assert-WebCondition ($null -ne $generalBid.workflowBinding -and $generalBid.workflowBinding.status -eq 'reference-only' -and -not $generalBid.workflowBinding.actionable) '일반경쟁입찰의 공유 Workflow가 확정 절차처럼 노출될 위험이 있음'
}

if ($null -ne $forms) {
    $formItems = @($forms.forms)
    Assert-WebCondition ($formItems.Count -eq 52) "Form Master 수 불일치: 기대 52, 실제 $($formItems.Count)"
    $formIds = @($formItems.formId)
    Assert-WebCondition (($formIds | Select-Object -Unique).Count -eq 52) 'Form ID 중복이 발견됨'
    $expectedFormIds = 1..52 | ForEach-Object { 'F-{0:D3}' -f $_ }
    Assert-WebCondition ((Compare-Object -ReferenceObject $expectedFormIds -DifferenceObject $formIds).Count -eq 0) 'Form ID F-001~F-052 전체가 존재하지 않음'
    Assert-WebCondition (@($formItems | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.manualNo) -or [string]::IsNullOrWhiteSpace([string]$_.title) }).Count -eq 0) '매뉴얼 번호 또는 서식명이 비어 있는 Form이 있음'

    $templateForms = @($formItems | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.hwpxTemplate) })
    Assert-WebCondition ($templateForms.Count -eq 48) "HWPX 템플릿 Form 수 불일치: 기대 48, 실제 $($templateForms.Count)"
    foreach ($form in $templateForms) {
        $templatePath = Join-Path $root ([string]$form.hwpxTemplate)
        Assert-WebCondition (Test-Path -LiteralPath $templatePath -PathType Leaf) "$($form.formId) HWPX 템플릿 파일 누락: $($form.hwpxTemplate)"
    }
}

if ($null -ne $fields) {
    $fieldItems = @($fields.fields)
    Assert-WebCondition ($fieldItems.Count -gt 0) 'Field Master가 비어 있음'
    $fieldIds = @($fieldItems.fieldId)
    Assert-WebCondition (($fieldIds | Select-Object -Unique).Count -eq $fieldItems.Count) 'Field ID 중복이 발견됨'
    Assert-WebCondition (@($fieldItems | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.fieldId) -or [string]::IsNullOrWhiteSpace([string]$_.label) }).Count -eq 0) 'Field ID 또는 한글명이 비어 있는 항목이 있음'

    if ($null -ne $forms) {
        $unknownFields = @($forms.forms | ForEach-Object { @($_.inputFieldIds) } | Where-Object { $_ -and $_ -notin $fieldIds } | Select-Object -Unique)
        Assert-WebCondition ($unknownFields.Count -eq 0) "Form Master가 존재하지 않는 Field를 참조함: $($unknownFields -join ', ')"
        $computedInputOverlaps = @($forms.forms | ForEach-Object {
            $form = $_
            @($form.inputFieldIds | Where-Object { $_ -in @($form.computedFieldIds) } | ForEach-Object { "$($form.formId):$_" })
        })
        Assert-WebCondition ($computedInputOverlaps.Count -eq 0) "계산 Field가 직접입력 Field와 중복됨: $($computedInputOverlaps -join ', ')"
        $invalidCoverage = @($forms.forms | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_.hwpxTemplate) -and (
                $null -eq $_.hwpxCoverage -or
                [string]::IsNullOrWhiteSpace([string]$_.hwpxCoverage.level) -or
                $null -eq $_.hwpxCoverage.supportedFieldIds -or
                [string]::IsNullOrWhiteSpace([string]$_.hwpxCoverage.outputCompletion)
            )
        })
        Assert-WebCondition ($invalidCoverage.Count -eq 0) "HWPX 템플릿 반영 범위가 정의되지 않은 Form이 있음: $($invalidCoverage.formId -join ', ')"
        $unknownScopes = @($forms.forms | Where-Object {
            @($_.contractMethods).Count -eq 0 -and (
                $null -eq $_.contractMethodScope -or
                $_.contractMethodScope.status -ne 'manual-confirmation-required'
            )
        })
        Assert-WebCondition ($unknownScopes.Count -eq 0) "계약방법 미확정 Form의 확인 필요 상태가 누락됨: $($unknownScopes.formId -join ', ')"
        $templateTokenFields = @{ '학교명' = 'C-01'; '학년도' = 'C-02'; '문서번호' = 'C-06'; '발행일' = 'C-05'; '관련문서' = 'D-03'; '제출기한' = 'B-07' }
        foreach ($templateForm in @($forms.forms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.hwpxTemplate) })) {
            $tokenFieldIds = @($templateForm.hwpxTokens | ForEach-Object { $templateTokenFields[[string]$_] } | Where-Object { $_ } | Select-Object -Unique)
            $declaredFieldIds = @($templateForm.inputFieldIds + $templateForm.referencedFieldIds + $templateForm.hwpxRequiredFieldIds | Select-Object -Unique)
            $missingTokenContracts = @($tokenFieldIds | Where-Object { $_ -notin $declaredFieldIds })
            Assert-WebCondition ($missingTokenContracts.Count -eq 0) "$($templateForm.formId) HWPX 토큰 Field 계약 누락: $($missingTokenContracts -join ', ')"
            $coverageDifference = @(Compare-Object -ReferenceObject $tokenFieldIds -DifferenceObject @($templateForm.hwpxCoverage.supportedFieldIds))
            Assert-WebCondition ($coverageDifference.Count -eq 0) "$($templateForm.formId) HWPX 지원 Field 범위가 실제 토큰과 다름"
        }
    }

    $protectedFormIds = @($fields.protectedForms.formId)
    Assert-WebCondition ((Compare-Object -ReferenceObject @('F-029', 'F-047', 'F-050') -DifferenceObject $protectedFormIds).Count -eq 0) '개인정보 가능 서식 보호 정책 대상이 F-029/F-047/F-050과 일치하지 않음'
    foreach ($protectedForm in @($fields.protectedForms)) {
        Assert-WebCondition (@($protectedForm.forbiddenFields).Count -gt 0) "$($protectedForm.formId) 금지 필드 목록이 비어 있음"
        Assert-WebCondition (@($protectedForm.forbiddenFields | Where-Object { -not $_.inputUiForbidden -or -not $_.storeForbidden -or -not $_.tokenForbidden -or -not $_.networkForbidden }).Count -eq 0) "$($protectedForm.formId) 보호 필드 경계가 완전하지 않음"
    }
    $satisfactionSurveyPolicy = @($fields.protectedForms | Where-Object { $_.formId -eq 'F-050' })[0]
    Assert-WebCondition ($null -ne $satisfactionSurveyPolicy -and (@($satisfactionSurveyPolicy.allowedFieldIds) -join ',') -eq 'C-01,C-02,D-02') 'F-050 허용 Field가 공통 업무값 범위를 벗어남'
    if ($null -ne $forms) {
        $satisfactionSurveyForm = @($forms.forms | Where-Object { $_.formId -eq 'F-050' })[0]
        Assert-WebCondition ($null -ne $satisfactionSurveyForm -and 'D-05' -notin @($satisfactionSurveyForm.inputFieldIds) -and 'D-05' -notin @($satisfactionSurveyForm.referencedFieldIds)) 'F-050에 자유서술 Field D-05가 다시 연결됨'
    }
}

$tokenScriptPath = Join-Path $root 'scripts\hwpx_token_fill.ps1'
Assert-WebCondition (Test-Path -LiteralPath $tokenScriptPath -PathType Leaf) 'HWPX 원본 토큰 검증 스크립트 누락'
if ((Test-Path -LiteralPath $tokenScriptPath -PathType Leaf) -and (Test-Path -LiteralPath $templateMappingPath -PathType Leaf)) {
    $tokenScript = Get-Content -LiteralPath $tokenScriptPath -Raw -Encoding UTF8
    $mapping = Get-Content -LiteralPath $templateMappingPath -Raw -Encoding UTF8
    $allowedTokenMatch = [regex]::Match($tokenScript, '\$AllowedTokens\s*=\s*@\((.*?)\)', [Text.RegularExpressions.RegexOptions]::Singleline)
    Assert-WebCondition $allowedTokenMatch.Success '원본 HWPX 허용 토큰 목록을 읽지 못함'
    if ($allowedTokenMatch.Success) {
        $tokenBlock = $allowedTokenMatch.Groups.Item(1).Value
        $tokenParts = $tokenBlock.Split([char]39)
        $tokens = @()
        for ($index = 1; $index -lt $tokenParts.Length; $index += 2) {
            $tokens += $tokenParts[$index]
        }
        foreach ($token in $tokens) {
            Assert-WebCondition ($mapping.Contains($token)) "Template Mapping에 허용 토큰 누락: $token"
        }
    }
    Assert-WebCondition ($mapping -match [regex]::Escape('성명|이름|담당자|전화|휴대|주소|이메일|서명|직인|계좌|주민|생년|학번|설문')) 'Template Mapping의 금지 토큰 패턴이 원본과 다름'
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { [Console]::Error.WriteLine("FAIL: $_") }
    exit 1
}

Write-Output 'PASS: WP-02 마스터 데이터 검증 전체 통과'
