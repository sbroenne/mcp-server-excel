#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Runs the native API coverage workflow through the real CLI executable.
.DESCRIPTION
    Retains the expanded native operation and save/reopen assertions. The CLI
    acceptance test hosts this workflow with a hard deadline and its own pipe.
#>
[CmdletBinding()]
param(
    [switch]$KeepFile,
    [string]$PipeName
)
$ErrorActionPreference = 'Stop'

# Find CLI executable (prefer Release build)
$cliPath = Join-Path $PSScriptRoot "..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"
if (-not (Test-Path $cliPath)) {
    $cliPath = Join-Path $PSScriptRoot "..\src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe"
}
if (-not (Test-Path $cliPath)) {
    Write-Error "CLI not found. Build first: dotnet build src/ExcelMcp.CLI"
    exit 1
}

$cli = (Resolve-Path $cliPath).Path
Write-Host "Using CLI: $cli" -ForegroundColor Cyan

$previousPipeName = $env:EXCELMCP_CLI_PIPE
$selectedPipeName = if (-not [string]::IsNullOrWhiteSpace($PipeName)) {
    $PipeName
}
else {
    "excelmcp-cli-workflow-$PID-$([Guid]::NewGuid().ToString('N'))"
}
$env:EXCELMCP_CLI_PIPE = $selectedPipeName
Write-Host "Using private CLI pipe: $selectedPipeName" -ForegroundColor DarkGray

function Reset-CliWorkflowEnvironment {
    $cleanupExitCode = 0
    try {
        & (Join-Path $PSScriptRoot 'Stop-ExcelMcpProcesses.ps1') -PipeName $selectedPipeName
        $cleanupExitCode = $LASTEXITCODE
    }
    finally {
        if ($null -eq $previousPipeName) {
            Remove-Item Env:EXCELMCP_CLI_PIPE -ErrorAction SilentlyContinue
        }
        else {
            $env:EXCELMCP_CLI_PIPE = $previousPipeName
        }
    }

    return $cleanupExitCode
}

# Generate unique test file
$testFile = Join-Path $env:TEMP "cli-workflow-test-$(Get-Random).xlsx"
$chartImagePath = [IO.Path]::ChangeExtension($testFile, '.png')
Write-Host "Test file: $testFile" -ForegroundColor Cyan

$passed = 0
$failed = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Action,
        [scriptblock]$Verify = $null
    )

    Write-Host "`n[$Name]" -ForegroundColor Yellow
    try {
        $global:LASTEXITCODE = 0
        $result = & $Action
        if ($LASTEXITCODE -ne 0) { throw "Command failed with exit code $LASTEXITCODE. $($result | ConvertTo-Json -Depth 10)" }
        if ($result.success -eq $false -or $result.errorMessage) { throw "Command reported failure: $($result | ConvertTo-Json -Depth 10)" }
        if ($Verify) {
            $verifyResult = & $Verify $result
            if (-not $verifyResult) {
                Write-Host "  FAIL: Verification failed" -ForegroundColor Red
                Write-Host "  Result: $result" -ForegroundColor Gray
                $script:failed++
                return $null
            }
        }
        Write-Host "  PASS" -ForegroundColor Green
        $script:passed++
        return $result
    }
    catch {
        Write-Host "  FAIL: $_" -ForegroundColor Red
        $script:failed++
        return $null
    }
}

# ============================================================================
# TEST WORKFLOW
# ============================================================================

try {
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Excel CLI Workflow Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

:workflow do {
# 1. Create session (auto-starts daemon, creates file)
$session = Test-Step "Create session (create file)" {
    & $cli -q session create $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -eq $true
}

if (-not $session.sessionId) {
    Write-Host "`nFATAL: Could not open session. Aborting." -ForegroundColor Red
    $failed++
    break workflow
}

$sessionId = $session.sessionId
Write-Host "  Session ID: $sessionId" -ForegroundColor Gray

# 2. Create worksheet (simpler than set-values with JSON)
Test-Step "Create worksheet 'Data'" {
    & $cli -q sheet create --session $sessionId --sheet-name Data | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 3. List worksheets
$sheets = Test-Step "List worksheets" {
    & $cli -q sheet list --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.worksheets -ne $null
}

Write-Host "  Sheets: $(($sheets.worksheets | Measure-Object).Count)" -ForegroundColor Gray

# 4. Format ranges (multi-value --range-addresses exercises string[] CLI option)
Test-Step "Format ranges on 'Data' (multi-value addresses)" {
    & $cli -q rangeformat format --session $sessionId --sheet-name Data --range-addresses "A1:A2" --range-addresses "C1:C2" --format-options '{"bold":true,"fillColor":"#FFFF00"}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

Test-Step "Apply native theme colors and a diagonal border" {
    & $cli -q rangeformat format --session $sessionId --sheet-name Data --range-addresses "AA10:AB11" --format-options '{"fontThemeColor":5,"fillThemeColor":6,"indentLevel":2,"horizontalAlignment":"left","borders":[{"position":"DiagonalUp","lineStyle":"dash","color":"#123456"}]}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

Test-Step "Read native theme formatting and selected border" {
    & $cli -q rangeformat get-format --session $sessionId --sheet-name Data --range-address "AA10" | ConvertFrom-Json
} -Verify {
    param($r)
    $cell = $r.cells[0].stored
    $border = $cell.borders | Where-Object edge -eq xlDiagonalUp
    $r.success -eq $true -and $cell.font.color.themeColor -eq 5 -and
    $cell.fill.color.themeColor -eq 6 -and $cell.indentLevel -eq 2 -and
    $border.lineStyle -eq -4115 -and $border.color.rgb -eq '#123456'
}

Test-Step "Capture a reusable native cell style" {
    & $cli -q workbook create-cell-style --session $sessionId --style-name WorkflowHighlight --source-sheet-name Data --source-cell-address AA10 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.style.name -eq 'WorkflowHighlight' -and $r.style.builtIn -eq $false
}

Test-Step "Apply the custom cell style" {
    & $cli -q rangeformat set-style --session $sessionId --sheet-name Data --range-address AD10 --style-name WorkflowHighlight | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

Test-Step "Update the native custom definition" {
    & $cli -q workbook update-cell-style --session $sessionId --style-name WorkflowHighlight --style-options '{"includeFont":true,"formatOptions":{"bold":true}}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.style.format.font.bold -eq $true
}

Test-Step "Read the existing custom style user" {
    & $cli -q rangeformat get-style --session $sessionId --sheet-name Data --range-address AD10 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.styleName -eq 'WorkflowHighlight' -and $r.isBuiltInStyle -eq $false
}

Test-Step "Read the complete custom style definition" {
    & $cli -q workbook get-cell-style --session $sessionId --style-name WorkflowHighlight | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.style.format.font.bold -eq $true -and $r.style.format.borders.Count -eq 6
}

Test-Step "Delete the custom cell style" {
    & $cli -q workbook delete-cell-style --session $sessionId --style-name WorkflowHighlight | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

Test-Step "Clone a native table style" {
    & $cli -q workbook create-table-style --session $sessionId --style-name WorkflowTableStyle --source-style-name TableStyleMedium2 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.style.builtIn -eq $false -and $r.style.elements.Count -eq 43
}

Test-Step "Update a native table-style element" {
    & $cli -q workbook update-table-style --session $sessionId --style-name WorkflowTableStyle --table-style-options '{"elements":[{"elementType":"xlHeaderRow","fillColor":"#123456","bold":false}]}' | ConvertFrom-Json
} -Verify {
    param($r)
    $header = $r.style.elements | Where-Object elementType -eq xlHeaderRow
    $r.success -eq $true -and $header.fill.color.rgb -eq '#123456'
}

Test-Step "Read the complete table-style definition" {
    & $cli -q workbook get-table-style --session $sessionId --style-name WorkflowTableStyle | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.style.elements.Count -eq 43
}

Test-Step "List every native table style" {
    & $cli -q workbook list-table-styles --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and ($r.styles | Where-Object name -eq WorkflowTableStyle).builtIn -eq $false
}

Test-Step "Delete the custom table style" {
    & $cli -q workbook delete-table-style --session $sessionId --style-name WorkflowTableStyle | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 5. Add a conditional-format rule with typed integer/boolean arguments
Test-Step "Read all stored and displayed formatting" {
    & $cli -q rangeformat get-format --session $sessionId --sheet-name Data --range-address "A1:C2" --view both | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.cellCount -eq 6 -and
    $r.cells.Count -eq 6 -and
    $r.cells[0].stored.font.bold -eq $true -and
    $r.cells[0].stored.fill.color.rgb -eq '#FFFF00' -and
    $r.cells[0].displayed.fill.color.rgb -eq '#FFFF00'
}

Test-Step "Add typed conditional-format rule" {
    & $cli -q conditionalformat add-rule --session $sessionId --sheet-name Data --range-address "B1:B10" --rule-type top10 --rank 7 --top10-percent true --font-bold true --font-italic false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

$conditionalFormatRules = Test-Step "Inspect typed conditional-format rule" {
    & $cli -q conditionalformat list-rules --session $sessionId --sheet-name Data --range-address "B1:B10" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.rules.Count -eq 1 -and
    $r.rules[0].top10.rank -eq 7 -and
    $r.rules[0].top10.percent -eq $true
}

$updatedConditionalRules = Test-Step "Update only the selected conditional rule" {
    $selected = $conditionalFormatRules.rules[0]
    & $cli -q conditionalformat update-rule --session $sessionId --sheet-name Data --rule-priority $selected.priority --expected-fingerprint $selected.fingerprint --options '{"rank":5,"stopIfTrue":false,"appliesTo":"B1:B8"}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.rules.Count -eq 1 -and
    $r.rules[0].top10.rank -eq 5 -and $r.rules[0].stopIfTrue -eq $false -and
    $r.rules[0].appliesTo -eq '$B$1:$B$8'
}
Test-Step "Add a separate disposable conditional rule" {
    & $cli -q conditionalformat add-rule --session $sessionId --sheet-name Data --range-address W10:W12 --rule-type expression --formula1 '=TRUE()' --stop-if-true false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
$rulesBeforeMove = Test-Step "Read fresh worksheet rule selections" {
    & $cli -q conditionalformat list-worksheet-rules --session $sessionId --sheet-name Data | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.rules.Count -eq 2
}
$movedRules = Test-Step "Move only the disposable rule to first priority" {
    $selected = @($rulesBeforeMove.rules | Where-Object type -EQ expression)[0]
    & $cli -q conditionalformat set-rule-priority --session $sessionId --sheet-name Data --rule-priority $selected.priority --expected-fingerprint $selected.fingerprint --new-priority 1 | ConvertFrom-Json
} -Verify {
    param($r)
    $selected = @($r.rules | Where-Object type -EQ expression)[0]
    $r.success -eq $true -and $r.rules.Count -eq 2 -and $selected.priority -eq 1
}
Test-Step "Delete only the disposable rule with its fresh selection" {
    $selected = @($movedRules.rules | Where-Object type -EQ expression)[0]
    & $cli -q conditionalformat delete-rule --session $sessionId --sheet-name Data --rule-priority $selected.priority --expected-fingerprint $selected.fingerprint | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.rules.Count -eq 1 -and $r.rules[0].top10.rank -eq 5
}

# Keep Data for save/reopen verification; deletion uses a separate sheet.
Test-Step "Write persisted value" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[424242]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Protected write rejects occupied destination without mutation" {
    $rejected = & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[1]]' | ConvertFrom-Json
    if ($LASTEXITCODE -ne 1 -or $rejected.success -ne $false -or $rejected.errorCategory -ne 'Conflict' -or $rejected.errorMessage -notmatch '\$A\$1') {
        throw "Expected a categorized occupied-cell failure: $($rejected | ConvertTo-Json -Depth 10)"
    }
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address A1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -eq 424242
}
Test-Step "Explicit allow permits intentional replacement" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address A1 --values '[[424242]]' --overwrite-policy allow | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Discover constants in the exact requested range" {
    & $cli -q range get-special-cells --session $sessionId --sheet-name Data --range-address "A1:A3" --cell-kind constants | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.sheetName -eq 'Data' -and
    $r.rangeAddress -eq '$A$1:$A$3' -and
    $r.cellKind -eq 'constants' -and
    $r.cellCount -eq 1 -and
    $r.areas.Count -eq 1 -and
    $r.areas[0] -eq '$A$1'
}
Test-Step "Write content for formatting-only paste" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address H1 --values '[[7]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Copy formats without replacing occupied content" {
    & $cli -q range copy --session $sessionId --source-sheet Data --source-range "A1:C2" --target-sheet Data --target-range H1 --paste-kind formats | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.destinationAddress -eq '$H$1:$J$2' -and
    $r.pasteKind -eq 'formats'
}
Test-Step "Verify formatting-only paste preserved content" {
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address H1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -eq 7
}
Test-Step "Verify copied font and fill" {
    & $cli -q rangeformat get-format --session $sessionId --sheet-name Data --range-address H1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.cells[0].stored.font.bold -eq $true -and
    $r.cells[0].stored.fill.color.rgb -eq '#FFFF00'
}
Test-Step "Write dynamic-array source" {
    & $cli -q range set-formulas --session $sessionId --sheet-name Data --range-address "F1" --formulas '[["=SEQUENCE(3)"]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native dynamic-array source and all results" {
    & $cli -q range get-spill-info --session $sessionId --sheet-name Data --range-address "F1:F3" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.capability -eq 'supported' -and
    $r.cellCount -eq 3 -and
    $r.cells[0].state -eq 'source' -and
    $r.cells[2].state -eq 'result' -and
    $r.cells[2].sourceAddress -eq '$F$1' -and
    $r.cells[2].spillAddress -eq '$F$1:$F$3'
}
Test-Step "Write native pattern seeds" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address "P1:P2" --values '[[1],[3]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Extend native pattern" {
    & $cli -q rangeedit auto-fill --session $sessionId --sheet-name Data --source-range "P1:P2" --destination-range "P1:P4" --fill-type series | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native pattern result" {
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address P4 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -eq 7
}
Test-Step "Write relative R1C1 formula" {
    & $cli -q range set-formulas --session $sessionId --sheet-name Data --range-address Q1 --formulas '[["=RC[-1]*2"]]' --reference-style r1c1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Fill native relative formulas" {
    & $cli -q rangeedit fill --session $sessionId --sheet-name Data --range-address "Q1:Q4" --direction down | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read R1C1 formula and calculated result" {
    & $cli -q range get-formulas --session $sessionId --sheet-name Data --range-address Q4 --reference-style r1c1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.formulas[0][0] -eq '=RC[-1]*2' -and $r.values[0][0] -eq 14
}
Test-Step "Trace complete native local precedents" {
    & $cli -q range trace-precedents --session $sessionId --sheet-name Data --range-address Q4 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.nodes.Count -eq 2 -and
    $r.edges.Count -eq 1 -and
    $r.coverage.workbookComplete -eq $false -and
    $r.coverage.nativeTraversalComplete -eq $true -and
    $r.unresolved.Count -eq 0
}
Test-Step "Trace native dependents with explicit unresolved leaf coverage" {
    & $cli -q range trace-dependents --session $sessionId --sheet-name Data --range-address P4 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and
    $r.nodes.Count -eq 2 -and
    $r.edges.Count -eq 1 -and
    $r.coverage.workbookComplete -eq $false -and
    $r.coverage.nativeTraversalComplete -eq $false -and
    $r.unresolved.Count -eq 1
}
Test-Step "Write PivotTable calculation source" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address U1:V3 --values '[["Region","Sales"],["North",100],["South",300]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Create calculation PivotTable" {
    & $cli -q pivottable create-from-range --session $sessionId --source-sheet Data --source-range U1:V3 --destination-sheet Data --destination-cell X1 --pivot-table-name CalculationPivot | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Add calculation PivotTable row field" {
    & $cli -q pivottablefield add-row-field --session $sessionId --pivot-table-name CalculationPivot --field-name Region | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Add separately named calculation value field" {
    & $cli -q pivottablefield add-value-field --session $sessionId --pivot-table-name CalculationPivot --field-name Sales --custom-name "Total Sales" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Set native percentage of grand total" {
    & $cli -q pivottablefield set-field-calculation --session $sessionId --pivot-table-name CalculationPivot --field-name "Total Sales" --calculation PercentOfTotal | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.calculation -eq 'PercentOfTotal' -and $r.function -eq 'Sum'
}
Test-Step "Verify native percentage results" {
    & $cli -q pivottablecalc get-data --session $sessionId --pivot-table-name CalculationPivot | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[1][1] -eq 0.25 -and $r.values[2][1] -eq 0.75
}
Test-Step "Set native Pivot style and repeated labels" {
    & $cli -q pivottablecalc set-layout-options --session $sessionId --pivot-table-name CalculationPivot --layout-options '{"rowLayout":1,"repeatLabels":true,"styleName":"PivotStyleMedium9","preserveFormatting":true}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.styleName -eq 'PivotStyleMedium9' -and $r.rowFields[0].repeatLabels -eq $true
}
Test-Step "Add native Pivot label filter" {
    & $cli -q pivottablefield add-field-filter --session $sessionId --pivot-table-name CalculationPivot --field-name Region --filter-options '{"type":"CaptionEquals","text1":"North"}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.filters[0].value1 -eq 'North'
}
Test-Step "Read every native Pivot calculated filter" {
    & $cli -q pivottablefield get-field-filters --session $sessionId --pivot-table-name CalculationPivot --field-name Region | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.filters.Count -eq 1 -and $r.filters[0].type -eq 'CaptionEquals'
}
Test-Step "Clear only selected Pivot calculated filters" {
    & $cli -q pivottablefield clear-field-filters --session $sessionId --pivot-table-name CalculationPivot --field-name Region | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.filters.Count -eq 0
}
Test-Step "Read native Pivot source and shared users" {
    & $cli -q pivottable get-source --session $sessionId --pivot-table-name CalculationPivot | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.recordCount -eq 2 -and $r.sharedPivotTables[0] -eq 'CalculationPivot'
}
Test-Step "Replace only selected Pivot cache without rebuilding fields" {
    & $cli -q pivottable set-source --session $sessionId --pivot-table-name CalculationPivot --source-sheet-name Data --source-range-address U1:V3 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.recordCount -eq 2 -and $r.connectedSlicerCaches.Count -eq 0
}
Test-Step "Create first drawing object" {
    & $cli -q drawing add-shape --session $sessionId --sheet-name Data --name DrawFirst --left 20 --top 20 --width 40 --height 30 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObject.name -eq 'DrawFirst'
}
Test-Step "Create second drawing object" {
    & $cli -q drawing add-shape --session $sessionId --sheet-name Data --name DrawSecond --left 100 --top 70 --width 40 --height 30 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObject.name -eq 'DrawSecond'
}
Test-Step "Duplicate drawing with actual point offsets" {
    & $cli -q drawing duplicate-object --session $sessionId --sheet-name Data --object-name DrawFirst --new-name DrawThird --offset-left 180 --offset-top 70 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects[0].name -eq 'DrawThird' -and $r.drawingObjects[0].left -eq 200
}
Test-Step "Group named drawings" {
    & $cli -q drawing group-objects --session $sessionId --sheet-name Data --object-names '["DrawFirst","DrawSecond"]' --group-name DrawGroup | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects[0].name -eq 'DrawGroup' -and $r.drawingObjects[0].children.Count -eq 2
}
Test-Step "Read complete grouped drawing members" {
    & $cli -q drawing get-object --session $sessionId --sheet-name Data --object-name DrawGroup | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObject.children.Count -eq 2
}
Test-Step "Ungroup and expose actual member names" {
    & $cli -q drawing ungroup-object --session $sessionId --sheet-name Data --object-name DrawGroup | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects.Count -eq 2 -and $r.drawingObjects.name -contains 'DrawFirst'
}
Test-Step "Align selected drawing edges" {
    & $cli -q drawing align-objects --session $sessionId --sheet-name Data --object-names '["DrawFirst","DrawSecond","DrawThird"]' --alignment Top | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects.Count -eq 3 -and ($r.drawingObjects.top | Select-Object -Unique).Count -eq 1
}
Test-Step "Distribute native drawing gaps" {
    & $cli -q drawing distribute-objects --session $sessionId --sheet-name Data --object-names '["DrawFirst","DrawSecond","DrawThird"]' --distribution Horizontal | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects[1].left -eq 110
}
Test-Step "Send drawing behind other objects" {
    & $cli -q drawing set-z-order --session $sessionId --sheet-name Data --object-name DrawThird --z-order SendToBack | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.drawingObjects[0].zOrderPosition -eq 1
}
Test-Step "Write chart depth source" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address AA1:AC4 --values '[["Category","First","Second"],["A",10,100],["B",20,200],["C",30,300]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Create native chart" {
    & $cli -q chart create-from-range --session $sessionId --sheet-name Data --source-range-address AA1:AC4 --chart-type ColumnClustered --chart-name DepthChart | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.chartName -eq 'DepthChart'
}
Test-Step "Set native combo series type" {
    & $cli -q chartconfig set-series-chart-type --session $sessionId --chart-name DepthChart --series-index 2 --chart-type LineMarkers | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Assign secondary chart axes" {
    & $cli -q chartconfig set-series-axis-group --session $sessionId --chart-name DepthChart --series-index 2 --axis-group Secondary | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.axisGroup -eq 'Secondary'
}
Test-Step "Read actual combo series settings" {
    & $cli -q chartconfig get-series-settings --session $sessionId --chart-name DepthChart --series-index 2 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.chartType -eq 'LineMarkers' -and $r.axisGroup -eq 'Secondary' -and $r.pointCount -eq 2 -and $r.name -eq 'B'
}
Test-Step "Format only one chart point" {
    & $cli -q chartconfig set-point-format --session $sessionId --chart-name DepthChart --series-index 1 --point-index 2 --point-options '{"fillColor":"#FF0000","lineColor":"#0000FF","lineWeight":2}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.fillColor -eq '#FF0000'
}
Test-Step "Set native fixed error bars" {
    & $cli -q chartconfig set-error-bars --session $sessionId --chart-name DepthChart --series-index 1 --error-bar-options '{"kind":"Fixed","amount":2,"endStyle":"NoCap"}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.hasErrorBars -eq $true -and $r.endStyle -eq 'NoCap'
}
Test-Step "Read honest native error-bar limits" {
    & $cli -q chartconfig get-error-bars --session $sessionId --chart-name DepthChart --series-index 1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.hasErrorBars -eq $true -and $r.settingsReadable -eq $false
}
Test-Step "Export an actual chart image" {
    & $cli -q chart export-image --session $sessionId --chart-name DepthChart --target-path $chartImagePath | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and (Get-Item -LiteralPath $chartImagePath).Length -gt 1000
}
Test-Step "Set exact cell protection flags" {
    & $cli -q rangelink set-cell-protection --session $sessionId --sheet-name Data --range-address "S1,S3" --locked false --formula-hidden true | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read complete cell protection without hiding gap state" {
    & $cli -q rangelink get-cell-protection --session $sessionId --sheet-name Data --range-address "S1:S3" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.cells.Count -eq 3 -and
    $r.cells[0].locked -eq $false -and $r.cells[0].formulaHidden -eq $true -and
    $r.cells[1].locked -eq $true -and $r.cells[1].formulaHidden -eq $false
}
Test-Step "Protect worksheet with explicit row-formatting permission" {
    & $cli -q worksheetstyle set-protection --session $sessionId --sheet-name Data --is-protected true --options '{"allowFormattingRows":true}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native worksheet permissions" {
    & $cli -q worksheetstyle get-protection --session $sessionId --sheet-name Data | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.protectContents -eq $true -and
    $r.permissions.allowFormattingRows -eq $true -and $r.permissions.allowSorting -eq $false
}
Test-Step "Unprotect worksheet for remaining workflow" {
    & $cli -q worksheetstyle set-protection --session $sessionId --sheet-name Data --is-protected false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native calculation settings" {
    & $cli -q calculationmode get-settings --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $script:previousCalculationMode = $r.mode
    $r.success -eq $true -and $r.settingsScope -eq 'application' -and
    $r.maximumIterations -gt 0 -and $r.precisionAsDisplayed -eq $false
}
Test-Step "Set manual calculation without changing omitted settings" {
    & $cli -q calculationmode set-settings --session $sessionId --mode manual | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.mode -eq 'manual'
}
Test-Step "Rebuild native formula dependencies" {
    & $cli -q calculationmode calculate --session $sessionId --scope application --kind rebuild | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Restore previous calculation mode" {
    & $cli -q calculationmode set-settings --session $sessionId --mode $script:previousCalculationMode | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.mode -eq $script:previousCalculationMode
}
Test-Step "Keep stored numeric precision" {
    & $cli -q calculationmode set-precision --session $sessionId --precision-as-displayed false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.precisionAsDisplayed -eq $false
}
Test-Step "Inspect only the owned workbook window context" {
    & $cli -q window get-context --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.availability -eq 'available' -and $r.windows.Count -gt 0
}
Test-Step "Read all native workbook theme definitions" {
    & $cli -q workbook get-theme --session $sessionId | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.colors.Count -eq 12 -and $r.majorFonts.Count -eq 3 -and $r.minorFonts.Count -eq 3
}
Test-Step "Hide exact rows without hiding the gap" {
    & $cli -q rangeformat set-visibility --session $sessionId --sheet-name Data --range-address "A40,A42" --axis rows --hidden true | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read exact row visibility and its limits" {
    & $cli -q rangeformat get-visibility --session $sessionId --sheet-name Data --range-address "A40:A42" --axis rows | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.items.Count -eq 3 -and
    $r.items[0].hidden -eq $true -and $r.items[1].hidden -eq $false -and
    $r.items[2].hidden -eq $true -and $r.items[0].hiddenCause -eq 'undetermined'
}
Test-Step "Restore row visibility" {
    & $cli -q rangeformat set-visibility --session $sessionId --sheet-name Data --range-address "A40,A42" --axis rows --hidden false | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Write native series seed" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address R1 --values '[[5]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Create native stepped series" {
    & $cli -q rangeedit create-series --session $sessionId --sheet-name Data --range-address "R1:R4" --orientation columns --step-value 5 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native stepped series result" {
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address R4 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -eq 20
}
Test-Step "Write duplicate records" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address "X20:Y23" --values '[["Key","Amount"],[1,10],[1,20],[2,30]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Remove native duplicates with exact counts" {
    & $cli -q rangeedit remove-duplicates --session $sessionId --sheet-name Data --range-address "X20:Y23" --key-columns '[1]' --has-headers true | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.removedRows -eq 1 -and
    $r.remainingRows -eq 2 -and $r.remainingRange -eq '$X$20:$Y$22'
}
Test-Step "Verify native duplicate removal" {
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address "X21:Y23" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][1] -eq 10 -and
    $r.values[1][1] -eq 30 -and $null -eq $r.values[2][0]
}
Test-Step "Write text parsing input" {
    & $cli -q range set-values --session $sessionId --sheet-name Data --range-address "AB20:AB21" --values '[["001,10,"],["002,20,"]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Split text with exact native output bounds" {
    & $cli -q rangeedit text-to-columns --session $sessionId --sheet-name Data --source-range "AB20:AB21" --destination-cell AD20 --options '{"comma":true,"fields":[{"position":1,"dataType":"Text"}]}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.outputColumns -eq 3 -and
    $r.destinationRange -eq '$AD$20:$AF$21'
}
Test-Step "Verify native text parsing" {
    & $cli -q range get-values --session $sessionId --sheet-name Data --range-address "AD20:AF21" | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.values[0][0] -ceq '001' -and
    $r.values[1][0] -ceq '002' -and $r.values[1][1] -eq 20 -and
    [string]::IsNullOrEmpty($r.values[0][2])
}
Test-Step "Create ordinary filtering worksheet" {
    & $cli -q sheet create --session $sessionId --sheet-name Filters | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Write ordinary filtering source" {
    & $cli -q range set-values --session $sessionId --sheet-name Filters --range-address A1:B6 --values '[["Category","Amount"],["A",10],["B",20],["C",30],["A",40],["B",50]]' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Apply native two-condition ordinary filter" {
    & $cli -q rangeedit apply-filter --session $sessionId --sheet-name Filters --range-address A1:B6 --column-index 2 --filter-options '{"filterOperator":"And","criteria1":">=20","criteria2":"<=40"}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read both native ordinary criteria" {
    & $cli -q rangeedit get-filters --session $sessionId --sheet-name Filters --range-address A1:B6 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.columnFilters.Count -eq 2 -and
    $r.columnFilters[1].filterOperator -eq 'And' -and
    $r.columnFilters[1].criteria1.value -eq '>=20' -and
    $r.columnFilters[1].criteria2.value -eq '<=40'
}
Test-Step "Verify native matching row count" {
    & $cli -q range get-special-cells --session $sessionId --sheet-name Filters --range-address A2:A6 --cell-kind Visible | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.cellCount -eq 3
}
Test-Step "Clear only the ordinary filter" {
    & $cli -q rangeedit clear-filters --session $sessionId --sheet-name Filters --range-address A1:B6 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Set native report layout without changing orientation" {
    & $cli -q worksheetstyle set-page-setup --session $sessionId --sheet-name Filters --page-setup-options '{"printArea":"A1:B6","leftMargin":36,"centerHeader":"Report","zoomPercent":100}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native report layout" {
    & $cli -q worksheetstyle get-page-setup --session $sessionId --sheet-name Filters | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.printArea -eq '$A$1:$B$6' -and $r.leftMargin -eq 36 -and $r.centerHeader -eq 'Report'
}
Test-Step "Replace manual report page breaks" {
    & $cli -q worksheetstyle set-page-breaks --session $sessionId --sheet-name Filters --page-break-options '{"rows":[4],"columns":[]}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Read native manual report page break" {
    & $cli -q worksheetstyle get-page-breaks --session $sessionId --sheet-name Filters | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and @($r.horizontal | Where-Object { $_.isManual -and $_.position -eq 4 }).Count -eq 1
}
Test-Step "Create native report slicer" {
    & $cli -q slicer create-slicer --session $sessionId --pivot-table-name CalculationPivot --field-name Region --slicer-name ReportRegions --destination-sheet Data --position AB1 | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Update only report slicer layout" {
    & $cli -q slicer update-slicer --session $sessionId --slicer-name ReportRegions --slicer-options '{"width":240,"height":180,"columnCount":2,"caption":"Regions","displayHeader":false}' | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.slicer.width -eq 240 -and $r.slicer.columnCount -eq 2
}
Test-Step "Read complete native report control" {
    & $cli -q slicer get-slicer --session $sessionId --slicer-name ReportRegions | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true -and $r.slicer.caption -eq 'Regions' -and $r.slicer.availableItems.Count -eq 2 -and $r.slicer.connectedPivotTables[0] -eq 'CalculationPivot'
}
Test-Step "Delete report slicer without changing Pivot calculation" {
    & $cli -q slicer delete-slicer --session $sessionId --slicer-name ReportRegions | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Create disposable worksheet" {
    & $cli -q sheet create --session $sessionId --sheet-name Disposable | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}
Test-Step "Delete disposable worksheet" {
    & $cli -q sheet delete --session $sessionId --sheet-name Disposable | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 7. Close session (with save)
Test-Step "Close session (with save)" {
    & $cli -q session close --session $sessionId --save | ConvertFrom-Json
} -Verify {
    param($r)
    $r.success -eq $true
}

# 8. Reopen saved file (session open - exercises Workbooks.Open path distinct from Add+SaveAs)
#    This step would catch deployment issues like missing office.dll (issue #487) because
#    ExcelBatch.ctor runs AutomationSecurity setup before opening any workbook.
$reopenSession = Test-Step "Reopen saved file (session open)" {
    & $cli -q session open $testFile | ConvertFrom-Json
} -Verify {
    param($r)
    $r.sessionId -and $r.success -eq $true
}

# 9. List worksheets in reopened session (proves the file loaded correctly)
if ($reopenSession -and $reopenSession.sessionId) {
    $reopenSessionId = $reopenSession.sessionId
    Test-Step "Verify value survived save and reopen" {
        & $cli -q range get-values --session $reopenSessionId --sheet-name Data --range-address A1 | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -and $r.values.Count -eq 1 -and $r.values[0][0] -eq 424242
    }

    Test-Step "Verify Pivot calculation survived save and reopen" {
        & $cli -q pivottablefield list-fields --session $reopenSessionId --pivot-table-name CalculationPivot | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true -and $r.valueFields.Count -eq 1 -and
        $r.valueFields[0].fieldName -eq 'Total Sales' -and $r.valueFields[0].calculation -eq 'PercentOfTotal' -and
        $r.valueFields[0].function -eq 'Sum'
    }

    Test-Step "Verify selected conditional rule survived save and reopen" {
        & $cli -q conditionalformat list-worksheet-rules --session $reopenSessionId --sheet-name Data | ConvertFrom-Json
    } -Verify {
        param($r)
        $original = @($r.rules | Where-Object appliesTo -EQ '$B$1:$B$8')
        $copied = @($r.rules | Where-Object appliesTo -EQ '$I$1:$I$2')
        $verified = $r.success -eq $true -and $r.rules.Count -eq 2 -and
        $original.Count -eq 1 -and $copied.Count -eq 1 -and
        $original[0].top10.rank -eq 5 -and $original[0].stopIfTrue -eq $false -and
        $copied[0].top10.rank -eq 5 -and $copied[0].stopIfTrue -eq $false
        if (-not $verified) {
            throw "Saved conditional rules differ: $($r | ConvertTo-Json -Depth 15 -Compress)"
        }
        $verified
    }

    # 10. Close reopened session
    Test-Step "Close reopened session" {
        & $cli -q session close --session $reopenSessionId | ConvertFrom-Json
    } -Verify {
        param($r)
        $r.success -eq $true
    }
}

# 11. Verify file exists
Test-Step "Verify file exists" {
    if (Test-Path $testFile) {
        $size = (Get-Item $testFile).Length
        "File size: $size bytes"
    } else {
        throw "File not found"
    }
} -Verify {
    param($r)
    $r -match "bytes"
}
} while ($false)

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "TEST SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $passed" -ForegroundColor Green
Write-Host "Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "Test file: $testFile" -ForegroundColor Gray

if ($KeepFile) {
    Write-Host "(Test file kept for inspection)" -ForegroundColor Yellow
}

if ($failed -gt 0) {
    Write-Host "`nSome tests FAILED!" -ForegroundColor Red
    $workflowExitCode = 1
} else {
    Write-Host "`nAll tests PASSED!" -ForegroundColor Green
    $workflowExitCode = 0
}
}
finally {
    try { $cleanupExitCode = Reset-CliWorkflowEnvironment }
    finally {
        if (-not $KeepFile -and (Test-Path -LiteralPath $testFile)) {
            Remove-Item -LiteralPath $testFile -Force
        }
        if (-not $KeepFile -and (Test-Path -LiteralPath $chartImagePath)) {
            Remove-Item -LiteralPath $chartImagePath -Force
        }
    }
}

if ($cleanupExitCode -ne 0) {
    Write-Error "Owned CLI cleanup failed with exit code $cleanupExitCode."
    exit $cleanupExitCode
}

exit $workflowExitCode
