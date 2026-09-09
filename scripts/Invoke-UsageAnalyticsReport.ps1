<#
.SYNOPSIS
    Generates and validates a Copilot interpretation for the usage analytics report.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$AnalyticsPath,

    [Parameter(Mandatory = $true)]
    [string]$InterpretationPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [ValidateRange(1, 5)]
    [int]$MaxAttempts = 3,

    [scriptblock]$CopilotInvoker = {
        param([string[]]$Arguments)
        $output = & copilot @Arguments 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = @($output) }
    }
)

$ErrorActionPreference = "Stop"
$completeScript = Join-Path $PSScriptRoot "Complete-UsageAnalyticsReport.ps1"
$resolvedAnalyticsPath = [IO.Path]::GetFullPath($AnalyticsPath)
$resolvedInterpretationPath = [IO.Path]::GetFullPath($InterpretationPath)
$resolvedOutputPath = [IO.Path]::GetFullPath($OutputPath)
$analyticsDirectory = [IO.Path]::GetDirectoryName($resolvedAnalyticsPath)
$analyticsFileName = [IO.Path]::GetFileName($resolvedAnalyticsPath)
$interpretationFileName = [IO.Path]::GetFileName($resolvedInterpretationPath)

if (-not (Test-Path -LiteralPath $resolvedAnalyticsPath -PathType Leaf)) {
    throw "Sanitized analytics input '$resolvedAnalyticsPath' does not exist."
}
if ([IO.Path]::GetDirectoryName($resolvedInterpretationPath) -ne $analyticsDirectory) {
    throw "Analytics input and interpretation output must use the same directory."
}

$requirements = @"
Read '$analyticsFileName', which contains only pre-checked anonymous totals. Write an interesting, useful '$interpretationFileName' for a general public audience. The entire file, including headings and blank lines, must contain between 100 and 3,500 characters. Use exactly these H2 headings: What changed; How well it worked; How people use it; What we will improve. Do more than repeat the tables. Connect the 90-day overview, weekly history, release-upgrade trend, latest comparison, repeat use, popular areas, and common actions. Explain what the relationship between users and actions may suggest. Mention advanced uses when the numbers support them. Workbook opening and closing are intentionally excluded as setup work, so do not discuss them. For reliability, clearly distinguish expected negative diagnostic results from failures, distinguish the fixed failure classes, and keep unclassified failures visible. Do not reinterpret releases before categorizedReliabilityMinimumVersion because those rows lack outcome labels. Do not claim to fix or explain Excel performance from duration alone. State reliability limitations plainly when categorized data is not yet available. Draw only cautious conclusions, using words such as suggests or may rather than claiming a cause. Format English numbers consistently: use comma thousands separators for counts of 1,000 or more and periods as decimal separators. Use one to three short paragraphs under each heading, short sentences, and everyday language. Explain internal names in plain English and never include raw action names. Do not use technical analytics or software-error jargon, including p50, p95, p99, percentile, invocation, cohort, telemetry, sanitized, AggregateException, COMException, or tail latency. Every numeric claim must use an exact number already present in '$analyticsFileName'. Do not add links, HTML, identifiers, geography, messages, stack traces, or guesses presented as fact. Clearly separate observations from planned improvements. Write only the requested file.
"@

$lastValidationError = $null
for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    Remove-Item -LiteralPath $resolvedInterpretationPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $resolvedOutputPath -Force -ErrorAction SilentlyContinue

    $prompt = if ($attempt -eq 1) {
        $requirements
    }
    else {
        "The previous '$interpretationFileName' failed validation: $lastValidationError Replace it with a new draft that follows every requirement below.`n`n$requirements"
    }

    Push-Location $analyticsDirectory
    try {
        $result = & $CopilotInvoker -Arguments @(
            "-p", $prompt,
            "--allow-tool=write",
            "--no-ask-user"
        )
    }
    finally {
        Pop-Location
    }

    if ($null -eq $result -or
        $null -eq $result.PSObject.Properties["ExitCode"] -or
        $null -eq $result.PSObject.Properties["Output"]) {
        throw "The Copilot invoker returned an invalid result."
    }
    if ($result.ExitCode -ne 0) {
        throw "Copilot interpretation generation failed: $(@($result.Output) -join "`n")"
    }

    try {
        & $completeScript `
            -AnalyticsPath $resolvedAnalyticsPath `
            -InterpretationPath $resolvedInterpretationPath `
            -OutputPath $resolvedOutputPath
        Write-Host "Usage analytics interpretation passed validation on attempt $attempt."
        return
    }
    catch {
        $lastValidationError = $_.Exception.Message
        if ($attempt -lt $MaxAttempts) {
            Write-Warning "Copilot interpretation attempt $attempt failed validation; regenerating."
        }
    }
}

throw "Copilot interpretation failed validation after $MaxAttempts attempts. Last error: $lastValidationError"
