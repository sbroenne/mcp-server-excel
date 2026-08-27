<#
.SYNOPSIS
    Validates Copilot interpretation and adds it to a sanitized analytics report.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$AnalyticsPath,

    [Parameter(Mandatory = $true)]
    [string]$InterpretationPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$analyticsText = [IO.File]::ReadAllText([IO.Path]::GetFullPath($AnalyticsPath))
$analytics = $analyticsText | ConvertFrom-Json
$interpretation = [IO.File]::ReadAllText([IO.Path]::GetFullPath($InterpretationPath)).Trim()

if ($interpretation.Length -lt 100 -or $interpretation.Length -gt 4000) {
    throw "Interpretation must contain between 100 and 4000 characters."
}

$requiredHeadings = @(
    "## What changed",
    "## Reliability and performance",
    "## Adoption",
    "## Recommendations"
)
foreach ($heading in $requiredHeadings) {
    if ($interpretation -notmatch "(?m)^$([regex]::Escape($heading))\s*$") {
        throw "Interpretation is missing required heading '$heading'."
    }
}

$forbiddenPatterns = [ordered]@{
    "link" = 'https?://|\[[^\]]+\]\([^)]+\)'
    "HTML" = '<[A-Za-z!/][^>]*>'
    "email" = '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'
    "Windows path" = '[A-Za-z]:\\'
    "UNC path" = '\\\\[^\\\s]+\\'
    "identifier field" = '(?i)\b(UserId|SessionId|FileSessionId|ClientIP|ClientCity|ClientCountryOrRegion|StackTrace)\b'
}
foreach ($entry in $forbiddenPatterns.GetEnumerator()) {
    if ($interpretation -match $entry.Value) {
        throw "Interpretation contains forbidden $($entry.Key) content."
    }
}

$allowedNumbers = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($match in [regex]::Matches($analyticsText, '(?<![\w.])-?\d+(?:\.\d+)?')) {
    [void]$allowedNumbers.Add($match.Value)
}
foreach ($match in [regex]::Matches($interpretation, '(?<![\w.])-?\d+(?:\.\d+)?')) {
    if (-not $allowedNumbers.Contains($match.Value)) {
        throw "Interpretation contains unsupported numeric claim '$($match.Value)'."
    }
}

$analytics | Add-Member -NotePropertyName interpretation -NotePropertyValue $interpretation
$analytics | Add-Member -NotePropertyName interpretationModel -NotePropertyValue "GitHub Copilot CLI"
$json = $analytics | ConvertTo-Json -Depth 10
[IO.File]::WriteAllText(
    [IO.Path]::GetFullPath($OutputPath),
    $json + "`n",
    [Text.UTF8Encoding]::new($false))
