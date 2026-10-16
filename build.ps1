#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
. (Join-Path $PSScriptRoot 'scripts\ValidationExecution.ps1')

if (-not $Arguments.Count -or $Arguments[0] -in @('help', '--help', '-h')) {
    'ExcelMcp: build.ps1 plan|build|test|validate [--full|--staged|--base-ref REF|--paths-json JSON|--area AREA|--plan FILE]'
    'Options: --head-ref REF --comparison MergeBase|Direct --output FILE --group Fast|Process|Tooling|Excel'
    '         --no-build --list-tests --list-projects --results-directory DIR --deadline-seconds SECONDS'
    $global:LASTEXITCODE = 0
    return
}
$verb = $Arguments[0]
if ($verb -notin @('plan', 'build', 'test', 'validate')) { throw "Unknown build command: $verb." }
$options = @{}
$flags = @('--full', '--staged', '--no-build', '--list-tests', '--list-projects')
$values = @('--base-ref', '--head-ref', '--comparison', '--paths-json', '--area', '--plan', '--output', '--group', '--results-directory', '--deadline-seconds')
for ($i = 1; $i -lt $Arguments.Count; $i++) {
    $key = $Arguments[$i]
    if ($options.ContainsKey($key)) { throw "Duplicate option: $key." }
    if ($key -in $flags) { $options[$key] = $true }
    elseif ($key -in $values -and $i + 1 -lt $Arguments.Count -and -not $Arguments[$i + 1].StartsWith('--')) {
        $options[$key] = $Arguments[++$i]
    } else { throw "Unknown option or missing value: $key." }
}
if ($options['--plan']) { $plan = Read-ValidationPlan $options['--plan'] }
else {
    $selection = @{}
    foreach ($entry in @(@('--full', 'Full'), @('--staged', 'Staged'), @('--base-ref', 'BaseRef'), @('--head-ref', 'HeadRef'), @('--comparison', 'Comparison'))) {
        if ($options.ContainsKey($entry[0])) { $selection[$entry[1]] = $options[$entry[0]] }
    }
    if ($options.ContainsKey('--paths-json')) { $selection.Paths = @($options['--paths-json'] | ConvertFrom-Json) }
    if ($options['--area']) {
        if ($options['--area'] -notmatch '^[A-Za-z][A-Za-z0-9]*$') { throw 'Area must be a command area name.' }
        $selection.Paths = @("src/ExcelMcp.Core/Commands/$($options['--area'])/Commands.cs")
    }
    if (-not $selection.Count) { $selection.BaseRef = 'origin/main' }
    $plan = Get-ValidationPlan @selection
}
if ($verb -eq 'plan') {
    $json = $plan | ConvertTo-Json -Depth 10
    if ($options['--output']) {
        $parent = Split-Path -Parent $options['--output']
        if ($parent) { New-Item -ItemType Directory $parent -Force | Out-Null }
        Set-Content -LiteralPath $options['--output'] -Value $json -Encoding utf8
    } else { $json }
    $global:LASTEXITCODE = 0
    return
}
$group = $options['--group']
if ($options['--list-projects']) { Get-SelectedBuildProjects $plan $group; return }
$deadline = if ($options['--deadline-seconds']) { [int]$options['--deadline-seconds'] } else { 0 }
if ($options.ContainsKey('--deadline-seconds') -and ($deadline -lt 1 -or $deadline -gt 28800)) { throw 'A hard deadline must be 1-28800 seconds.' }
if ($verb -eq 'build' -or -not $options['--no-build']) { Invoke-SelectedBuild $plan $group }
if ($verb -eq 'validate' -and $plan.DocumentationCounts -and (-not $group -or $group -eq $plan.DocumentationChecksGroup)) {
    & (Join-Path $PSScriptRoot 'scripts\Build-AgentSkills.ps1') -GenerateOnly
}
if ($verb -ne 'build') {
    $results = if ($options['--results-directory']) { $options['--results-directory'] } else { Join-Path $PSScriptRoot "TestResults\selected-$([Guid]::NewGuid().ToString('N'))" }
    foreach ($kind in $(if ($group -and $group -ne 'Excel') { @($group) } elseif ($group -eq 'Excel') { @() } else { $plan.CiTestGroups })) {
        $settings = @{ Plan = $plan; Group = $kind; ResultsDirectory = Join-Path $results $kind; ListTests = [bool]$options['--list-tests'] }
        if ($deadline) { $settings.DeadlineSeconds = $deadline }
        Invoke-SelectedFreeTests @settings
    }
    if ((-not $group -and $plan.Excel) -or $group -eq 'Excel') {
        $settings = @{ Plan = $plan; ResultsDirectory = $results; ListTests = [bool]$options['--list-tests'] }
        if ($deadline) { $settings.DeadlineSeconds = $deadline }
        Invoke-ExcelPlan @settings
    }
}
$global:LASTEXITCODE = 0
