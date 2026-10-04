$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$workflow = Get-Content -LiteralPath (Join-Path $root '.github\workflows\copilot-setup-steps.yml') -Raw
$match = [regex]::Match($workflow, '(?m)- name: Install repository tooling\r?\n(?:\s+shell: [^\r\n]+\r?\n)?\s+run: (?<command>[^\r\n]+)')
if (-not $match.Success) { throw 'The repository npm setup step was not found.' }
$command = $match.Groups['command'].Value.Trim().Trim("'")
$fixture = Join-Path ([IO.Path]::GetTempPath()) "copilot-npm-$([Guid]::NewGuid().ToString('N'))"
$programFiles = Join-Path $fixture 'Program Files'
$npmPath = Join-Path $programFiles 'nodejs\npm.cmd'
$previousProgramFiles = $env:ProgramFiles
$previousEnvironment = $env:RUNNER_ENVIRONMENT
$previousNpm = Get-Item Function:\global:npm -ErrorAction SilentlyContinue
$global:CopilotSetupNpmHostedCalls = 0
try {
    New-Item -ItemType Directory -Path (Split-Path -Parent $npmPath) -Force | Out-Null
    @'
@echo off
if not "%1"=="ci" exit /b 25
echo protected-npm-ci
echo cwd-%CD%
exit /b 0
'@ | Set-Content -LiteralPath $npmPath -Encoding ASCII
    $env:ProgramFiles = $programFiles
    $env:RUNNER_ENVIRONMENT = 'self-hosted'
    function global:npm { throw 'Copied runtime npm shim cannot find npm-prefix.js.' }
    Push-Location $root
    try { $output = @(Invoke-Expression $command) }
    finally { Pop-Location }
    if ($output -notcontains 'protected-npm-ci') {
        throw 'Self-hosted setup must use the protected npm.cmd, not the injected runtime shim.'
    }

    $helper = Join-Path $root 'scripts\Invoke-CopilotSetupNpm.ps1'
    $npmSteps = @('Install repository tooling', 'Install VS Code extension dependencies', 'Install shared npm tooling')
    foreach ($name in $npmSteps) {
        $step = [regex]::Match($workflow, "(?m)- name: $([regex]::Escape($name))\r?\n(?:\s+working-directory: (?<directory>[^\r\n]+)\r?\n)?(?:\s+shell: (?<shell>[^\r\n]+)\r?\n)?\s+run: (?<command>[^\r\n]+)")
        if (-not $step.Success) { throw "The npm setup step is missing: $name" }
        if ($step.Groups['shell'].Value.Trim() -ne 'pwsh') {
            throw "The cloud agent ignores job-level shell defaults; the PowerShell npm helper needs an explicit pwsh step: $name"
        }
        $directory = if ($step.Groups['directory'].Success) {
            Join-Path $root $step.Groups['directory'].Value.Trim().Replace('/', '\')
        } else { $root }
        $relative = $step.Groups['command'].Value.Trim().Trim("'")
        if ([IO.Path]::GetFullPath((Join-Path $directory $relative)) -ine $helper) {
            throw "The npm helper does not resolve from the step's working directory: $name"
        }
        Push-Location $directory
        try { $output = @(Invoke-Expression $relative) }
        finally { Pop-Location }
        if ($output -notcontains 'protected-npm-ci' -or $output -notcontains "cwd-$directory") {
            throw "Protected npm must retain the step's working directory: $name"
        }
    }
    $env:RUNNER_ENVIRONMENT = 'github-hosted'
    function global:npm {
        if ($args.Count -ne 1 -or $args[0] -ne 'ci') { throw 'Unexpected hosted npm arguments.' }
        $global:CopilotSetupNpmHostedCalls++
        $global:LASTEXITCODE = 0
    }
    & $helper
    if ($global:CopilotSetupNpmHostedCalls -ne 1) { throw 'Hosted setup must retain its configured npm.' }

    $env:RUNNER_ENVIRONMENT = 'self-hosted'
    '@exit /b 23' | Set-Content -LiteralPath $npmPath -Encoding ASCII
    $failure = $null
    try { & $helper }
    catch { $failure = $_.Exception.Message }
    if ($failure -notmatch '23') { throw 'A failed npm installation must preserve its exit code.' }
    Remove-Item -LiteralPath $npmPath
    $failure = $null
    try { & $helper }
    catch { $failure = $_.Exception.Message }
    if ($failure -notmatch 'protected npm') { throw 'Missing protected npm must fail explicitly.' }

    if ([regex]::Matches($workflow, 'run: .*Invoke-CopilotSetupNpm\.ps1').Count -ne $npmSteps.Count -or
        $workflow -match 'run: npm ci') {
        throw 'All setup npm installations must use the same protected-runtime selection.'
    }
}
finally {
    $env:ProgramFiles = $previousProgramFiles
    $env:RUNNER_ENVIRONMENT = $previousEnvironment
    Remove-Item Function:\global:npm -ErrorAction SilentlyContinue
    if ($previousNpm) { Set-Item Function:\global:npm -Value $previousNpm.ScriptBlock }
    Remove-Variable CopilotSetupNpmHostedCalls -Scope Global -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Recurse -Force }
}
Write-Output 'Copilot setup uses protected npm on the desktop, preserves hosted npm, and reports failures.'
