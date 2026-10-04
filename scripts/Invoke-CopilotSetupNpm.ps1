$ErrorActionPreference = 'Stop'
$npm = 'npm'
if ($env:RUNNER_ENVIRONMENT -eq 'self-hosted') {
    $npm = Join-Path $env:ProgramFiles 'nodejs\npm.cmd'
    if (-not (Test-Path -LiteralPath $npm -PathType Leaf)) {
        throw 'The qualified desktop is missing its protected npm command.'
    }
}
& $npm ci
if ($LASTEXITCODE -ne 0) {
    throw "Copilot npm dependency installation failed with exit code $LASTEXITCODE."
}
