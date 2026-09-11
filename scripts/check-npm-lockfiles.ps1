<#
.SYNOPSIS
    Rejects fixed download URLs in all tracked npm lockfiles.
.DESCRIPTION
    Checks package-lock.json and npm-shrinkwrap.json, including nested projects,
    but not node_modules. Diagnostics never include dependency values or URLs.
    Use -Staged in commit hooks to check index contents rather than working files.
#>
[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [switch]$Staged
)

$ErrorActionPreference = "Stop"

function Invoke-LockfileGit {
    param([string[]]$GitArguments)

    $info = [Diagnostics.ProcessStartInfo]::new("git")
    $info.WorkingDirectory = $RepositoryRoot
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in $GitArguments) { $info.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($info)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(15000)) {
            $process.Kill($true)
            throw "Git lockfile inspection exceeded 15 seconds."
        }
        $output = $stdout.GetAwaiter().GetResult()
        $null = $stderr.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) { throw "Git could not read tracked npm lockfiles." }
        return $output
    }
    finally { $process.Dispose() }
}

function Test-FixedDownloadUrl {
    param($Value)

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            $child = $Value[$key]
            if ($key -in @("resolved", "version") -and $child -is [string] -and
                $child -match '^(?:[a-z][a-z0-9+.-]*:)?//') {
                return $true
            }
            if (Test-FixedDownloadUrl $child) { return $true }
        }
    }
    elseif ($Value -is [array]) {
        foreach ($child in $Value) {
            if (Test-FixedDownloadUrl $child) { return $true }
        }
    }
    return $false
}

try {
    $tracked = (Invoke-LockfileGit @("ls-files", "--cached", "-z")) -split "`0"
    $lockfiles = @($tracked | Where-Object {
        $_ -match '(^|/)(package-lock\.json|npm-shrinkwrap\.json)$' -and
        $_ -notmatch '(^|/)node_modules/'
    } | Sort-Object -Unique)
    $failed = $false
    foreach ($path in $lockfiles) {
        if ($Staged) {
            $content = Invoke-LockfileGit @("show", ":$path")
        }
        else {
            $fullPath = Join-Path $RepositoryRoot $path
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                throw "A tracked npm lockfile is missing: $path"
            }
            $content = [IO.File]::ReadAllText($fullPath)
        }
        try { $lockfile = ConvertFrom-Json -InputObject $content -AsHashtable }
        catch { throw "Npm lockfile has invalid JSON: $path" }
        if ($lockfile -isnot [System.Collections.IDictionary]) {
            throw "Npm lockfile must contain a JSON object: $path"
        }
        if (Test-FixedDownloadUrl $lockfile) {
            Write-Host "Fixed download URL found in $path (URL omitted)." -ForegroundColor Red
            $failed = $true
        }
    }
    if ($failed) {
        Write-Host "Set omit-lockfile-registry-resolved=true in each project's .npmrc, then run npm install --package-lock-only --ignore-scripts there. Direct URL dependencies must be replaced with portable dependencies."
        exit 1
    }
    Write-Host "Checked $($lockfiles.Count) tracked npm lockfile(s): no fixed download URLs."
    exit 0
}
catch {
    Write-Host "Npm lockfile check failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
