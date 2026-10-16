#Requires -Version 7.4

[CmdletBinding()]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$root = $PSScriptRoot
$dotnet = (Get-Command dotnet -CommandType Application | Select-Object -First 1).Source
$dll = $env:EXCELMCP_BUILD_DLL
$executionDirectory = $null
$errorsFile = $null
$failure = $null
try {
    if (-not $dll) {
        $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($root)))
        $cache = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpBuild\$key"
        $executionDirectory = Join-Path $cache "runs\$([guid]::NewGuid().ToString('N'))"
        $dll = Join-Path $executionDirectory 'ExcelMcp.Build.dll'
        $mutex = [Threading.Mutex]::new($false, "ExcelMcp.Build.$key")
        $locked = $false
        try {
            try { $locked = $mutex.WaitOne([TimeSpan]::FromMinutes(2)) }
            catch [Threading.AbandonedMutexException] {
                $locked = $true
                [Console]::Error.WriteLine('Previous build-tool bootstrap stopped; rebuilding through the SDK.')
            }
            if (-not $locked) { throw 'Timed out waiting for build-tool bootstrap.' }
            Push-Location $root
            try {
                $output = & $dotnet build (Join-Path $root 'tools\ExcelMcp.Build\ExcelMcp.Build.csproj') `
                    -c Release --verbosity quiet --artifacts-path $cache -o $executionDirectory 2>&1
                foreach ($line in $output) { [Console]::Error.WriteLine($line) }
                if ($LASTEXITCODE -ne 0) { throw "Build-tool bootstrap failed with exit code $LASTEXITCODE." }
            } finally { Pop-Location }
        }
        finally {
            if ($locked) { $mutex.ReleaseMutex() }
            $mutex.Dispose()
        }
    }
    if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) { throw "Build-tool assembly does not exist: $dll" }
    $errorsFile = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpBuildErrors-$([Guid]::NewGuid().ToString('N')).log"
    & $dotnet $dll @Arguments 2> $errorsFile
    $code = $LASTEXITCODE
    $diagnostics = [IO.File]::ReadAllText($errorsFile)
    if ($diagnostics) { [Console]::Error.Write($diagnostics) }
    if ($code -ne 0) { throw "Build command failed with exit code $code.`n$diagnostics" }
}
catch { $failure = $_.Exception; throw }
finally {
    $cleanupErrors = [Collections.Generic.List[Exception]]::new()
    foreach ($path in @($errorsFile, $executionDirectory)) {
        if (-not $path) { continue }
        try {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
            }
        }
        catch { $cleanupErrors.Add($_.Exception) }
    }
    if ($cleanupErrors.Count) {
        if ($failure) { $cleanupErrors.Insert(0, $failure) }
        throw [AggregateException]::new('Build-tool cleanup failed.', $cleanupErrors)
    }
}
$global:LASTEXITCODE = 0
