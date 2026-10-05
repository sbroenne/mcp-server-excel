[CmdletBinding()]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$dll = $env:EXCELMCP_BUILD_DLL
if (-not $dll) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        $inputs = @(
            Get-ChildItem (Join-Path $root 'tools\ExcelMcp.Build') -File -Recurse |
                Where-Object { $_.Extension -in @('.cs', '.csproj') -and $_.FullName -notmatch '[\\/](bin|obj)[\\/]' }
            Get-Item (Join-Path $root 'Directory.Build.props')
            Get-Item (Join-Path $root 'Directory.Build.targets')
            Get-Item (Join-Path $root 'Directory.Packages.props')
            Get-Item (Join-Path $root 'global.json')
        ) | Sort-Object FullName
        $identity = $root + ($inputs | ForEach-Object {
            $_.FullName + ':' + [Convert]::ToBase64String($hash.ComputeHash([IO.File]::ReadAllBytes($_.FullName)))
        } | Out-String)
        $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($identity))).Replace('-', '')
    }
    finally { $hash.Dispose() }
    $cache = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpBuild\$key"
    $dll = Join-Path $cache 'bin\ExcelMcp.Build.dll'
    $stamp = Join-Path $cache 'bootstrap.complete'
    $mutex = [Threading.Mutex]::new($false, "ExcelMcp.Build.$key")
    $locked = $false
    try {
        try { $locked = $mutex.WaitOne([TimeSpan]::FromMinutes(2)) }
        catch [Threading.AbandonedMutexException] {
            $locked = $true
            [Console]::Error.WriteLine('Previous build-tool bootstrap stopped; checking its output.')
        }
        if (-not $locked) { throw 'Timed out waiting for build-tool bootstrap.' }
        if (-not (Test-Path -LiteralPath $stamp -PathType Leaf) -or -not (Test-Path -LiteralPath $dll -PathType Leaf)) {
            Push-Location $root
            try {
                $output = & (Get-Command dotnet -CommandType Application).Source build (Join-Path $root 'tools\ExcelMcp.Build\ExcelMcp.Build.csproj') `
                    -c Release --disable-build-servers --verbosity quiet `
                    "-p:BaseIntermediateOutputPath=$cache\obj\" "-p:MSBuildProjectExtensionsPath=$cache\obj\" `
                    "-p:BaseOutputPath=$cache\build\" -o (Join-Path $cache 'bin') 2>&1
                foreach ($line in $output) { [Console]::Error.WriteLine($line) }
                if ($LASTEXITCODE -ne 0) { throw "Build-tool bootstrap failed with exit code $LASTEXITCODE." }
                Set-Content -LiteralPath $stamp -Value $key -NoNewline
            } finally { Pop-Location }
        }
    }
    finally {
        if ($locked) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}
if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) { throw "Build-tool assembly does not exist: $dll" }
$errorsFile = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpBuildErrors-$([Guid]::NewGuid().ToString('N')).log"
try {
    & (Get-Command dotnet -CommandType Application).Source $dll @Arguments 2> $errorsFile
    $code = $LASTEXITCODE
    $diagnostics = [IO.File]::ReadAllText($errorsFile)
    if ($diagnostics) { [Console]::Error.Write($diagnostics) }
    if ($code -ne 0) { throw "Build command failed with exit code $code.`n$diagnostics" }
}
finally {
    if (Test-Path -LiteralPath $errorsFile) { Remove-Item -LiteralPath $errorsFile -Force }
}
$global:LASTEXITCODE = 0
