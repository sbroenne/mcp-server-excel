<#
.SYNOPSIS
    Tests portable npm lockfile validation without installing dependencies.
#>

$ErrorActionPreference = "Stop"
$guard = Join-Path $PSScriptRoot "check-npm-lockfiles.ps1"
$testRoot = Join-Path ([IO.Path]::GetTempPath()) "excelmcp-npm-tests-$([Guid]::NewGuid().ToString('N'))"
$testsRun = 0

function Invoke-TestProcess {
    param([string]$Command, [string[]]$Arguments)

    $info = [Diagnostics.ProcessStartInfo]::new($Command)
    $info.WorkingDirectory = $testRoot
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($info)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(30000)) {
            $process.Kill($true)
            throw "Test process exceeded 30 seconds."
        }
        return @{
            ExitCode = $process.ExitCode
            Output = $stdout.GetAwaiter().GetResult() + $stderr.GetAwaiter().GetResult()
        }
    }
    finally { $process.Dispose() }
}

function Invoke-TestGit {
    param([string[]]$GitArguments)
    $result = Invoke-TestProcess "git" $GitArguments
    if ($result.ExitCode -ne 0) { throw "Test Git command failed." }
}

function Set-Fixture {
    param([string]$Path, [string]$Content)
    $fullPath = Join-Path $testRoot $Path
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath)) | Out-Null
    [IO.File]::WriteAllText($fullPath, $Content)
}

function Assert-Guard {
    param([bool]$Success, [string]$ExpectedText, [switch]$Staged)
    $arguments = @("-NoProfile", "-File", $guard, "-RepositoryRoot", $testRoot)
    if ($Staged) { $arguments += "-Staged" }
    $result = Invoke-TestProcess "pwsh" $arguments
    if (($result.ExitCode -eq 0) -ne $Success -or $result.Output -notlike "*$ExpectedText*") {
        throw "Guard assertion failed: expected success=$Success and diagnostic '$ExpectedText'."
    }
    if ($result.Output -match "private-mirror|secret-token|secret-password") {
        throw "Guard exposed a download URL or credential."
    }
    $script:testsRun++
}

$portable = '{"lockfileVersion":3,"packages":{"":{"name":"fixture"},"node_modules/example":{"version":"1.2.3","integrity":"sha512-example","funding":{"url":"https://example.org/support"}},"node_modules/local":{"resolved":"file:../local","link":true}}}'
$unsafe = '{"lockfileVersion":3,"packages":{"node_modules/example":{"version":"1.2.3","resolved":"https://user:secret-password@private-mirror.invalid/example.tgz?secret-token","integrity":"sha512-example"}}}'

[IO.Directory]::CreateDirectory($testRoot) | Out-Null
try {
    Invoke-TestGit @("init", "--quiet")
    Set-Fixture "package-lock.json" $portable
    Invoke-TestGit @("add", "--", "package-lock.json")
    Assert-Guard $true "1 tracked npm lockfile"

    Set-Fixture "package-lock.json" $unsafe
    Assert-Guard $false "fixed download URL"
    Assert-Guard $true "1 tracked npm lockfile" -Staged
    Invoke-TestGit @("add", "--", "package-lock.json")
    Set-Fixture "package-lock.json" $portable
    Assert-Guard $false "fixed download URL" -Staged
    Invoke-TestGit @("add", "--", "package-lock.json")

    Set-Fixture "new nested/project/package-lock.json" $unsafe
    Assert-Guard $true "1 tracked npm lockfile"
    Invoke-TestGit @("add", "--", "new nested/project/package-lock.json")
    Assert-Guard $false "new nested/project/package-lock.json"
    Set-Fixture "new nested/project/package-lock.json" $portable
    Assert-Guard $true "2 tracked npm lockfile"

    Set-Fixture "node_modules/vendor/package-lock.json" $unsafe
    Invoke-TestGit @("add", "--", "node_modules/vendor/package-lock.json")
    Assert-Guard $true "2 tracked npm lockfile"

    foreach ($url in @("https://registry.npmjs.org/example.tgz", "http://private-mirror.invalid/example.tgz", "//private-mirror.invalid/example.tgz", "git+https://private-mirror.invalid/repo.git", "ftp://private-mirror.invalid/example.tgz")) {
        Set-Fixture "package-lock.json" (@{ lockfileVersion = 1; dependencies = @{ example = @{ version = "1.2.3"; resolved = $url } } } | ConvertTo-Json -Depth 10)
        Assert-Guard $false "fixed download URL"
    }

    Set-Fixture "package-lock.json" '{"lockfileVersion":1,"dependencies":{"example":{"version":"https://private-mirror.invalid/example.tgz"}}}'
    Assert-Guard $false "fixed download URL"
    Set-Fixture "package-lock.json" '{"lockfileVersion":2,"packages":{},"dependencies":{"example":{"version":"1.2.3","dependencies":{"nested":{"resolved":"https://private-mirror.invalid/nested.tgz"}}}}}'
    Assert-Guard $false "fixed download URL"

    Set-Fixture "package-lock.json" '{"private-mirror secret-password":'
    Assert-Guard $false "invalid JSON"
    Set-Fixture "package-lock.json" "null"
    Assert-Guard $false "JSON object"
    Set-Fixture "package-lock.json" $portable
    Set-Fixture "nested/npm-shrinkwrap.json" $unsafe
    Invoke-TestGit @("add", "--", "nested/npm-shrinkwrap.json")
    Assert-Guard $false "nested/npm-shrinkwrap.json"
    Set-Fixture "nested/npm-shrinkwrap.json" $portable
    Assert-Guard $true "3 tracked npm lockfile"

    Write-Host "Passed $testsRun npm lockfile regression checks."
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
