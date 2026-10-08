$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$properties = [xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)
$target = @($properties.Project.Target | Where-Object Name -EQ 'StopExcelCliService')
if ($target.Count -ne 1) { throw 'Expected one direct CLI-service stop target.' }
if ($target[0].Exec.Command -notlike '*scripts\Stop-ExcelCliService.ps1*') {
    throw 'The build must invoke the direct CLI-service stop script.'
}
if ($properties.OuterXml -match 'ExcelMcpCleanupRoot|ExcelMcpSkipCleanup|Stop-ExcelMcpProcesses') {
    throw 'The build still contains cleanup bootstrap machinery.'
}
$condition = [Security.SecurityElement]::Escape($target[0].Condition)
$directory = Join-Path ([IO.Path]::GetTempPath()) "excel-cleanup-condition-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $directory | Out-Null
try {
    foreach ($case in @(
        @{ Name = 'local'; CI = ''; Runner = ''; Project = 'ExcelMcp.CLI'; Expected = $true },
        @{ Name = 'desktop-ci'; CI = 'true'; Runner = 'self-hosted'; Project = 'ExcelMcp.CLI'; Expected = $true },
        @{ Name = 'hosted-ci'; CI = 'true'; Runner = 'github-hosted'; Project = 'ExcelMcp.CLI'; Expected = $false },
        @{ Name = 'unspecified-ci'; CI = 'true'; Runner = ''; Project = 'ExcelMcp.CLI'; Expected = $false },
        @{ Name = 'other-project'; CI = 'true'; Runner = 'self-hosted'; Project = 'ExcelMcp.Service'; Expected = $false }
    )) {
        $marker = Join-Path $directory "$($case.Name).txt"
        $escapedMarker = [Security.SecurityElement]::Escape($marker)
        $project = Join-Path $directory "$($case.Project).proj"
        @"
<Project>
  <Target Name="Probe" Condition="$condition">
    <WriteLinesToFile File="$escapedMarker" Lines="owned-cleanup-permitted" Overwrite="true" />
  </Target>
</Project>
"@ | Set-Content -LiteralPath $project -Encoding UTF8
        $output = & dotnet msbuild $project -nologo -verbosity:quiet -target:Probe `
            "-property:CI=$($case.CI)" "-property:RUNNER_ENVIRONMENT=$($case.Runner)" 2>&1
        if ($LASTEXITCODE -ne 0) { throw "MSBuild condition evaluation failed for $($case.Name): $output" }
        if ((Test-Path -LiteralPath $marker) -ne $case.Expected) {
            throw "Owned pre-build cleanup permission was incorrect for $($case.Name)."
        }
    }
}
finally { Remove-Item -LiteralPath $directory -Recurse -Force }
Write-Output 'Actual MSBuild service-stop condition preserves local, self-hosted, hosted and CLI-only boundaries.'
