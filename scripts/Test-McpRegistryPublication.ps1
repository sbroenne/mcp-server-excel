[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ServerJsonPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$')]
    [string]$Version,

    [ValidateRange(1, 100)]
    [int]$Attempts = 3,

    [ValidateRange(0, 3600)]
    [int]$RetrySeconds = 600
)

$ErrorActionPreference = 'Stop'
$registryName = 'io.github.sbroenne/mcp-server-excel'
$repositoryUrl = 'https://github.com/sbroenne/mcp-server-excel'
$nugetPackage = 'Sbroenne.ExcelMcp.McpServer'
$npmLauncherPackage = '@sbroenne/mcp-server-excel'
$npmRuntimePackage = '@sbroenne/mcp-server-excel-win32-x64'

if (-not (Test-Path -LiteralPath $ServerJsonPath -PathType Leaf)) {
    throw "MCP Registry metadata file was not found: $ServerJsonPath"
}

$server = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
if ($server.name -ne $registryName -or
    $server.version -ne $Version -or
    $server.repository.url -ne $repositoryUrl -or
    $server.repository.source -ne 'github') {
    throw "MCP Registry source metadata does not match release version '$Version'."
}

foreach ($expectedPackage in @(
        @{ RegistryType = 'nuget'; Identifier = $nugetPackage },
        @{ RegistryType = 'npm'; Identifier = $npmLauncherPackage }
    )) {
    $matches = @($server.packages | Where-Object {
            $_.registryType -eq $expectedPackage.RegistryType -and
            $_.identifier -eq $expectedPackage.Identifier -and
            $_.version -eq $Version
        })
    if ($matches.Count -ne 1) {
        throw "MCP Registry source metadata must contain exactly one matching '$($expectedPackage.Identifier)' package."
    }
}

for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
        $readmeResponse = Invoke-WebRequest "https://api.nuget.org/v3-flatcontainer/sbroenne.excelmcp.mcpserver/$Version/readme"
        $readme = if ($readmeResponse.Content -is [byte[]]) {
            [Text.Encoding]::UTF8.GetString($readmeResponse.Content)
        } else {
            [string]$readmeResponse.Content
        }
        $nugetRegistration = Invoke-RestMethod "https://api.nuget.org/v3/registration5-semver1/sbroenne.excelmcp.mcpserver/$Version.json"
        if ($nugetRegistration.catalogEntry -isnot [string]) {
            throw 'Published NuGet registration metadata is not ready.'
        }
        $nuget = Invoke-RestMethod $nugetRegistration.catalogEntry
        $launcher = Invoke-RestMethod "https://registry.npmjs.org/@sbroenne%2fmcp-server-excel/$Version"
        $runtime = Invoke-RestMethod "https://registry.npmjs.org/@sbroenne%2fmcp-server-excel-win32-x64/$Version"

        if ($readme -notmatch 'mcp-name:\s+io\.github\.sbroenne/mcp-server-excel' -or
            $nuget.id -ne $nugetPackage -or
            $nuget.version -ne $Version -or
            $launcher.name -ne $npmLauncherPackage -or
            $launcher.version -ne $Version -or
            $launcher.mcpName -ne $registryName -or
            $runtime.name -ne $npmRuntimePackage -or
            $runtime.version -ne $Version) {
            throw 'Published package metadata is not ready.'
        }

        Write-Output "Validated MCP Registry source, NuGet, and npm metadata for version $Version."
        return
    } catch {
        if ($attempt -eq $Attempts) { throw }
        Write-Warning "Package propagation pending: $_"
        Start-Sleep -Seconds $RetrySeconds
    }
}
