[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ServerJsonPath,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$mcpServerPackageIds = @(
    'Sbroenne.ExcelMcp.McpServer',
    '@sbroenne/mcp-server-excel'
)

if (-not (Test-Path -LiteralPath $ServerJsonPath -PathType Leaf)) {
    throw "MCP Registry metadata file was not found: $ServerJsonPath"
}

$server = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
if ($null -eq $server.PSObject.Properties['version']) {
    throw "MCP Registry metadata must contain a top-level 'version' property."
}

if ($null -eq $server.PSObject.Properties['packages']) {
    throw "MCP Registry metadata must contain a 'packages' array."
}

$server.version = $Version
foreach ($packageId in $mcpServerPackageIds) {
    $matchingPackages = @(
        $server.packages | Where-Object {
            $_.identifier -eq $packageId
        }
    )

    if ($matchingPackages.Count -ne 1) {
        throw "MCP Registry metadata must contain exactly one package with identifier '$packageId'."
    }

    $package = $matchingPackages[0]
    if ($null -eq $package.PSObject.Properties['version']) {
        throw "MCP Registry package '$packageId' must contain a 'version' property."
    }

    $package.version = $Version
}

$server | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $ServerJsonPath -NoNewline

$updatedServer = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
if ($updatedServer.version -ne $Version) {
    throw "MCP Registry metadata validation failed after stamping version '$Version'."
}

foreach ($packageId in $mcpServerPackageIds) {
    $updatedPackages = @(
        $updatedServer.packages | Where-Object {
            $_.identifier -eq $packageId
        }
    )

    if ($updatedPackages.Count -ne 1 -or $updatedPackages[0].version -ne $Version) {
        throw "MCP Registry metadata validation failed for package '$packageId' after stamping version '$Version'."
    }
}

Write-Output "Updated MCP Registry metadata to version $Version."
