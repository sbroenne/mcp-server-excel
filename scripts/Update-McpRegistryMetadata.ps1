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
$mcpServerPackageId = 'Sbroenne.ExcelMcp.McpServer'

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

$mcpPackages = @(
    $server.packages | Where-Object {
        $_.identifier -eq $mcpServerPackageId
    }
)

if ($mcpPackages.Count -ne 1) {
    throw "MCP Registry metadata must contain exactly one package with identifier '$mcpServerPackageId'."
}

$mcpPackage = $mcpPackages[0]
if ($null -eq $mcpPackage.PSObject.Properties['version']) {
    throw "MCP Registry package '$mcpServerPackageId' must contain a 'version' property."
}

$server.version = $Version
$mcpPackage.version = $Version
$content = ($server | ConvertTo-Json -Depth 20) -replace "`r?`n", "`n"
[System.IO.File]::WriteAllText(
    $ServerJsonPath,
    "$content`n",
    [System.Text.UTF8Encoding]::new($false))

$updatedServer = Get-Content -LiteralPath $ServerJsonPath -Raw | ConvertFrom-Json
$updatedMcpPackages = @(
    $updatedServer.packages | Where-Object {
        $_.identifier -eq $mcpServerPackageId
    }
)

if ($updatedServer.version -ne $Version -or
    $updatedMcpPackages.Count -ne 1 -or
    $updatedMcpPackages[0].version -ne $Version) {
    throw "MCP Registry metadata validation failed after stamping version '$Version'."
}

Write-Output "Updated MCP Registry metadata to version $Version."
