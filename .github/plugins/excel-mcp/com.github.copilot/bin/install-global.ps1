param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$PluginDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$WrapperPath = Join-Path $PluginDir "bin\start-mcp.ps1"
$UserMcpConfig = Join-Path $env:USERPROFILE ".copilot\mcp-config.json"

Write-Host "ExcelMcp Global Install Helper" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $WrapperPath)) {
    Write-Error "❌ Plugin wrapper not found at $WrapperPath"
    exit 1
}

$CopilotDir = Join-Path $env:USERPROFILE ".copilot"
if (-not (Test-Path $CopilotDir)) {
    Write-Host "[Install] Creating ~/.copilot directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $CopilotDir -Force | Out-Null
}

$config = [pscustomobject]@{
    mcpServers = [pscustomobject]@{}
}

if (Test-Path $UserMcpConfig) {
    Write-Host "[Install] Loading existing user MCP config..." -ForegroundColor Yellow
    $config = Get-Content $UserMcpConfig -Raw | ConvertFrom-Json
    if ($config -isnot [pscustomobject]) { throw "User MCP config must be a JSON object." }
    if ($null -eq $config.mcpServers) {
        $config | Add-Member -MemberType NoteProperty -Name mcpServers -Value ([pscustomobject]@{}) -Force
    }
    if ($config.mcpServers -isnot [pscustomobject]) { throw "mcpServers must be a JSON object." }
}

if ($config.mcpServers.PSObject.Properties.Name -contains "excel-mcp" -and -not $Force) {
    Write-Host "✅ excel-mcp server is already configured in user MCP config" -ForegroundColor Green
    Write-Host ""
    Write-Host "Run again with -Force to rewrite the wrapper path." -ForegroundColor Yellow
    exit 0
}

Write-Host "[Install] Adding excel-mcp to user MCP config..." -ForegroundColor Yellow

$excelMcpConfig = @{
    command = "powershell"
    args = @(
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        $WrapperPath
    )
}

$config.mcpServers | Add-Member -MemberType NoteProperty -Name "excel-mcp" -Value $excelMcpConfig -Force
$temporaryConfig = "$UserMcpConfig.$([Guid]::NewGuid().ToString('N')).tmp"
try {
    $json = $config | ConvertTo-Json -Depth 100 -WarningAction Stop
    [IO.File]::WriteAllText($temporaryConfig, $json, [Text.UTF8Encoding]::new($false))
    if (Test-Path -LiteralPath $UserMcpConfig) {
        [IO.File]::Replace($temporaryConfig, $UserMcpConfig, [System.Management.Automation.Language.NullString]::Value)
    } else {
        [IO.File]::Move($temporaryConfig, $UserMcpConfig)
    }
} finally {
    if (Test-Path -LiteralPath $temporaryConfig) { Remove-Item -LiteralPath $temporaryConfig }
}

Write-Host ""
Write-Host "✅ ExcelMcp MCP server installed globally!" -ForegroundColor Green
Write-Host "   Config:   $UserMcpConfig" -ForegroundColor Gray
Write-Host "   Wrapper:  $WrapperPath" -ForegroundColor Gray
Write-Host ""
Write-Host "MCP invocations use the latest npm package through npx when Node.js is available." -ForegroundColor Cyan
Write-Host "The plugin release downloader remains available as a fallback." -ForegroundColor Cyan
Write-Host "Verify installation:" -ForegroundColor Cyan
Write-Host "   copilot mcp list" -ForegroundColor Gray
