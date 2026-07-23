param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$PluginDir = Split-Path -Parent $PSScriptRoot
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
    $existingContent = Get-Content $UserMcpConfig -Raw | ConvertFrom-Json
    $config = [pscustomobject]@{
        mcpServers = if ($existingContent.mcpServers) {
            $existingContent.mcpServers
        } else {
            [pscustomobject]@{}
        }
    }
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
$config | ConvertTo-Json -Depth 10 | Set-Content $UserMcpConfig -Encoding UTF8

Write-Host ""
Write-Host "✅ ExcelMcp MCP server installed globally!" -ForegroundColor Green
Write-Host "   Config:   $UserMcpConfig" -ForegroundColor Gray
Write-Host "   Wrapper:  $WrapperPath" -ForegroundColor Gray
Write-Host ""
Write-Host "The first real MCP invocation will auto-download the newest Windows runtime." -ForegroundColor Cyan
Write-Host "Verify installation:" -ForegroundColor Cyan
Write-Host "   copilot mcp list" -ForegroundColor Gray
