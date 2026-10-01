[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
$marketplacePath = Join-Path $repoRoot ".github\plugin\marketplace.json"
$marketplace = Get-Content $marketplacePath -Raw | ConvertFrom-Json

foreach ($plugin in $marketplace.plugins) {
    $pluginRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $plugin.source))
    $manifestPath = Join-Path $pluginRoot "plugin.json"
    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $versionPath = Join-Path $pluginRoot "version.txt"
    $skillVersionPath = Join-Path $pluginRoot "skills\$($plugin.name)\VERSION"

    $versions = @(
        @(
            $plugin.version
            $manifest.version
            (Get-Content $versionPath -Raw).Trim()
            (Get-Content $skillVersionPath -Raw).Trim()
        ) | Select-Object -Unique
    )

    if ($versions.Count -ne 1) {
        throw "$($plugin.name) has inconsistent versions: $($versions -join ', ')"
    }

    if ($manifest.name -ne $plugin.name) {
        throw "$($plugin.name) does not match the name in plugin.json."
    }

    if ($manifest.'$schema' -ne "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json") {
        throw "$($plugin.name) does not target the Agent Plugins 1.0.0 manifest schema."
    }

    $skillPath = Join-Path $pluginRoot "skills\$($plugin.name)\SKILL.md"
    if (-not (Test-Path $skillPath -PathType Leaf)) {
        throw "$($plugin.name) is missing its skill at $skillPath."
    }

    foreach ($item in Get-ChildItem $pluginRoot -Recurse -Force) {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "$($plugin.name) contains a reparse point: $($item.FullName)"
        }
        if ($item.Name -eq "install-global.ps1") {
            throw "$($plugin.name) contains a retired global installation helper: $($item.FullName)"
        }
    }
}

$mcpPath = Join-Path $repoRoot "plugins\excel-mcp\mcp.json"
$mcpConfig = Get-Content $mcpPath -Raw | ConvertFrom-Json
if ($mcpConfig.'$schema' -ne "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json") {
    throw "excel-mcp does not target the Agent Plugins 1.0.0 MCP schema."
}

$server = $mcpConfig.mcpServers.'excel-mcp'
if ($server.type -ne "stdio" -or $server.command -ne "npx") {
    throw "excel-mcp must use a stdio server with a single executable command token."
}

if (@($server.args).Count -ne 2 -or $server.args[0] -ne "-y" -or
    $server.args[1] -ne "@sbroenne/mcp-server-excel@latest") {
    throw "excel-mcp must launch the latest published npm package through npx."
}

foreach ($file in Get-ChildItem (Join-Path $repoRoot "plugins") -Recurse -File -Filter "*.md") {
    $content = Get-Content $file.FullName -Raw
    $retiredDocumentation = @(
        @{ Pattern = [regex]::Escape("install-global.ps1"); Description = "retired global installation helper" }
        @{ Pattern = [regex]::Escape("ExcelMcp-CLI-latest-windows.zip"); Description = "nonexistent unversioned CLI release asset" }
        @{ Pattern = [regex]::Escape("excel-mcp-server.exe"); Description = "retired MCP executable name" }
        @{ Pattern = [regex]::Escape("excel-mcp-bundle.mcpb"); Description = "retired MCPB asset name" }
        @{ Pattern = [regex]::Escape("file(action: 'open', filePath"); Description = "retired MCP file path parameter" }
        @{ Pattern = [regex]::Escape("file(action: 'close', sessionId"); Description = "retired MCP session parameter" }
        @{ Pattern = '(?<![A-Za-z0-9-])--range-address(?![A-Za-z0-9-])'; Description = "retired CLI range flag" }
        @{ Pattern = '(?<![A-Za-z0-9-])--sheet-name(?![A-Za-z0-9-])'; Description = "retired CLI worksheet flag" }
        @{ Pattern = '(?<![A-Za-z0-9-])--source-table-name(?![A-Za-z0-9-])'; Description = "retired CLI PivotTable source flag" }
    )

    foreach ($entry in $retiredDocumentation) {
        if ([regex]::IsMatch($content, $entry.Pattern)) {
            throw "Outdated documentation in $($file.FullName): $($entry.Description)"
        }
    }

    $matches = [regex]::Matches(
        $content,
        '\[[^\]]+\]\((?!https?://|#|mailto:)([^)#]+)(?:#[^)]+)?\)')

    foreach ($match in $matches) {
        $relativePath = $match.Groups[1].Value -replace '/', '\'
        $targetPath = [IO.Path]::GetFullPath((Join-Path $file.DirectoryName $relativePath))

        if (-not (Test-Path $targetPath)) {
            throw "Broken local link in $($file.FullName): $relativePath"
        }
    }
}

foreach ($script in Get-ChildItem (Join-Path $repoRoot "plugins") -Recurse -File -Filter "*.ps1") {
    $tokens = $null
    $errors = $null
    [Management.Automation.Language.Parser]::ParseFile(
        $script.FullName,
        [ref]$tokens,
        [ref]$errors) | Out-Null

    if (@($errors).Count -gt 0) {
        throw "PowerShell syntax errors in $($script.FullName): $($errors -join '; ')"
    }
}

$tempProfile = Join-Path ([IO.Path]::GetTempPath()) ("excel-plugin-test-" + [Guid]::NewGuid().ToString("N"))
$originalUserProfile = $env:USERPROFILE
$originalHome = $env:HOME
$originalPath = $env:PATH

try {
    $echoScript = Join-Path $tempProfile "echo-argument.js"

    New-Item -ItemType Directory -Path $tempProfile -Force | Out-Null
    [IO.File]::WriteAllText(
        $echoScript,
        'WScript.StdOut.Write("pipeline-ok");',
        [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText(
        (Join-Path $tempProfile "npx.cmd"),
        "@echo off`r`n",
        [Text.Encoding]::ASCII)
    $npmBin = Join-Path $tempProfile "node_modules\npm\bin"
    New-Item -ItemType Directory -Path $npmBin -Force | Out-Null
    [IO.File]::WriteAllText(
        (Join-Path $npmBin "npx-cli.js"),
        @'
const { spawnSync } = require("node:child_process");
if (process.argv[2] !== "-y" || process.argv[3] !== "@sbroenne/excelcli@latest") {
    throw new Error(`Unexpected npx arguments: ${process.argv.slice(2).join(" ")}`);
}
const child = spawnSync(
    `${process.env.SystemRoot}\\System32\\cscript.exe`,
    process.argv.slice(4),
    { stdio: "inherit" });
if (child.error) {
    throw child.error;
}
process.exit(child.status ?? 1);
'@,
        [Text.UTF8Encoding]::new($false))

    $env:USERPROFILE = $tempProfile
    $env:HOME = $tempProfile
    $env:PATH = "$tempProfile;$originalPath"

    $wrapper = Join-Path $repoRoot "plugins\excel-cli\bin\start-cli.ps1"
    $captured = (& $wrapper //nologo $echoScript | Out-String).Trim()

    if ($captured -ne "pipeline-ok") {
        throw "CLI wrapper pipeline capture failed. Expected 'pipeline-ok', got '$captured'."
    }

} finally {
    $env:USERPROFILE = $originalUserProfile
    $env:HOME = $originalHome
    $env:PATH = $originalPath

    if (Test-Path $tempProfile) {
        Remove-Item $tempProfile -Recurse -Force
    }
}

Write-Output "All plugin validation checks passed."
