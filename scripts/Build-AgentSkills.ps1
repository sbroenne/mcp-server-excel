<#
.SYNOPSIS
    Builds the Excel MCP Agent Skills package for distribution.

.DESCRIPTION
    Creates distributable artifacts for Agent Skills:
    - excel-skills-v{version}.zip: Combined skill package with both excel-mcp and excel-cli
    - CLAUDE.md: Claude Code project instructions
    - .cursorrules: Cursor project rules

    MCP shared behavioral guidance from skills/shared/ is automatically copied
    to excel-mcp/references/ during packaging. The excel-cli skill uses its
    generated references/cli-commands.md file as the CLI-specific source of truth.

    Users install with: npx skills add sbroenne/mcp-server-excel

.PARAMETER OutputDir
    Output directory for artifacts. Default: artifacts/skills

.PARAMETER Version
    Package version. Required unless PopulateReferences is used.

.PARAMETER PopulateReferences
    Copy MCP shared references and regenerate CLI command reference files for local development (without packaging).

.EXAMPLE
    ./Build-AgentSkills.ps1 -Version 1.2.0

.EXAMPLE
    ./Build-AgentSkills.ps1 -OutputDir ./dist -Version 1.2.0

.EXAMPLE
    ./Build-AgentSkills.ps1 -PopulateReferences
#>
param(
    [string]$OutputDir = "artifacts/skills",
    [string]$Version = $null,
    [switch]$PopulateReferences
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $SkillsDir "shared"

# Generate a complete reference from the built CLI so aliases and branch commands cannot drift.
function Generate-CliReference {
    param(
        [string]$SkillPath,
        [string]$ExcelCliPath = $null
    )

    if (-not $ExcelCliPath) {
        $ExcelCliPath = Join-Path $RepoRoot "src/ExcelMcp.CLI/bin/Release/net10.0-windows/excelcli.exe"
    }
    if ($env:OS -ne "Windows_NT" -and [System.IO.Path]::GetExtension($ExcelCliPath) -eq ".exe") {
        Write-Warning "Skipping CLI reference generation because the Windows executable cannot run on this host"
        return
    }
    if (-not (Test-Path $ExcelCliPath)) {
        throw "excelcli not found at $ExcelCliPath. Build it first with: dotnet build src/ExcelMcp.CLI -c Release"
    }

    function Get-HelpSection {
        param([string[]]$Lines, [string]$Header)

        $start = [Array]::IndexOf($Lines, $Header)
        if ($start -lt 0) {
            return @()
        }

        $section = [System.Collections.Generic.List[string]]::new()
        for ($index = $start + 1; $index -lt $Lines.Count; $index++) {
            if ($Lines[$index] -match '^[A-Z][A-Z ]+:$') {
                break
            }
            if ($section.Count -eq 0 -and [string]::IsNullOrWhiteSpace($Lines[$index])) {
                continue
            }
            $section.Add($Lines[$index])
        }
        return @($section)
    }

    function Join-WrappedText {
        param([System.Collections.Generic.List[string]]$Lines)
        return ((($Lines | ForEach-Object { $_.Trim() }) -join " ") -replace '\s+', ' ').Trim()
    }

    function Get-HelpEntries {
        param(
            [string[]]$Lines,
            [string]$Header,
            [ValidateSet("Command", "Argument", "Option")]
            [string]$Kind
        )

        $pattern = switch ($Kind) {
            "Command" { '^\s{4}(?<spec>\S+(?:\s+<[^>]+>)?)\s{2,}(?<description>.*)$' }
            "Argument" { '^\s{4}(?<spec><[^>]+>)\s{2,}(?<description>.*)$' }
            "Option" { '^\s{4,}(?<spec>(?:-\w,\s+)?--[\w-]+(?:\s+<[^>]+>)?)\s{2,}(?<description>.*)$' }
        }

        $entries = [System.Collections.Generic.List[object]]::new()
        $current = $null
        foreach ($line in (Get-HelpSection -Lines $Lines -Header $Header)) {
            if ($line -match $pattern) {
                if ($null -ne $current) {
                    $current.Description = Join-WrappedText -Lines $current.DescriptionLines
                    $entries.Add($current)
                }
                $current = [PSCustomObject]@{
                    Spec = $Matches.spec.Trim()
                    Description = ""
                    DescriptionLines = [System.Collections.Generic.List[string]]::new()
                }
                if (-not [string]::IsNullOrWhiteSpace($Matches.description)) {
                    $current.DescriptionLines.Add($Matches.description)
                }
            }
            elseif ($null -ne $current -and -not [string]::IsNullOrWhiteSpace($line)) {
                $current.DescriptionLines.Add($line)
            }
        }
        if ($null -ne $current) {
            $current.Description = Join-WrappedText -Lines $current.DescriptionLines
            $entries.Add($current)
        }
        return @($entries)
    }

    function Add-ParameterTable {
        param(
            [System.Collections.Generic.List[string]]$Markdown,
            [string[]]$HelpLines,
            [string[]]$KnownTokens = @()
        )

        function Restore-KnownTokens {
            param([string]$Text, [string[]]$Tokens)

            foreach ($token in ($Tokens | Sort-Object Length -Descending)) {
                $pattern = (($token.ToCharArray() | ForEach-Object {
                    [regex]::Escape([string]$_)
                }) -join '\s*')
                $Text = [regex]::Replace($Text, $pattern, $token)
            }
            $Text = [regex]::Replace(
                $Text,
                "'(?<left>[A-Za-z0-9]*[a-z][A-Z][A-Za-z0-9]*)\s+(?<right>[a-z][A-Za-z0-9]*)'",
                "'`${left}`${right}'")
            return $Text
        }

        $parameters = [System.Collections.Generic.List[object]]::new()
        foreach ($argument in (Get-HelpEntries -Lines $HelpLines -Header "ARGUMENTS:" -Kind Argument)) {
            if ($argument.Spec -ne "<ACTION>") {
                $parameters.Add([PSCustomObject]@{
                    Name = $argument.Spec.ToLowerInvariant()
                    Description = Restore-KnownTokens -Text $argument.Description -Tokens $KnownTokens
                })
            }
        }
        foreach ($option in (Get-HelpEntries -Lines $HelpLines -Header "OPTIONS:" -Kind Option)) {
            $name = [regex]::Match($option.Spec, '--[\w-]+').Value
            if ($name -and $name -ne "--help") {
                $parameters.Add([PSCustomObject]@{
                    Name = $name
                    Description = Restore-KnownTokens -Text $option.Description -Tokens $KnownTokens
                })
            }
        }
        if ($parameters.Count -eq 0) {
            return
        }

        $Markdown.Add("| Parameter | Description |")
        $Markdown.Add("|-----------|-------------|")
        foreach ($parameter in $parameters) {
            $description = $parameter.Description.Replace("|", "\|")
            $Markdown.Add("| ``$($parameter.Name)`` | $description |")
        }
        $Markdown.Add("")
    }

    Write-Host "  Generating CLI command reference from excelcli..." -ForegroundColor Cyan
    $mainHelp = @(& $ExcelCliPath --help 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to run '$ExcelCliPath --help'."
    }

    $cliAssemblyPath = [System.IO.Path]::ChangeExtension($ExcelCliPath, ".dll")
    if (-not (Test-Path $cliAssemblyPath)) {
        throw "CLI assembly not found at $cliAssemblyPath."
    }
    $assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path $cliAssemblyPath))
    $typesByCommand = @{}
    foreach ($type in @($assembly.GetTypes() | Where-Object {
        $_.Namespace -eq "Sbroenne.ExcelMcp.CLI.Generated" -and
        $_.Name -like "*Command" -and
        $_.Name -ne "CliCommandRegistration"
    })) {
        $typesByCommand[($type.Name -replace 'Command$', '').ToLowerInvariant()] = $type
    }
    $typesByCommand["calculationmode"] = $typesByCommand["calculation"]
    $typesByCommand["datamodelrelationship"] = $typesByCommand["datamodelrel"]
    $typesByCommand["worksheetstyle"] = $typesByCommand["sheetstyle"]

    $content = [System.Collections.Generic.List[string]]::new()
    $content.Add("# CLI Command Reference")
    $content.Add("")
    $content.Add("> Auto-generated from the built ``excelcli`` runtime. Use these exact command and parameter names.")
    $content.Add("")

    $commands = Get-HelpEntries -Lines $mainHelp -Header "COMMANDS:" -Kind Command
    foreach ($command in ($commands | Sort-Object { $_.Spec.Split(' ')[0] })) {
        $commandName = $command.Spec.Split(' ')[0]
        $help = @(& $ExcelCliPath $commandName --help 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to run '$ExcelCliPath $commandName --help'."
        }

        $descriptionLines = [System.Collections.Generic.List[string]]::new()
        foreach ($line in (Get-HelpSection -Lines $help -Header "DESCRIPTION:")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $descriptionLines.Add($line)
            }
        }
        $description = if ($descriptionLines.Count -gt 0) {
            Join-WrappedText -Lines $descriptionLines
        } else {
            $command.Description
        }
        $actions = @()

        $content.Add("### $commandName")
        $content.Add("")
        $content.Add($description)
        $content.Add("")

        if ($typesByCommand.ContainsKey($commandName)) {
            $commandType = $typesByCommand[$commandName]
            $instance = [Activator]::CreateInstance($commandType)
            $property = $commandType.GetProperty(
                "ValidActions",
                [System.Reflection.BindingFlags]"Public,NonPublic,Instance")
            $actions = @($property.GetValue($instance))
            $formattedActions = ($actions | ForEach-Object { [string]::Concat('`', $_, '`') }) -join ", "
            $content.Add("**Actions:** $formattedActions")
            $content.Add("")
        }

        $subcommands = Get-HelpEntries -Lines $help -Header "COMMANDS:" -Kind Command
        if ($subcommands.Count -gt 0) {
            foreach ($subcommand in $subcommands) {
                $subcommandName = $subcommand.Spec.Split(' ')[0]
                $subcommandHelp = @(& $ExcelCliPath $commandName $subcommandName --help 2>&1)
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to run '$ExcelCliPath $commandName $subcommandName --help'."
                }
                $content.Add("#### $commandName $subcommandName")
                $content.Add("")
                $content.Add($subcommand.Description)
                $content.Add("")
                Add-ParameterTable -Markdown $content -HelpLines $subcommandHelp -KnownTokens @($subcommandName)
            }
        } else {
            Add-ParameterTable -Markdown $content -HelpLines $help -KnownTokens $actions
        }
    }

    $content.Add("## Common Pitfalls")
    $content.Add("")
    $content.Add("- ``--values-file`` requires an existing JSON or CSV file; use ``--values`` for inline JSON.")
    $content.Add("- ``--timeout`` ranges are action-specific: session open/create accepts 10-3600; Power Query refresh/refresh-all accepts 0-2147483 (0 keeps the default); other generated timeout actions accept 1-2147483.")
    $content.Add("- ``pythoninexcel get-result --max-wait-seconds`` must be at least 1 and shorter than the session operation timeout.")
    $content.Add("- ``--values`` and list parameters use JSON arrays; range values use a two-dimensional array.")
    $content.Add("- Power Query operations may take 30 seconds or longer; use a deliberate data-operation timeout or 0 for the default.")
    $content.Add("")

    $refsDir = Join-Path $SkillPath "references"
    New-Item -ItemType Directory -Path $refsDir -Force | Out-Null
    $outputFile = Join-Path $refsDir "cli-commands.md"
    $content -join "`n" | Set-Content -Path $outputFile -Encoding UTF8 -NoNewline
    Write-Host "  Generated: cli-commands.md" -ForegroundColor Green
}

# Function to copy shared references to a skill's references folder
function Copy-SharedReferences {
    param(
        [string]$SkillPath,
        [string]$SkillName
    )

    $RefsDir = Join-Path $SkillPath "references"

    # Create references directory if it doesn't exist
    if (-not (Test-Path $RefsDir)) {
        New-Item -ItemType Directory -Path $RefsDir -Force | Out-Null
    }

    if (Test-Path $SharedDir) {
        $FilesToCopy = @(Get-ChildItem -Path $SharedDir -File -Filter "*.md")
        $CopiedCount = 0
        foreach ($sourceFile in $FilesToCopy) {
            $destination = Join-Path $RefsDir $sourceFile.Name
            if ($SkillName -eq "excel-cli") {
                $cliSyntaxNotice = "> **CLI syntax note:** This shared domain guide may use MCP-style ``tool(action: ...)`` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live ``--help``; notably, MCP ``file`` open/close maps to CLI ``session`` open/close, and MCP ``worksheet`` maps to CLI ``sheet``."
                $adaptedContent = "$cliSyntaxNotice`r`n`r`n$(Get-Content -Path $sourceFile.FullName -Raw)"
                Set-Content -Path $destination -Value $adaptedContent -Encoding UTF8 -NoNewline
            } else {
                Copy-Item -Path $sourceFile.FullName -Destination $destination -Force
            }
            $CopiedCount++
        }
        Write-Host "  Copied $CopiedCount shared references to $SkillName/references/" -ForegroundColor Green
    } else {
        Write-Warning "Shared directory not found: $SharedDir"
    }
}

# Handle -PopulateReferences mode (for development)
if ($PopulateReferences) {
    Write-Host "Populating references from shared/ for local development..." -ForegroundColor Cyan

    # Copy to excel-mcp
    $McpPath = Join-Path $SkillsDir "excel-mcp"
    if (Test-Path $McpPath) {
        Copy-SharedReferences -SkillPath $McpPath -SkillName "excel-mcp"
    }

    # Copy to excel-cli
    $CliPath = Join-Path $SkillsDir "excel-cli"
    if (Test-Path $CliPath) {
        Copy-SharedReferences -SkillPath $CliPath -SkillName "excel-cli"
        # Generate CLI command reference from excelcli --help
        Generate-CliReference -SkillPath $CliPath
    }

    Write-Host ""
    Write-Host "Done! References populated for local development." -ForegroundColor Green
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "Version is required. Pass -Version <version>."
}
$Version = $Version.Trim()

Write-Host "Building Agent Skills package v$Version" -ForegroundColor Cyan
Write-Host "Source: $SkillsDir"
Write-Host "Output: $OutputDir"
Write-Host ""

# Create output directory
$OutputPath = Join-Path $RepoRoot $OutputDir
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Build combined skills package
Write-Host "Building combined skills package..." -ForegroundColor Yellow

# Create staging directory
$StagingDir = Join-Path ([System.IO.Path]::GetTempPath()) "excel-skills-$([guid]::NewGuid().ToString('N').Substring(0,8))"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

try {
    # Create skills/ subdirectory (the standard location for npx skills add)
    $SkillsStagingDir = Join-Path $StagingDir "skills"
    New-Item -ItemType Directory -Path $SkillsStagingDir -Force | Out-Null

    # Copy excel-mcp skill
    $McpSource = Join-Path $SkillsDir "excel-mcp"
    if (Test-Path $McpSource) {
        Copy-Item -Path $McpSource -Destination "$SkillsStagingDir/excel-mcp" -Recurse
        Copy-SharedReferences -SkillPath "$SkillsStagingDir/excel-mcp" -SkillName "excel-mcp"
        Set-Content -Path "$SkillsStagingDir/excel-mcp/VERSION" -Value $Version -Encoding UTF8 -NoNewline
    } else {
        Write-Warning "excel-mcp skill not found"
    }

    # Copy excel-cli skill
    $CliSource = Join-Path $SkillsDir "excel-cli"
    if (Test-Path $CliSource) {
        Copy-Item -Path $CliSource -Destination "$SkillsStagingDir/excel-cli" -Recurse
        Copy-SharedReferences -SkillPath "$SkillsStagingDir/excel-cli" -SkillName "excel-cli"
        # Generate CLI command reference from excelcli --help
        Generate-CliReference -SkillPath "$SkillsStagingDir/excel-cli"
        Set-Content -Path "$SkillsStagingDir/excel-cli/VERSION" -Value $Version -Encoding UTF8 -NoNewline
    } else {
        Write-Warning "excel-cli skill not found"
    }

    # Copy skills README to root of package
    $SkillsReadme = Join-Path $SkillsDir "README.md"
    if (Test-Path $SkillsReadme) {
        Copy-Item -Path $SkillsReadme -Destination $StagingDir
    }

    # Create ZIP archive
    $ZipName = "excel-skills-v$Version.zip"
    $ZipPath = Join-Path $OutputPath $ZipName

    if (Test-Path $ZipPath) {
        Remove-Item $ZipPath -Force
    }

    Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal
    Write-Host "  Created: $ZipName" -ForegroundColor Green

} finally {
    if (Test-Path $StagingDir) {
        Remove-Item $StagingDir -Recurse -Force
    }
}

# Copy CLAUDE.md and .cursorrules
Write-Host "Copying platform-specific files..." -ForegroundColor Yellow

$ClaudeSrc = Join-Path $SkillsDir "CLAUDE.md"
if (Test-Path $ClaudeSrc) {
    Copy-Item -Path $ClaudeSrc -Destination $OutputPath
    Write-Host "  Created: CLAUDE.md" -ForegroundColor Green
}

$CursorSrc = Join-Path $SkillsDir ".cursorrules"
if (Test-Path $CursorSrc) {
    Copy-Item -Path $CursorSrc -Destination $OutputPath
    Write-Host "  Created: .cursorrules" -ForegroundColor Green
}

# Generate manifest
$Manifest = @{
    name = "excel-skills"
    version = $Version
    description = "Excel MCP Server Agent Skills for AI coding assistants"
    platforms = @("github-copilot", "claude-code", "cursor", "windsurf", "gemini-cli", "goose", "codex", "opencode", "amp", "kilo", "roo", "trae")
    skills = @(
        @{
            name = "excel-mcp"
            path = "skills/excel-mcp"
            description = "MCP Server skill - for conversational AI (Claude Desktop, VS Code Chat)"
            target = "MCP Server"
        }
        @{
            name = "excel-cli"
            path = "skills/excel-cli"
            description = "CLI skill - for coding agents (Copilot, Cursor, Windsurf)"
            target = "CLI Tool"
        }
    )
    installation = @{
        npx = "npx skills add sbroenne/mcp-server-excel"
        selectSkill = "npx skills add sbroenne/mcp-server-excel --skill excel-cli"
        installBoth = "npx skills add sbroenne/mcp-server-excel --skill '*'"
    }
    files = @(
        @{ name = "CLAUDE.md"; type = "config"; description = "Claude Code project instructions" }
        @{ name = ".cursorrules"; type = "config"; description = "Cursor project rules" }
    )
    repository = "https://github.com/sbroenne/mcp-server-excel"
    documentation = "https://excelmcpserver.dev/"
    buildDate = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
}

$ManifestPath = Join-Path $OutputPath "manifest.json"
$Manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $ManifestPath -Encoding UTF8
Write-Host "  Created: manifest.json" -ForegroundColor Green

Write-Host ""
Write-Host "Build complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Output files in: $OutputPath" -ForegroundColor Cyan
Get-ChildItem $OutputPath | ForEach-Object {
    $Size = if ($_.Length -gt 1MB) { "{0:N2} MB" -f ($_.Length / 1MB) }
            elseif ($_.Length -gt 1KB) { "{0:N2} KB" -f ($_.Length / 1KB) }
            else { "{0} bytes" -f $_.Length }
    Write-Host "  $($_.Name) ($Size)"
}

Write-Host ""
Write-Host "Installation:" -ForegroundColor Cyan
Write-Host "  npx skills add sbroenne/mcp-server-excel" -ForegroundColor White
Write-Host "  (users will be prompted to select excel-cli, excel-mcp, or both)"
