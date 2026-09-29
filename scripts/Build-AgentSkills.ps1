<#
.SYNOPSIS
    Builds the Excel MCP Agent Skills package for distribution.

.DESCRIPTION
    Creates distributable artifacts for Agent Skills:
    - excel-skills-v{version}.zip: Combined skill package with both excel-mcp and excel-cli
    GenerateOnly renders both skills and their complete references into ignored
    output. Packaging consumes that prepared directory without regenerating it.
    Users install with: npx skills add sbroenne/mcp-server-excel-plugins

.PARAMETER OutputDir
    Output directory for artifacts. Default: artifacts/skills

.PARAMETER Version
    Package version. Required unless GenerateOnly is used.

.PARAMETER GenerateOnly
    Generate complete skill directories after the Release solution build, without packaging.

.EXAMPLE
    ./Build-AgentSkills.ps1 -Version 1.2.0

.EXAMPLE
    ./Build-AgentSkills.ps1 -OutputDir ./dist -Version 1.2.0

.EXAMPLE
    ./Build-AgentSkills.ps1 -GenerateOnly
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [string]$Version,
    [switch]$GenerateOnly,
    [string]$SkillsDirectory,
    [string]$ManifestPath,
    [string]$CliExecutable
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot
$SkillsDir = Join-Path $RepoRoot "skills"
$SharedDir = Join-Path $SkillsDir "shared"
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')

function ConvertTo-PlainHelpLines {
    param([AllowEmptyCollection()][object[]]$Lines)

    return @($Lines | ForEach-Object {
        ([string]$_) -replace "`e\[[0-?]*[ -/]*[@-~]", ''
    })
}

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
        throw "Complete CLI reference generation requires Windows."
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
    $mainHelp = ConvertTo-PlainHelpLines @(& $ExcelCliPath --help 2>&1)
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
        $help = ConvertTo-PlainHelpLines @(& $ExcelCliPath $commandName --help 2>&1)
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
                $subcommandHelp = ConvertTo-PlainHelpLines @(& $ExcelCliPath $commandName $subcommandName --help 2>&1)
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
    $content.Add("- ``--timeout`` ranges are action-specific: session open/create/test accepts 10-3600; Power Query refresh/refresh-all accepts 0-2147483 (0 keeps the default); other generated timeout actions accept 1-2147483.")
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
        if ($FilesToCopy.Count -eq 0) { throw "No shared reference files found: $SharedDir" }
        $CopiedCount = 0
        foreach ($sourceFile in $FilesToCopy) {
            $destination = Join-Path $RefsDir $sourceFile.Name
            if ($SkillName -eq "excel-cli") {
                $cliSyntaxNotice = "> **CLI syntax note:** This shared domain guide may use MCP-style ``tool(action: ...)`` examples as conceptual shorthand. Do not translate or paste those calls mechanically. Use the exact commands and kebab-case options in [cli-commands.md](./cli-commands.md) or live ``--help``; notably, MCP ``file`` open/close maps to CLI ``session`` open/close, and MCP ``worksheet`` maps to CLI ``sheet``."
                $sourceContent = (Get-Content -Path $sourceFile.FullName -Raw) -replace "`r`n?", "`n"
                $adaptedContent = "$cliSyntaxNotice`n`n$sourceContent"
                Set-Content -Path $destination -Value $adaptedContent -Encoding UTF8 -NoNewline
            } else {
                Copy-Item -Path $sourceFile.FullName -Destination $destination -Force
            }
            $CopiedCount++
        }
        Write-Host "  Copied $CopiedCount shared references to $SkillName/references/" -ForegroundColor Green
    } else {
        throw "Required shared references not found: $SharedDir"
    }
}

if ($GenerateOnly) {
    if (-not $ManifestPath) {
        $ManifestPath = Join-Path $RepoRoot 'src\ExcelMcp.Core\obj\GeneratedFiles\ExcelMcp.Generators\Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator\_SkillManifest.g.cs'
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw "Required manifest not found: $ManifestPath" }
    if (-not $OutputDir) { $OutputDir = 'artifacts\generated-skills' }
    if (-not $Version) { $Version = (Get-Content (Join-Path $RepoRoot 'package.json') -Raw | ConvertFrom-Json).version }
}
elseif (-not $OutputDir) {
    $OutputDir = 'artifacts\skills'
}
if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "Version is required. Pass -Version <version>."
}
$Version = $Version.Trim()
$OutputPath = [IO.Path]::GetFullPath($OutputDir, $RepoRoot)
Assert-PackageOutputPath -Path $OutputPath -RepoRoot $RepoRoot
if ($OutputPath -eq [IO.Path]::GetPathRoot($OutputPath) -or
    $OutputPath -eq $RepoRoot -or
    $OutputPath.StartsWith("$SkillsDir$([IO.Path]::DirectorySeparatorChar)", [StringComparison]::OrdinalIgnoreCase) -or
    $OutputPath -eq $SkillsDir) {
    throw "Skill output must not overlap source files: $OutputPath"
}
if (-not $SkillsDirectory) { $SkillsDirectory = Join-Path $RepoRoot 'artifacts\generated-skills' }
$StagingDir = Join-Path ([IO.Path]::GetTempPath()) "excel-skills-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
try {
    $SkillsStagingDir = Join-Path $StagingDir "skills"
    New-Item -ItemType Directory -Path $SkillsStagingDir -Force | Out-Null
    foreach ($component in @('cli', 'mcp')) {
        $name = "excel-$component"
        $destination = Join-Path $SkillsStagingDir $name
        if ($GenerateOnly) {
            $assets = Join-Path $SkillsDir "assets\$name"
            Copy-Item -LiteralPath $assets -Destination $destination -Recurse
            $projectName = if ($component -eq 'cli') { 'CLI' } else { 'McpServer' }
            $target = if ($component -eq 'cli') { 'GenerateCliSkill' } else { 'GenerateMcpSkill' }
            dotnet msbuild (Join-Path $RepoRoot "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj") `
                "-target:$target" -p:Configuration=Release "-p:SkillOutputRoot=$SkillsStagingDir" `
                "-p:SkillManifestPath=$ManifestPath" -nodeReuse:false -verbosity:minimal
            if ($LASTEXITCODE -ne 0) { throw "$name rendering failed with exit code $LASTEXITCODE." }
            Copy-SharedReferences -SkillPath $destination -SkillName $name
            if ($component -eq 'cli') { Generate-CliReference -SkillPath $destination -ExcelCliPath $CliExecutable }
        }
        else {
            Copy-Item -LiteralPath (Join-Path $SkillsDirectory $name) -Destination $destination -Recurse
        }
        foreach ($required in @('SKILL.md', 'README.md', 'references\range.md')) {
            if (-not (Test-Path -LiteralPath (Join-Path $destination $required))) { throw "$name is missing $required." }
        }
        Set-Content (Join-Path $destination 'VERSION') $Version -Encoding utf8 -NoNewline
    }
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    if ($GenerateOnly) {
        foreach ($name in @('excel-cli', 'excel-mcp')) {
            $destination = Join-Path $OutputPath $name
            Install-PackageOutput -Source (Join-Path $SkillsStagingDir $name) -Destination $destination
        }
        Write-Host "Generated complete skills at $OutputPath"
    }
    else {
        Copy-Item (Join-Path $SkillsDir 'README.md') $StagingDir
        $zip = Join-Path $StagingDir "excel-skills-v$Version.zip"
        Compress-Archive -LiteralPath $SkillsStagingDir,(Join-Path $StagingDir 'README.md') -DestinationPath $zip
        Install-PackageOutput -Source $zip -Destination (Join-Path $OutputPath (Split-Path $zip -Leaf))
        Write-Host "Created $OutputPath\excel-skills-v$Version.zip"
    }
}
finally { Remove-Item -LiteralPath $StagingDir -Recurse -Force }
$global:LASTEXITCODE = 0
