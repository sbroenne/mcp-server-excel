function Invoke-TypedPackage {
    param(
        [Parameter(Mandatory)][hashtable]$Options,
        [Parameter(Mandatory)][string]$SourceRoot
    )
    $toolRoot = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { $SourceRoot }
    $bootstrap = Join-Path $toolRoot 'build.ps1'
    if (-not (Test-Path -LiteralPath $bootstrap -PathType Leaf)) {
        throw "The shared package tool bootstrap is missing: $bootstrap"
    }
    $optionsFile = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpPackageOptions-$([Guid]::NewGuid().ToString('N')).json"
    try {
        # PowerShell binds omitted strings as empty; preserve the typed optional defaults.
        $typed = @{}
        foreach ($name in $Options.Keys) {
            if ($Options[$name] -is [string] -and $Options[$name].Length -eq 0) { continue }
            $typed[$name] = $Options[$name]
        }
        [IO.File]::WriteAllText($optionsFile, ($typed | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
        & $bootstrap package --root $toolRoot --source-root $SourceRoot --package-options $optionsFile
        if ($LASTEXITCODE -ne 0) { throw "Package operation '$($Options.Operation)' failed with exit code $LASTEXITCODE." }
    }
    finally {
        if (Test-Path -LiteralPath $optionsFile) { Remove-Item -LiteralPath $optionsFile -Force }
    }
}

function Assert-PackageOutputPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$RepoRoot,
        [string[]]$Inputs = @()
    )
    Invoke-TypedPackage -SourceRoot $RepoRoot -Options @{
        Operation = 'AssertOutput'; OutputDirectory = $Path; Inputs = @($Inputs)
    }
}

function Publish-PackageRuntime {
    param(
        [ValidateSet('Cli', 'Mcp')][string]$Component,
        [string]$RepoRoot,
        [string]$Version,
        [string]$OutputDirectory,
        [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64'
    )
    Invoke-TypedPackage -SourceRoot $RepoRoot -Options @{
        Operation = 'PublishRuntime'; Component = $Component; Version = $Version
        OutputDirectory = $OutputDirectory; Architecture = $Architecture
    }
}

function Assert-PackageRuntimeArchitecture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('x64', 'arm64')][string]$Architecture
    )
    Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
        Operation = 'RuntimeArchitecture'; RuntimeExecutable = $Path; Architecture = $Architecture
    }
}

function Install-PackageOutput {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    Invoke-TypedPackage -SourceRoot (Split-Path $PSScriptRoot -Parent) -Options @{
        Operation = 'InstallOutput'; Source = $Source; Destination = $Destination
    }
}
