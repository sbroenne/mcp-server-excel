function Assert-PackageOutputPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$RepoRoot,
        [string[]]$Inputs = @()
    )
    $output = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $repo = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $allowed = @("$repo\artifacts", "$repo\plugins", "$repo\mcpb\artifacts")
    if ($output -eq [IO.Path]::GetPathRoot($Path).TrimEnd('\', '/') -or
        $repo -eq $output -or $repo.StartsWith("$output\", [StringComparison]::OrdinalIgnoreCase) -or
        ($output.StartsWith("$repo\", [StringComparison]::OrdinalIgnoreCase) -and
            -not @($allowed | Where-Object { $output -eq $_ -or $output.StartsWith("$_\", [StringComparison]::OrdinalIgnoreCase) }).Count)) {
        throw "Unsafe package output directory: $output"
    }
    foreach ($inputPath in $Inputs) {
        if (-not $inputPath) { continue }
        $inputFull = [IO.Path]::GetFullPath($inputPath, $RepoRoot).TrimEnd('\', '/')
        if ($output -eq $inputFull -or $inputFull.StartsWith("$output\", [StringComparison]::OrdinalIgnoreCase) -or
            $output.StartsWith("$inputFull\", [StringComparison]::OrdinalIgnoreCase)) {
            throw "Package output overlaps a prepared input: $output"
        }
    }
    for ($ancestor = $output; $ancestor; $ancestor = Split-Path $ancestor -Parent) {
        if ((Test-Path -LiteralPath $ancestor) -and
            ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Package output must not traverse a link: $ancestor"
        }
    }
}

function Publish-PackageRuntime {
    param(
        [ValidateSet('Cli', 'Mcp')][string]$Component,
        [string]$RepoRoot,
        [string]$Version,
        [string]$OutputDirectory
    )
    $projectName = if ($Component -eq 'Cli') { 'CLI' } else { 'McpServer' }
    dotnet publish (Join-Path $RepoRoot "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj") `
        -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false `
        -p:PublishReadyToRun=false -p:NuGetAudit=false "-p:Version=$Version" `
        -o $OutputDirectory --verbosity minimal
    if ($LASTEXITCODE -ne 0) { throw "$Component runtime publish failed with exit code $LASTEXITCODE." }
}

function Install-PackageOutput {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )
    if ((Test-Path -LiteralPath $Destination) -and
        ((Get-Item -LiteralPath $Destination -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "Package destination must not be a link: $Destination"
    }
    $temporary = "$Destination.$([Guid]::NewGuid().ToString('N')).tmp"
    $backup = "$temporary.bak"
    try {
        Copy-Item -LiteralPath $Source -Destination $temporary -Recurse
        if (Test-Path -LiteralPath $Destination) {
            Move-Item -LiteralPath $Destination -Destination $backup
        }
        try {
            Move-Item -LiteralPath $temporary -Destination $Destination
        } catch {
            if (Test-Path -LiteralPath $backup) { Move-Item -LiteralPath $backup -Destination $Destination }
            throw
        }
        if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Recurse -Force }
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
    }
}
