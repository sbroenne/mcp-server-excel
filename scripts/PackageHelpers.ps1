function Assert-PackageOutputPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$RepoRoot,
        [string[]]$Inputs = @()
    )
    $output = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $repo = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $separator = [IO.Path]::DirectorySeparatorChar
    $comparison = if ($IsWindows) { [StringComparison]::OrdinalIgnoreCase } else { [StringComparison]::Ordinal }
    $allowed = @(
        (Join-Path $repo 'artifacts'),
        (Join-Path $repo 'plugins'),
        (Join-Path $repo 'mcpb/artifacts')
    )
    $trustedBaselines = @(
        $repo,
        [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
    )
    function Test-PathWithin([string]$Candidate, [string]$Parent) {
        return $Candidate.Equals($Parent, $comparison) -or
            $Candidate.StartsWith("$Parent$separator", $comparison)
    }
    if ($output -eq [IO.Path]::GetPathRoot($Path).TrimEnd('\', '/') -or
        (Test-PathWithin $repo $output) -or
        ((Test-PathWithin $output $repo) -and
            -not @($allowed | Where-Object { Test-PathWithin $output $_ }).Count)) {
        throw "Unsafe package output directory: $output"
    }
    foreach ($inputPath in $Inputs) {
        if (-not $inputPath) { continue }
        $inputFull = [IO.Path]::GetFullPath($inputPath, $RepoRoot).TrimEnd('\', '/')
        if ((Test-PathWithin $inputFull $output) -or (Test-PathWithin $output $inputFull)) {
            throw "Package output overlaps a prepared input: $output"
        }
    }
    for ($ancestor = $output; $ancestor; $ancestor = Split-Path $ancestor -Parent) {
        if ((Test-Path -LiteralPath $ancestor) -and
            ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -and
            -not @($trustedBaselines | Where-Object { Test-PathWithin $_ $ancestor }).Count) {
            throw "Package output must not traverse a link: $ancestor"
        }
    }
}

function Publish-PackageRuntime {
    param(
        [ValidateSet('Cli', 'Mcp')][string]$Component,
        [string]$RepoRoot,
        [string]$Version,
        [string]$OutputDirectory,
        [ValidateSet('x64', 'arm64')][string]$Architecture = 'x64',
        [ValidateSet('win-x64', 'win-arm64', 'osx-arm64')]
        [string]$RuntimeIdentifier = 'win-x64'
    )
    if (-not $PSBoundParameters.ContainsKey('RuntimeIdentifier')) {
        $RuntimeIdentifier = "win-$Architecture"
    }
    $projectName = if ($Component -eq 'Cli') { 'CLI' } else { 'McpServer' }
    dotnet publish (Join-Path $RepoRoot "src\ExcelMcp.$projectName\ExcelMcp.$projectName.csproj") `
        -c Release -r $RuntimeIdentifier --self-contained true -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true -p:PublishTrimmed=false `
        -p:PublishReadyToRun=false -p:NuGetAudit=false "-p:Version=$Version" `
        -o $OutputDirectory --verbosity minimal
    if ($LASTEXITCODE -ne 0) { throw "$Component runtime publish failed with exit code $LASTEXITCODE." }
}

function Assert-PackageRuntimeArchitecture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('x64', 'arm64')][string]$Architecture
    )
    $stream = [IO.File]::OpenRead($Path)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($stream.Length -lt 64 -or $reader.ReadUInt16() -ne 0x5a4d) {
            throw "Runtime is not a Windows executable: $Path"
        }
        $stream.Position = 0x3c
        $headerOffset = $reader.ReadUInt32()
        if ($headerOffset -lt 64 -or $headerOffset -gt $stream.Length - 6) {
            throw "Runtime has an invalid executable header: $Path"
        }
        $stream.Position = $headerOffset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Runtime has an invalid PE signature: $Path"
        }
        $machine = $reader.ReadUInt16()
        $expected = if ($Architecture -eq 'arm64') { 0xaa64 } else { 0x8664 }
        if ($machine -ne $expected) {
            throw ("Runtime machine type 0x{0:x4} does not match {1}: {2}" -f $machine, $Architecture, $Path)
        }
    } finally {
        $reader.Dispose()
        $stream.Dispose()
    }
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
            $installationError = $_
            if (-not (Test-Path -LiteralPath $backup)) { throw $installationError }
            try {
                Move-Item -LiteralPath $backup -Destination $Destination
            } catch {
                $recoveryError = $_
                $destinationState = if (Test-Path -LiteralPath $Destination) {
                    "Destination exists at '$Destination'."
                } else {
                    "Destination is absent at '$Destination'."
                }
                $backupState = if (Test-Path -LiteralPath $backup) {
                    "Recovery backup retained at '$backup'."
                } else {
                    "Recovery backup is absent from '$backup'."
                }
                throw "Package installation failed: $($installationError.Exception.Message). Recovery also failed: $($recoveryError.Exception.Message). $destinationState $backupState"
            }
            throw $installationError
        }
        if (Test-Path -LiteralPath $backup) {
            try {
                Remove-Item -LiteralPath $backup -Recurse -Force
            } catch {
                Write-Warning "Package output installed successfully at '$Destination', but backup cleanup failed: $($_.Exception.Message). Backup retained at '$backup'."
            }
        }
    } finally {
        if (Test-Path -LiteralPath $temporary) {
            try {
                Remove-Item -LiteralPath $temporary -Recurse -Force
            } catch {
                Write-Warning "Temporary package output cleanup failed: $($_.Exception.Message). Temporary output retained at '$temporary'."
            }
        }
    }
}
