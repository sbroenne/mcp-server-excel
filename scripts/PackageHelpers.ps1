#Requires -Version 7.4

function Test-PackagePathContains {
    param([string]$Directory, [string]$Path)
    $directoryPath = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Directory))
    $full = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Path))
    $full.Equals($directoryPath, [StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith("$directoryPath$([IO.Path]::DirectorySeparatorChar)", [StringComparison]::OrdinalIgnoreCase)
}

function Assert-PackageNoLinks {
    param([string]$Path)
    for ($current = [IO.Path]::GetFullPath($Path); $current; $current = [IO.Path]::GetDirectoryName($current)) {
        if ([IO.Path]::Exists($current) -and ([IO.File]::GetAttributes($current) -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Package output must not traverse a link: $current"
        }
    }
}

function Assert-PackageOutputPath {
    param([string]$Path, [string]$RepoRoot, [string[]]$Inputs = @())
    $output = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Path, $RepoRoot))
    $allowed = @('artifacts', 'plugins', 'mcpb\artifacts') |
        Where-Object { Test-PackagePathContains (Join-Path $RepoRoot $_) $output }
    if ($output -eq [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetPathRoot($output)) -or
        (Test-PackagePathContains $output $RepoRoot) -or
        ((Test-PackagePathContains $RepoRoot $output) -and -not $allowed)) {
        throw "Unsafe package output directory: $output"
    }
    foreach ($inputPath in $Inputs) {
        if (-not $inputPath) { continue }
        $full = [IO.Path]::GetFullPath($inputPath, $RepoRoot)
        if ((Test-PackagePathContains $output $full) -or (Test-PackagePathContains $full $output)) {
            throw "Package output overlaps a prepared input: $output"
        }
    }
    Assert-PackageNoLinks $output
}

function Copy-PackageItem {
    param([string]$Source, [string]$Destination)
    if ([IO.File]::GetAttributes($Source) -band [IO.FileAttributes]::ReparsePoint) {
        throw "Package source must not be a link: $Source"
    }
    Assert-PackageNoLinks $Destination
    if ([IO.Directory]::Exists($Source)) {
        [IO.Directory]::CreateDirectory($Destination) | Out-Null
        foreach ($entry in [IO.Directory]::EnumerateFileSystemEntries($Source)) {
            Copy-PackageItem $entry (Join-Path $Destination ([IO.Path]::GetFileName($entry)))
        }
    } else {
        [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Destination)) | Out-Null
        [IO.File]::Copy($Source, $Destination, $true)
    }
}

function Remove-PackageItem {
    param([string]$Path)
    Assert-PackageNoLinks $Path
    if ([IO.Directory]::Exists($Path)) { [IO.Directory]::Delete($Path, $true) }
    elseif ([IO.File]::Exists($Path)) { [IO.File]::Delete($Path) }
}

function Move-PackageItem {
    param([string]$Source, [string]$Destination)
    if ([IO.Directory]::Exists($Source)) { [IO.Directory]::Move($Source, $Destination) }
    else { [IO.File]::Move($Source, $Destination) }
}

function Install-PackageOutput {
    param(
        [string]$Source, [string]$Destination,
        [scriptblock]$Move = { param($from, $to) Move-PackageItem $from $to },
        [scriptblock]$Delete = { param($path) Remove-PackageItem $path }
    )
    Assert-PackageNoLinks $Destination
    $temporary = "$Destination.$([Guid]::NewGuid().ToString('N')).tmp"
    $backup = "$temporary.bak"
    try {
        Copy-PackageItem $Source $temporary
        if ([IO.Path]::Exists($Destination)) { & $Move $Destination $backup }
        try { & $Move $temporary $Destination }
        catch {
            $primary = $_
            if ([IO.Path]::Exists($backup)) {
                try { & $Move $backup $Destination }
                catch {
                    throw [AggregateException]::new(
                        "Package installation failed: $primary. Recovery also failed: $_. Destination present: $([IO.Path]::Exists($Destination)) at '$Destination'. Recovery backup retained at '$backup'.",
                        [Exception[]]@($primary.Exception, $_.Exception))
                }
            }
            throw
        }
        if ([IO.Path]::Exists($backup)) {
            try { & $Delete $backup }
            catch { Write-Warning "Package output installed successfully at '$Destination', but backup cleanup failed: $_. Backup retained at '$backup'." }
        }
    } finally {
        if ([IO.Path]::Exists($temporary)) {
            try { & $Delete $temporary }
            catch { Write-Warning "Temporary package output cleanup failed: $_. Temporary output retained at '$temporary'." }
        }
    }
}

function Remove-PackageStaging {
    param(
        [string]$Path, [double]$TimeoutMilliseconds = 120000, [double]$RetryMilliseconds = 500,
        [scriptblock]$Remove = { param($path) Remove-PackageItem $path }
    )
    if ($TimeoutMilliseconds -lt 0 -or $RetryMilliseconds -lt 0) { throw 'Staging timeout and retry interval must be non-negative.' }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $attempts = 0
    while ([IO.Directory]::Exists($Path)) {
        $attempts++
        $lastError = $null
        try { & $Remove $Path }
        catch [IO.IOException] { $lastError = $_ }
        catch [UnauthorizedAccessException] { $lastError = $_ }
        if (-not [IO.Directory]::Exists($Path)) { return }
        $remaining = $TimeoutMilliseconds - $clock.Elapsed.TotalMilliseconds
        if ($remaining -le 0) {
            throw [IO.IOException]::new(
                "Failed to remove package staging directory '$Path' after $attempts attempt(s) within $TimeoutMilliseconds ms. Stale staging remains. Last error: $lastError")
        }
        if ($RetryMilliseconds -gt 0) { Start-Sleep -Milliseconds ([int][Math]::Min($RetryMilliseconds, $remaining)) }
    }
}

function Invoke-PackageStage {
    param([string]$StageName, [scriptblock]$Action)
    $stage = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp$StageName-$([Guid]::NewGuid().ToString('N'))"
    [IO.Directory]::CreateDirectory($stage) | Out-Null
    $primary = $null
    try { & $Action $stage }
    catch { $primary = $_; throw }
    finally {
        try { Remove-PackageStaging $stage }
        catch {
            if ($primary) { throw [AggregateException]::new('Packaging and staging cleanup failed.', [Exception[]]@($primary.Exception, $_.Exception)) }
            throw
        }
    }
}

function Get-PackageVersion {
    param([string]$Version)
    if ([string]::IsNullOrWhiteSpace($Version) -or $Version.Trim() -cnotmatch '^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$') {
        throw 'Version is required. Pass a valid package version.'
    }
    $Version.Trim()
}

function Read-PackageJson {
    param([string]$Path)
    $value = [IO.File]::ReadAllText($Path) | ConvertFrom-Json -AsHashtable
    if ($value -isnot [Collections.IDictionary]) { throw "Expected a JSON object: $Path" }
    return $value
}

function Write-PackageJson {
    param([string]$Path, [Collections.IDictionary]$Value)
    [IO.File]::WriteAllText($Path, (($Value | ConvertTo-Json -Depth 50).Replace("`r`n", "`n").Replace("`r", "`n")))
}

function Assert-PackageRuntimeArchitecture {
    param([string]$Path, [ValidateSet('x64', 'arm64')][string]$Architecture)
    $reader = [IO.BinaryReader]::new([IO.File]::OpenRead($Path))
    try {
        if ($reader.BaseStream.Length -lt 64 -or $reader.ReadUInt16() -ne 0x5a4d) { throw "Runtime is not a Windows executable: $Path" }
        $reader.BaseStream.Position = 0x3c
        $offset = $reader.ReadUInt32()
        if ($offset -lt 64 -or $offset -gt $reader.BaseStream.Length - 6) { throw "Runtime has an invalid executable header: $Path" }
        $reader.BaseStream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw "Runtime has an invalid PE signature: $Path" }
        $machine = $reader.ReadUInt16()
        $expected = if ($Architecture -eq 'x64') { 0x8664 } else { 0xaa64 }
        if ($machine -ne $expected) { throw ('Runtime machine type 0x{0:x4} does not match {1}: {2}' -f $machine, $Architecture, $Path) }
    } finally { $reader.Dispose() }
}

function Invoke-PackageCommand {
    param([string]$Directory, [string]$Executable, [string[]]$Arguments, [double]$TimeoutSeconds = 1200)
    if ($Executable -eq 'npm' -and $IsWindows) {
        $node = (Get-Command node.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        $npm = Join-Path (Split-Path $node -Parent) 'node_modules\npm\bin\npm-cli.js'
        if (-not [IO.File]::Exists($npm)) { throw "The Node.js npm CLI is missing: $npm" }
        $Executable = $node
        $Arguments = @($npm) + $Arguments
    }
    $info = [Diagnostics.ProcessStartInfo]::new($Executable)
    $info.WorkingDirectory = $Directory
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.CreateNoWindow = $true
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    foreach ($key in @($info.Environment.Keys)) {
        if ($key.StartsWith('GIT_', [StringComparison]::OrdinalIgnoreCase)) { $info.Environment.Remove($key) | Out-Null }
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw "Could not start $Executable." }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $completion = [Threading.Tasks.Task]::WhenAll([Threading.Tasks.Task[]]@($process.WaitForExitAsync(), $stdout, $stderr))
        try { $null = $completion.WaitAsync([TimeSpan]::FromSeconds($TimeoutSeconds)).GetAwaiter().GetResult() }
        catch {
            $primary = $_
            try {
                if (-not $process.HasExited) { $process.Kill($true) }
                $null = $completion.WaitAsync([TimeSpan]::FromSeconds(10)).GetAwaiter().GetResult()
            } catch {
                throw [AggregateException]::new("$Executable failed and process/output cleanup failed.", [Exception[]]@($primary.Exception, $_.Exception))
            }
            throw "$Executable exceeded its hard deadline. $($stdout.Result) $($stderr.Result)"
        }
        if ($process.ExitCode -ne 0) { throw "$Executable failed with exit code $($process.ExitCode).`n$($stdout.Result)`n$($stderr.Result)" }
        if ($stderr.Result) { [Console]::Error.Write($stderr.Result) }
        return $stdout.Result
    } finally { $process.Dispose() }
}

function Expand-PackageTar {
    param([string]$Archive, [string]$Destination)
    Assert-PackageNoLinks $Destination
    [IO.Directory]::CreateDirectory($Destination) | Out-Null
    $file = [IO.File]::OpenRead($Archive)
    try {
        $gzip = [IO.Compression.GZipStream]::new($file, [IO.Compression.CompressionMode]::Decompress)
        try {
            $reader = [System.Formats.Tar.TarReader]::new($gzip)
            try {
                $paths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                while ($entry = $reader.GetNextEntry()) {
                    $path = [IO.Path]::GetFullPath($entry.Name.Replace('/', [IO.Path]::DirectorySeparatorChar), $Destination)
                    if (-not (Test-PackagePathContains $Destination $path) -or -not $paths.Add($path)) {
                        throw "Unsafe or duplicate npm archive entry: $($entry.Name)"
                    }
                    Assert-PackageNoLinks $path
                    switch ($entry.EntryType.ToString()) {
                        'Directory' { [IO.Directory]::CreateDirectory($path) | Out-Null }
                        { $_ -in 'RegularFile', 'V7RegularFile' } {
                            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path)) | Out-Null
                            $entry.ExtractToFile($path, $false)
                        }
                        default { throw "Unsupported npm archive entry: $($entry.Name) ($($entry.EntryType))." }
                    }
                }
            } finally { $reader.Dispose() }
        } finally { $gzip.Dispose() }
    } finally { $file.Dispose() }
}

function New-StandalonePackage {
    param([string]$Component, [string]$Version, [string]$Executable, [string]$Architecture, [string]$OutputDirectory, [string]$RepoRoot)
    $Version = Get-PackageVersion $Version
    Assert-PackageRuntimeArchitecture $Executable $Architecture
    $baseName, $executableName = switch ($Component) {
        'Cli' { 'ExcelMcp-CLI'; 'excelcli.exe' }
        'Mcp' { 'ExcelMcp-MCP-Server'; 'mcp-excel.exe' }
        default { throw "Unknown standalone package component: $Component" }
    }
    Invoke-PackageStage 'standalone' {
        param($stage)
        Copy-PackageItem $Executable (Join-Path $stage $executableName)
        foreach ($name in 'README.md', 'LICENSE', 'CHANGELOG.md') { Copy-PackageItem (Join-Path $RepoRoot $name) (Join-Path $stage $name) }
        $suffix = if ($Architecture -eq 'x64') { 'windows' } else { 'windows-arm64' }
        $archive = Join-Path $OutputDirectory "$baseName-$Version-$suffix.zip"
        [IO.Compression.ZipFile]::CreateFromDirectory($stage, $archive)
        $archive
    }
}
