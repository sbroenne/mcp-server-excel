#Requires -Version 7.4

function Initialize-ExtensionPackage {
    param([string]$RepoRoot, [string]$Version, [string]$SkillsDirectory, [switch]$SkipExtensionTests)
    $stage = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpextension-$([Guid]::NewGuid().ToString('N'))"
    [IO.Directory]::CreateDirectory($stage) | Out-Null
    try {
        foreach ($entry in [IO.Directory]::EnumerateFileSystemEntries((Join-Path $RepoRoot 'vscode-extension'))) {
            $name = [IO.Path]::GetFileName($entry)
            if ($name -in @('node_modules', 'bin', 'out', 'skills') -or [IO.Path]::GetExtension($entry) -eq '.vsix') { continue }
            Copy-PackageItem $entry (Join-Path $stage $name)
        }
        $skill = Join-Path $stage 'skills\excel-mcp-report-formatting'
        Copy-PackageItem (Join-Path $SkillsDirectory 'excel-mcp-report-formatting') $skill
        [IO.File]::WriteAllText((Join-Path $skill 'VERSION'), $Version)
        Copy-PackageItem (Join-Path $RepoRoot 'CHANGELOG.md') (Join-Path $stage 'CHANGELOG.md')
        $path = Join-Path $stage 'package.json'
        $manifest = Read-PackageJson $path
        $manifest.version = $Version
        # Only this disposable, already-compiled source skips vsce's repeated compilation.
        $manifest.scripts.Remove('vscode:prepublish') | Out-Null
        Write-PackageJson $path $manifest
        Invoke-PackageCommand $stage 'npm' @('ci', '--ignore-scripts') | Out-Host
        Invoke-PackageCommand $stage 'npm' @('run', 'compile') | Out-Host
        Invoke-PackageCommand $stage 'npm' @('run', 'lint') | Out-Host
        if (-not $SkipExtensionTests) {
            Invoke-PackageCommand $stage 'npm' @('run', 'typecheck:tests') | Out-Host
            Invoke-PackageCommand $stage 'npm' @('test') | Out-Host
        }
        return $stage
    } catch {
        $primary = $_
        try { Remove-PackageStaging $stage }
        catch { throw [AggregateException]::new('Extension source checks and cleanup failed.', [Exception[]]@($primary.Exception, $_.Exception)) }
        throw
    }
}

function Read-VsixEntry {
    param([IO.Compression.ZipArchive]$Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { throw "VSIX is missing $Name." }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Assert-ExtensionPackage {
    param([string]$Path, [string]$Stage, [string]$Architecture, [string]$Version)
    $target = "win32-$Architecture"
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($name in 'extension/bin/Sbroenne.ExcelMcp.McpServer.exe', 'extension/out/extension.js', 'extension/out/prerequisites.js') {
            if (-not $zip.GetEntry($name)) { throw "VSIX is missing $name." }
        }
        $inspection = Join-Path ([IO.Path]::GetDirectoryName($Path)) "$target-server-inspection.exe"
        try {
            [IO.Compression.ZipFileExtensions]::ExtractToFile($zip.GetEntry('extension/bin/Sbroenne.ExcelMcp.McpServer.exe'), $inspection, $false)
            Assert-PackageRuntimeArchitecture $inspection $Architecture
        } finally { Remove-PackageItem $inspection }
        foreach ($file in [IO.Directory]::EnumerateFiles((Join-Path $Stage 'skills\excel-mcp-report-formatting'), '*', [IO.SearchOption]::AllDirectories)) {
            $relative = [IO.Path]::GetRelativePath($Stage, $file).Replace('\', '/')
            if (-not $zip.GetEntry("extension/$relative")) { throw "VSIX is missing $relative." }
        }
        if ((Read-VsixEntry $zip 'extension/skills/excel-mcp-report-formatting/VERSION').Trim() -cne $Version) { throw 'VSIX skill version does not match the package.' }
        $manifest = Read-VsixEntry $zip 'extension/package.json' | ConvertFrom-Json -AsHashtable
        if ($manifest.version -cne $Version -or ($manifest.extensionKind -join ',') -cne 'ui' -or ($manifest.os -join ',') -cne 'win32') {
            throw 'VSIX version or local Windows host metadata is incorrect.'
        }
        $metadata = [xml](Read-VsixEntry $zip 'extension.vsixmanifest')
        if ($metadata.SelectSingleNode("//*[local-name()='Identity']").GetAttribute('TargetPlatform') -cne $target) { throw "VSIX target does not match $target." }
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -cmatch '^extension/(?:node_modules|tests|scripts|\.vitest|coverage|out/tests)/|^extension/(?:vitest\.config\.|tsconfig(?:\.test)?\.json$|bin/excelcli)|^extension/(?:.*/)?(?:AGENTS|CLAUDE)\.md$') {
                throw "VSIX contains development files or the CLI: $($entry.FullName)"
            }
        }
        Write-Host "Verified VSIX target: $target"
    } finally { $zip.Dispose() }
}

function Complete-ExtensionPackages {
    param([string]$Stage, [string]$Version, [string]$OutputDirectory, [Collections.IDictionary]$Prepared)
    $bundled = Join-Path $Stage 'bin\Sbroenne.ExcelMcp.McpServer.exe'
    foreach ($architecture in 'x64', 'arm64') {
        $runtime = if ($architecture -eq 'x64') { $Prepared['Mcp'] } else { $Prepared['Mcp-arm64'] }
        Copy-PackageItem $runtime $bundled
        $filename = if ($architecture -eq 'x64') { "excel-mcp-$Version.vsix" } else { "excel-mcp-$Version-win32-arm64.vsix" }
        $path = Join-Path $OutputDirectory $filename
        Invoke-PackageCommand $Stage 'npm' @('exec', '--', 'vsce', 'package', '--no-dependencies', '--target', "win32-$architecture", '--out', $path) | Out-Host
        Assert-ExtensionPackage $path $Stage $architecture $Version
    }
    $nativeRuntime = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [Runtime.InteropServices.Architecture]::Arm64) { $Prepared['Mcp-arm64'] } else { $Prepared['Mcp'] }
    Copy-PackageItem $nativeRuntime $bundled
    $debug = Join-Path $OutputDirectory 'extension'
    foreach ($entry in [IO.Directory]::EnumerateFileSystemEntries($Stage)) {
        if ([IO.Path]::GetFileName($entry) -ne 'node_modules') { Copy-PackageItem $entry (Join-Path $debug ([IO.Path]::GetFileName($entry))) }
    }
}
