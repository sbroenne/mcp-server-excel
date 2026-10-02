[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$')][string]$Version,
    [Parameter(Mandatory)][ValidatePattern('^[\w.-]+/[\w.-]+$')][string]$Repository,
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$SourceCommit,
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$ReleaseCommit,
    [Parameter(Mandatory)][string]$MetadataPatch,
    [Parameter(Mandatory)][string]$AssetDirectory,
    [Parameter(Mandatory)][string]$PublishDirectory,
    [Parameter(Mandatory)][string]$NotesFile,
    [switch]$AllowMutableRepair
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$tag = "v$Version"

function Invoke-Gh {
    param([string[]]$Arguments, [switch]$AllowNotFound)
    $PSNativeCommandUseErrorActionPreference = $false
    $output = @(& gh @Arguments 2>&1)
    $text = $output | Out-String
    if ($LASTEXITCODE -ne 0) {
        if ($AllowNotFound -and $text -match 'HTTP 404') { return $null }
        throw "GitHub command failed: $text"
    }
    return $text
}

function Get-Release {
    $json = Invoke-Gh -Arguments @('api', "repos/$Repository/releases/tags/$tag") -AllowNotFound
    if ($null -eq $json) { return $null }
    $release = $json | ConvertFrom-Json
    if ($release.tag_name -cne $tag -or $release.draft -isnot [bool] -or
        $release.immutable -isnot [bool]) {
        throw 'GitHub returned an invalid release identity or state.'
    }
    return $release
}

function Assert-Assets {
    param($Release, [switch]$AllowMissing)
    $missing = @()
    foreach ($file in $payload) {
        $matches = @($Release.assets | Where-Object { $_.name -ceq $file.Name })
        if ($matches.Count -eq 0 -and $AllowMissing) {
            $missing += $file.FullName
            continue
        }
        if ($matches.Count -ne 1) { throw "Exactly one release asset is required: $($file.Name)" }
        $digest = 'sha256:' + (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($matches[0].digest -cne $digest) {
            throw "Missing or mismatched GitHub SHA-256 digest: $($file.Name)"
        }
    }
    return $missing
}

if (-not (Test-Path -LiteralPath $MetadataPatch -PathType Leaf) -or
    -not (Test-Path -LiteralPath $NotesFile -PathType Leaf)) {
    throw 'The exact metadata patch and release notes are required.'
}
if (Test-Path -LiteralPath $PublishDirectory) {
    throw 'Use a new empty publication directory; existing content is not overwritten.'
}
$requiredNames = @(
    "ExcelMcp-CLI-$Version-windows.zip",
    "ExcelMcp-MCP-Server-$Version-windows.zip",
    "excel-plugins-v$Version.zip",
    "excel-skills-v$Version.zip",
    "excel-mcp-$Version.vsix",
    "excel-mcp-$Version-win32-arm64.vsix",
    "excel-mcp-$Version.mcpb"
)
$files = @(Get-ChildItem -LiteralPath $AssetDirectory -Recurse -File)
$selected = foreach ($name in $requiredNames) {
    $matches = @($files | Where-Object { $_.Name -ceq $name })
    if ($matches.Count -ne 1) { throw "Exactly one prepared artifact is required: $name" }
    $matches[0]
}
New-Item -ItemType Directory -Path $PublishDirectory | Out-Null
foreach ($file in $selected) {
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $PublishDirectory $file.Name)
}
Copy-Item -LiteralPath $MetadataPatch -Destination (Join-Path $PublishDirectory 'release-metadata.patch')
$artifacts = @($selected | Sort-Object Name | ForEach-Object {
    [ordered]@{
        name = $_.Name
        sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
})
$inputs = [ordered]@{
    kind = 'build-input-record-not-attestation'
    repository = $Repository
    workflowPath = '.github/workflows/release.yml'
    sourceCommit = $SourceCommit
    metadataPatchSha256 = (Get-FileHash -LiteralPath $MetadataPatch -Algorithm SHA256).Hash.ToLowerInvariant()
    releaseCommit = $ReleaseCommit
    tag = $tag
    artifacts = $artifacts
}
$utf8 = [Text.UTF8Encoding]::new($false)
$record = ($inputs | ConvertTo-Json -Depth 5).Replace("`r`n", "`n") + "`n"
[IO.File]::WriteAllText((Join-Path $PublishDirectory 'RELEASE-INPUTS.json'), $record, $utf8)
$lines = @(Get-ChildItem -LiteralPath $PublishDirectory -File | Sort-Object Name | ForEach-Object {
    (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $_.Name
})
[IO.File]::WriteAllText((Join-Path $PublishDirectory 'SHA256SUMS'), ($lines -join "`n") + "`n", $utf8)
$payload = @(Get-ChildItem -LiteralPath $PublishDirectory -File | Sort-Object Name)

$release = Get-Release
if ($null -eq $release) {
    Write-Output "Creating draft $tag before uploading verified assets."
    Invoke-Gh -Arguments @('release', 'create', $tag, '--repo', $Repository, '--draft',
        '--verify-tag', '--target', $ReleaseCommit, '--title', "ExcelMcp $Version",
        '--notes-file', $NotesFile) | Out-Null
    $release = Get-Release
    if ($null -eq $release -or -not $release.draft) { throw 'GitHub did not create the expected draft.' }
}
if ($release.draft) {
    if ($release.immutable) { throw 'An immutable release cannot be edited as a draft.' }
    Invoke-Gh -Arguments (@('release', 'upload', $tag, '--repo', $Repository, '--clobber') + $payload.FullName) | Out-Null
    $release = Get-Release
    if ($null -eq $release -or -not $release.draft) { throw 'Release state changed during draft preparation.' }
    Assert-Assets -Release $release | Out-Null
    if (@($release.assets).Count -ne $payload.Count) { throw 'Unexpected draft assets must be reviewed before publication.' }
    Invoke-Gh -Arguments @('release', 'edit', $tag, '--repo', $Repository, '--draft=false',
        '--title', "ExcelMcp $Version", '--notes-file', $NotesFile) | Out-Null
    $release = Get-Release
    if ($null -eq $release -or $release.draft) { throw 'GitHub did not publish the verified draft.' }
    Assert-Assets -Release $release | Out-Null
    Write-Output "Published $tag after verifying every draft asset."
} else {
    $missing = @(Assert-Assets -Release $release -AllowMissing)
    if ($missing.Count -gt 0) {
        if ($release.immutable -or -not $AllowMutableRepair) {
            throw 'Published assets are missing. Immutable releases cannot be repaired; mutable repair requires explicit authorization.'
        }
        Write-Output "Repairing only missing assets on explicitly authorized mutable release $tag."
        Invoke-Gh -Arguments (@('release', 'upload', $tag, '--repo', $Repository) + $missing) | Out-Null
        Assert-Assets -Release (Get-Release) | Out-Null
    }
    Write-Output "Verified published $tag without replacing existing assets or editing release metadata."
}
