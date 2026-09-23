<#
.SYNOPSIS
    Compiles pending .changeset/*.md fragments into CHANGELOG.md for a release.

.DESCRIPTION
    Wraps `npx changeset version`, which consumes every pending changeset fragment,
    bumps the (bookkeeping-only) root package.json version, and inserts a new section
    at the top of CHANGELOG.md. This script then:
      1. Normalizes the changesets-generated version header to the Keep a Changelog
         style already used in this file: `## [X.Y.Z] - YYYY-MM-DD`.
      2. Synchronizes all persistent source-tree version metadata with the real
         release version. Build-time placeholder manifests remain unchanged.
      3. Extracts the newly-inserted section body to a separate file so it can be
         used verbatim as GitHub Release notes.

    Safe to run locally for a dry run: it mutates CHANGELOG.md, release metadata,
    and deletes consumed fragments in .changeset/, same as the real release step.

.PARAMETER Version
    The version being released, e.g. "1.9.1" (no leading "v").

.PARAMETER Date
    Release date in YYYY-MM-DD format. Defaults to today (UTC).

.PARAMETER RepoRoot
    Path to the repository root (where package.json and CHANGELOG.md live).
    Defaults to the current directory.

.PARAMETER OutputNotesPath
    Path to write the extracted release-notes body to. Defaults to
    "release_notes_body.md" in RepoRoot.

.EXAMPLE
    pwsh scripts/Build-Changelog.ps1 -Version 1.9.1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$Date = (Get-Date -AsUTC -Format 'yyyy-MM-dd'),

    [string]$RepoRoot = (Get-Location).Path,

    [string]$OutputNotesPath
)

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path $RepoRoot).Path
$changelogPath = Join-Path $RepoRoot 'CHANGELOG.md'
$packageJsonPath = Join-Path $RepoRoot 'package.json'
$updateReleaseVersionScript = Join-Path $PSScriptRoot 'Update-ReleaseVersionMetadata.ps1'
if (-not $OutputNotesPath) {
    $OutputNotesPath = Join-Path $RepoRoot 'release_notes_body.md'
}

if (-not (Test-Path $changelogPath)) {
    throw "CHANGELOG.md not found at $changelogPath"
}
if (-not (Test-Path $packageJsonPath)) {
    throw "package.json not found at $packageJsonPath (required to host the changesets tool)"
}
if (-not (Test-Path $updateReleaseVersionScript)) {
    throw "Release version metadata updater not found at $updateReleaseVersionScript"
}
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version '$Version' must be a plain semver value without a leading 'v' (e.g. 1.9.1)."
}

# --- Step 1: snapshot the changelog body (everything after the title line) before
# changesets mutates the file. changesets always inserts its new section
# immediately after line 1, leaving everything else untouched, so a suffix match
# after the run tells us exactly what it inserted.
$beforeLines = Get-Content -LiteralPath $changelogPath
if ($beforeLines.Count -lt 1) {
    throw "CHANGELOG.md is empty — expected at least a title line."
}
$titleLine = $beforeLines[0]
$beforeBody = ($beforeLines | Select-Object -Skip 1) -join "`n"

# --- Step 2: run changeset version from the repo root.
Push-Location $RepoRoot
try {
    & npx changeset version
    if ($LASTEXITCODE -ne 0) {
        throw "npx changeset version failed with exit code $LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

# --- Step 3: isolate the newly-inserted section using the first non-blank line
# from the previous changelog body as an anchor. Changesets can normalize legacy
# content while rewriting the file, so comparing the entire body byte-for-byte is
# too strict. The final file is reassembled from the untouched snapshot below.
$afterLines = Get-Content -LiteralPath $changelogPath
$anchorLine = $beforeLines |
    Select-Object -Skip 1 |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -First 1

if ($null -eq $anchorLine) {
    $newSection = (($afterLines | Select-Object -Skip 1) -join "`n").Trim("`r", "`n")
}
else {
    $anchorIndexes = for ($i = 1; $i -lt $afterLines.Count; $i++) {
        if ($afterLines[$i] -ceq $anchorLine) {
            $i
        }
    }

    if ($anchorIndexes.Count -ne 1) {
        throw "Could not uniquely locate the pre-existing CHANGELOG.md anchor '$anchorLine' after " +
              "'npx changeset version' ran. Aborting without writing changes to avoid corrupting the changelog."
    }

    $newSectionLines = if ($anchorIndexes[0] -gt 1) {
        $afterLines[1..($anchorIndexes[0] - 1)]
    }
    else {
        @()
    }
    $newSection = ($newSectionLines -join "`n").Trim("`r", "`n")
}

if ([string]::IsNullOrWhiteSpace($newSection)) {
    & $updateReleaseVersionScript -RepoRoot $RepoRoot -Version $Version
    Write-Output 'No pending changesets found — nothing to add to the changelog.'
    Set-Content -LiteralPath $OutputNotesPath -Value "_No changes recorded for this release._" -NoNewline
    exit 0
}

# --- Step 4: normalize the first non-blank line (changesets' own version header,
# e.g. "## excelmcp@1.9.1" or "## 1.9.1") to the Keep a Changelog style used here.
$newSectionLines = $newSection -split "`r?`n"
$headerIdx = 0
while ($headerIdx -lt $newSectionLines.Count -and [string]::IsNullOrWhiteSpace($newSectionLines[$headerIdx])) {
    $headerIdx++
}
if ($headerIdx -ge $newSectionLines.Count -or $newSectionLines[$headerIdx] -notmatch '^##\s') {
    throw "Expected the changesets-generated section to start with a '## ' version header, got: '$($newSectionLines[$headerIdx])'"
}
$newSectionLines[$headerIdx] = "## [$Version] - $Date"
$newSection = ($newSectionLines -join "`n").Trim()

# --- Step 5: reassemble CHANGELOG.md as: title + preamble + new section + prior
# versions. The preamble is any prose between the title and the first `## [` version
# heading. Inserting the new section *after* the preamble — rather than immediately
# after the title line — keeps the preamble pinned at the top. Inserting after the
# title (the previous behavior) pushed the preamble one release further down every
# time, which is how it eventually ended up buried among old version entries.
$beforeBodyLines = $beforeBody -split "`n"
$firstVersionIdx = -1
for ($i = 0; $i -lt $beforeBodyLines.Count; $i++) {
    if ($beforeBodyLines[$i] -match '^##\s+\[') { $firstVersionIdx = $i; break }
}
if ($firstVersionIdx -le 0) {
    # No preamble (body starts at, or before, the first version heading).
    $preambleBlock = ''
    $priorVersions = $beforeBody.Trim("`r", "`n")
}
else {
    $preambleBlock = (($beforeBodyLines[0..($firstVersionIdx - 1)]) -join "`n").Trim("`r", "`n")
    $priorVersions = (($beforeBodyLines[$firstVersionIdx..($beforeBodyLines.Count - 1)]) -join "`n").Trim("`r", "`n")
}
$sections = @($titleLine)
if (-not [string]::IsNullOrWhiteSpace($preambleBlock)) { $sections += $preambleBlock }
$sections += $newSection
if (-not [string]::IsNullOrWhiteSpace($priorVersions)) { $sections += $priorVersions }
$finalContent = ($sections -join "`n`n").TrimEnd() + "`n"
Set-Content -LiteralPath $changelogPath -Value $finalContent -NoNewline

# --- Step 6: synchronize every persistent source-tree version. This runs after
# changesets so its calculated package.json bump cannot override the selected
# release version.
& $updateReleaseVersionScript -RepoRoot $RepoRoot -Version $Version

# --- Step 7: write the release-notes body (verbatim section content, no header
# duplication needed since GitHub Release titles already carry the version).
$notesBody = ($newSectionLines[($headerIdx + 1)..($newSectionLines.Count - 1)] -join "`n").Trim()
if ([string]::IsNullOrWhiteSpace($notesBody)) {
    $notesBody = '_No notable changes recorded for this release._'
}
Set-Content -LiteralPath $OutputNotesPath -Value $notesBody -NoNewline

Write-Output "CHANGELOG.md updated with ## [$Version] - $Date"
Write-Output "Release notes body written to $OutputNotesPath"
