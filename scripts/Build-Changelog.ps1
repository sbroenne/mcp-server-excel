<#
.SYNOPSIS
    Compiles pending .changeset/*.md fragments into CHANGELOG.md for a release.

.DESCRIPTION
    Wraps `npx changeset version`, which consumes every pending changeset fragment,
    bumps the (bookkeeping-only) root package.json version, and inserts a new section
    at the top of CHANGELOG.md. This script then:
      1. Normalizes the changesets-generated version header to the Keep a Changelog
         style already used in this file: `## [X.Y.Z] - YYYY-MM-DD`.
      2. Forces the root package.json version to exactly match the real release
         version (the source of truth remains the git tag / workflow input — this
         package.json only exists to host the changesets tool and must not drift).
      3. Extracts the newly-inserted section body to a separate file so it can be
         used verbatim as GitHub Release notes.

    Safe to run locally for a dry run: it mutates CHANGELOG.md, package.json, and
    deletes consumed fragments in .changeset/, same as the real release step. Use
    `git checkout -- CHANGELOG.md package.json` and `git clean -fd .changeset` to
    revert a local dry run.

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
if (-not $OutputNotesPath) {
    $OutputNotesPath = Join-Path $RepoRoot 'release_notes_body.md'
}

if (-not (Test-Path $changelogPath)) {
    throw "CHANGELOG.md not found at $changelogPath"
}
if (-not (Test-Path $packageJsonPath)) {
    throw "package.json not found at $packageJsonPath (required to host the changesets tool)"
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

# --- Step 3: isolate the newly-inserted section via suffix match against the
# untouched body captured in Step 1.
$afterLines = Get-Content -LiteralPath $changelogPath
$afterRest = ($afterLines | Select-Object -Skip 1) -join "`n"

$idx = $afterRest.IndexOf($beforeBody, [System.StringComparison]::Ordinal)
if ($idx -lt 0) {
    throw "Could not locate the pre-existing CHANGELOG.md content after 'npx changeset version' ran. " +
          "Aborting without writing changes to avoid corrupting the changelog. " +
          "This usually means CHANGELOG.md's existing content was hand-edited in a way changesets doesn't expect."
}
$newSection = $afterRest.Substring(0, $idx).Trim("`r", "`n")

if ([string]::IsNullOrWhiteSpace($newSection)) {
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

# --- Step 5: reassemble CHANGELOG.md: title + normalized new section + untouched body.
$beforeBodyTrimmed = $beforeBody.TrimStart("`r", "`n")
$finalContent = "$titleLine`n`n$newSection`n`n$beforeBodyTrimmed".TrimEnd() + "`n"
Set-Content -LiteralPath $changelogPath -Value $finalContent -NoNewline

# --- Step 6: keep package.json's bookkeeping version in exact sync with the real
# release version (changesets' own auto-bump may not match if a contributor picked
# a different bump type than the maintainer ultimately released).
$packageJson = Get-Content -LiteralPath $packageJsonPath -Raw | ConvertFrom-Json
$packageJson.version = $Version
# Preserve key order/formatting reasonably well: ConvertTo-Json + npm-style 2-space indent.
($packageJson | ConvertTo-Json -Depth 10) -replace "`r?`n", "`n" | Set-Content -LiteralPath $packageJsonPath -NoNewline
Add-Content -LiteralPath $packageJsonPath -Value "`n"

# --- Step 7: write the release-notes body (verbatim section content, no header
# duplication needed since GitHub Release titles already carry the version).
$notesBody = ($newSectionLines[($headerIdx + 1)..($newSectionLines.Count - 1)] -join "`n").Trim()
if ([string]::IsNullOrWhiteSpace($notesBody)) {
    $notesBody = '_No notable changes recorded for this release._'
}
Set-Content -LiteralPath $OutputNotesPath -Value $notesBody -NoNewline

Write-Output "CHANGELOG.md updated with ## [$Version] - $Date"
Write-Output "Release notes body written to $OutputNotesPath"
