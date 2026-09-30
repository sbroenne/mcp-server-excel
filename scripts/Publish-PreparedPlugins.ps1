[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceDirectory,
    [Parameter(Mandatory)][string]$PublishedRepoDirectory,
    [Parameter(Mandatory)][string]$BuiltPluginsDirectory,
    [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][string]$SourceCommit,
    [switch]$ManualRepair,
    [switch]$Preview
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
$tag = "v$Version"
if ($Version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$' -or $SourceCommit -notmatch '^[a-f0-9]{40}$') {
    throw 'An exact release version and source commit are required.'
}
$source = (Resolve-Path -LiteralPath $SourceDirectory).Path
$published = (Resolve-Path -LiteralPath $PublishedRepoDirectory).Path
$built = (Resolve-Path -LiteralPath $BuiltPluginsDirectory).Path
if ((git -C $source rev-parse "refs/tags/$tag^{commit}") -ne $SourceCommit) { throw 'Source release tag/commit mismatch.' }
if ((git -C $source rev-parse HEAD) -ne $SourceCommit -or @(git -C $source status --porcelain).Count) {
    throw 'Source checkout must be clean and match the exact release commit.'
}
if (@(git -C $published status --porcelain).Count) { throw 'Published checkout must be clean.' }
$baselineCommit = git -C $published rev-parse HEAD
$manifestPath = Join-Path $published '.github\plugin\marketplace.json'
if (-not (Test-Path -LiteralPath $manifestPath)) { $manifestPath = Join-Path $published 'marketplace.json' }
$marketplace = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$versions = @($marketplace.plugins.version | Select-Object -Unique)
if ($versions.Count -ne 1 -or $versions[0] -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$') {
    throw 'Published marketplace must have one valid release version.'
}
$current = $versions[0]
if ([version]$current -gt [version]$Version) { throw 'Downgrade publish blocked.' }
$tagExists = @(git -C $published tag --list $tag).Count -eq 1
if ($tagExists -and $current -ne $Version) { throw 'Existing tag conflicts with marketplace version.' }
$currentTag = "v$current"
$currentTagExists = @(git -C $published tag --list $currentTag).Count -eq 1
if (-not $currentTagExists -and -not ($ManualRepair -and $current -eq $Version)) {
    throw 'Current publication tag is missing; use authorized exact-release manual repair.'
}
$candidate = Join-Path ([IO.Path]::GetTempPath()) "excel-plugin-publication-$([Guid]::NewGuid().ToString('N'))"
$archive = "$candidate.zip"
$pathspecFile = "$candidate.paths"
New-Item -ItemType Directory -Path $candidate | Out-Null
try {
    git -C $published archive --format=zip "--output=$archive" HEAD
    Expand-Archive -LiteralPath $archive -DestinationPath $candidate
    # Remove only previously source-owned overlay files. Keep unrelated root files (e.g. LICENSE).
    $oldSourceTag = "refs/tags/$currentTag"
    git -C $source rev-parse --verify "$oldSourceTag^{commit}" | Out-Null
    $oldOverlay = @(git -C $source ls-tree -r --name-only $oldSourceTag -- .github/plugins/marketplace-repo)
    foreach ($oldPath in $oldOverlay) {
        $relative = $oldPath.Substring('.github/plugins/marketplace-repo/'.Length).Replace('/', [IO.Path]::DirectorySeparatorChar)
        $file = Join-Path $candidate $relative
        if (Test-Path -LiteralPath $file -PathType Leaf) { Remove-Item -LiteralPath $file }
    }
    & (Join-Path $source 'scripts\Sync-PublishedPluginRepo.ps1') `
        -PublishedRepoDir $candidate -BuiltPluginsDir $built -Version $Version
    git -C $candidate init --quiet
    # Compare the bytes Git will distribute, including its ordinary text clean conversion.
    git -C $candidate -c core.autocrlf=true add --all --force
    $tree = git -C $candidate write-tree
    $evidenceText = node (Join-Path $PSScriptRoot 'PluginContent.mjs') $published $candidate $tree $Version `
        $ManualRepair.IsPresent.ToString().ToLowerInvariant() $tagExists.ToString().ToLowerInvariant()
    $evidence = $evidenceText | ConvertFrom-Json
    if ($tagExists -and $evidence.changedPaths.Count -and -not $ManualRepair) {
        throw 'Existing immutable tag has different prepared content; automatic publication blocked.'
    }
    $actualTag = $currentTag
    $actualCommit = if ($currentTagExists) { git -C $published rev-parse "$currentTag^{commit}" } else { $baselineCommit }
    $status = $evidence.status
    if ($evidence.write -and -not $Preview) {
        # The destination has not been staged, committed, pushed or tagged before this decision.
        if ((git -C $published rev-parse HEAD) -ne $baselineCommit -or @(git -C $published status --porcelain).Count) {
            throw 'Published checkout changed while preparing output.'
        }
        $candidateFiles = @(git -C $candidate -c core.quotepath=false ls-tree -r --name-only $tree)
        $baselineFiles = @(git -C $published -c core.quotepath=false ls-tree -r --name-only HEAD)
        foreach ($file in $baselineFiles | Where-Object { $_ -cnotin $candidateFiles }) {
            Remove-Item -LiteralPath (Join-Path $published $file.Replace('/', [IO.Path]::DirectorySeparatorChar))
        }
        foreach ($file in $candidateFiles) {
            $relative = $file.Replace('/', [IO.Path]::DirectorySeparatorChar)
            $destination = Join-Path $published $relative
            New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $candidate $relative) -Destination $destination -Force
        }
        git -C $published config user.name 'ExcelMcp Publisher Bot'
        git -C $published config user.email 'excelmcp-publisher[bot]@users.noreply.github.com'
        $paths = @($baselineFiles + $candidateFiles | Sort-Object -Unique)
        [IO.File]::WriteAllText($pathspecFile, ($paths -join "`0") + "`0", [Text.UTF8Encoding]::new($false))
        git --literal-pathspecs -C $published -c core.autocrlf=true add --all --force "--pathspec-from-file=$pathspecFile" --pathspec-file-nul
        if ((git -C $published write-tree) -ne $tree) {
            throw 'Destination staged tree differs from the validated prepared publication; commit/push/tag blocked.'
        }
        if (@(git -C $published diff --cached --name-only).Count) {
            git -C $published commit -m "Release $tag" -m "Source release commit: $SourceCommit"
            git -C $published push origin HEAD:main
        }
        if (-not $tagExists) {
            git -C $published tag -a $tag -m "Release $tag"
            git -C $published push origin $tag
        }
        $actualTag = $tag
        $actualCommit = git -C $published rev-parse "$tag^{commit}"
    }
    $result = [ordered]@{
        status = if ($Preview) { 'preview' } else { $status }
        decision = $status
        published_tag = $actualTag
        published_commit = $actualCommit
        changed_plugins = @($evidence.changedPlugins)
        changed_paths = @($evidence.changedPaths)
        handoff = $evidence.handoff -and -not $tagExists -and -not $Preview
        baseline_commit = $baselineCommit
        baseline_fingerprint = $evidence.baselineFingerprint
        candidate_fingerprint = $evidence.candidateFingerprint
        repaired_stamps = @($evidence.repairs)
    }
    $result | ConvertTo-Json -Depth 10
    if ($env:GITHUB_OUTPUT) {
        "status=$($result.status)" >> $env:GITHUB_OUTPUT
        "published_tag=$actualTag" >> $env:GITHUB_OUTPUT
        "published_commit=$actualCommit" >> $env:GITHUB_OUTPUT
        "changed_plugins=$(ConvertTo-Json -InputObject @($result.changed_plugins) -Compress)" >> $env:GITHUB_OUTPUT
        "handoff=$($result.handoff.ToString().ToLowerInvariant())" >> $env:GITHUB_OUTPUT
    }
    if ($env:GITHUB_STEP_SUMMARY) {
        $message = if ($status -eq 'skipped') { 'Plugin publication skipped: no content changes.' } else { 'Plugin publication prepared/published.' }
        "$message Actual published tag: ``$actualTag``; commit: ``$actualCommit``. Changed plugins: $($result.changed_plugins -join ', ')." >> $env:GITHUB_STEP_SUMMARY
    }
}
finally {
    if (Test-Path -LiteralPath $archive) { Remove-Item -LiteralPath $archive }
    if (Test-Path -LiteralPath $pathspecFile) { Remove-Item -LiteralPath $pathspecFile }
    Remove-Item -LiteralPath $candidate -Recurse -Force
}
