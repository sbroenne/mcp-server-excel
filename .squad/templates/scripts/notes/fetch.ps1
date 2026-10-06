#!/usr/bin/env pwsh
# Fetch into remote-tracking refs; never overwrite local notes.
# -Setup migrates the legacy refspec. -Merge reconciles divergent notes.
[CmdletBinding()]
param(
    [string]$Remote = "origin",
    [string]$RepoPath = ".",
    [switch]$Setup,
    [switch]$Merge,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $false
$repo = (Resolve-Path $RepoPath).Path
$refspec = "+refs/notes/squad/*:refs/notes/remotes/$Remote/squad/*"

function Invoke-Git {
    param([string[]]$GitArgs)
    $output = & git -C $repo @GitArgs 2>&1
    if ($LASTEXITCODE -ne 0) { throw "git $GitArgs failed: $output" }
    return $output
}

if ($Setup) {
    $key = "remote.$Remote.fetch"
    $existing = & git -C $repo config --get-all $key
    if ($LASTEXITCODE -notin 0, 1) { throw "Could not read $key" }
    foreach ($legacy in @("refs/notes/*:refs/notes/*", "+refs/notes/*:refs/notes/*")) {
        if ($existing -contains $legacy) {
            Invoke-Git -GitArgs @("config", "--fixed-value", "--unset-all", $key, $legacy) | Out-Null
        }
    }
    if ($existing -notcontains $refspec) {
        Invoke-Git -GitArgs @("config", "--add", $key, $refspec) | Out-Null
    }
}

Invoke-Git -GitArgs @("fetch", $Remote, $refspec) | Out-Null
$prefix = "refs/notes/remotes/$Remote/"
$refs = Invoke-Git -GitArgs @("for-each-ref", "${prefix}squad/", "--format=%(refname)")
foreach ($remoteRef in $refs) {
    $namespace = $remoteRef.Substring($prefix.Length)
    $localRef = "refs/notes/$namespace"
    & git -C $repo show-ref --verify --quiet $localRef
    $exists = $LASTEXITCODE
    if ($exists -eq 1) {
        $sha = Invoke-Git -GitArgs @("rev-parse", $remoteRef)
        Invoke-Git -GitArgs @("update-ref", $localRef, $sha, "") | Out-Null
    } elseif ($exists -ne 0) {
        throw "Could not check local notes ref $localRef"
    } elseif ($Merge) {
        Invoke-Git -GitArgs @("notes", "--ref=$namespace", "merge", "-s", "cat_sort_uniq", $remoteRef) | Out-Null
    }
}

if (-not $Quiet) {
    Write-Host "[notes/fetch] Notes fetched into remote-tracking refs."
    if ($Merge) { Write-Host "[notes/fetch] Notes merge complete." }
}
