#!/usr/bin/env pwsh
# scripts/notes/write-note.ps1
# ─────────────────────────────────────────────────────────────────────────────
# Helper for agents to write notes without wrestling with JSON escaping.
# Validates namespace ownership, handles conflicts, pushes automatically.
#
# Usage:
#   ./scripts/notes/write-note.ps1 -Agent data -Type decision \
#       -Content '{"decision":"Use JWT","reasoning":"..."}' \
#       [-Commit HEAD] [-Promote] [-Archive]
# ─────────────────────────────────────────────────────────────────────────────

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Agent,

    [Parameter(Mandatory)]
    [ValidateSet("decision","research","review","security-review","progress",
                 "api-contract","risk-assessment","routing-discovery","counter-argument")]
    [string]$Type,

    [Parameter(Mandatory)]
    [string]$Content,   # JSON object with type-specific fields

    [string]$Commit     = "HEAD",
    [string]$RepoPath   = ".",
    [string]$Remote     = "origin",
    [switch]$Promote,   # set promote_to_permanent: true
    [switch]$Archive,   # set archive_on_close: true
    [switch]$NoPush,    # skip auto-push
    [switch]$Quiet
)

function Log ([string]$msg, [string]$color = "White") {
    if (-not $Quiet) { Write-Host "[notes/write] $msg" -ForegroundColor $color }
}

$repo      = Resolve-Path $RepoPath
$namespace = "squad/$($Agent.ToLower())"

# ── Validate JSON content ────────────────────────────────────────────────────
try {
    $parsed = $Content | ConvertFrom-Json -ErrorAction Stop
} catch {
    Write-Error "Content must be valid JSON. Got: $Content"
    exit 1
}

# ── Build full note object ────────────────────────────────────────────────────
$note = [ordered]@{
    agent     = (Get-Culture).TextInfo.ToTitleCase($Agent.ToLower())
    timestamp = [System.DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    type      = $Type
}

# Merge content fields into note
$parsed.PSObject.Properties | ForEach-Object { $note[$_.Name] = $_.Value }

# Add flag fields
if ($Promote)  { $note["promote_to_permanent"] = $true }
if ($Archive)  { $note["archive_on_close"]     = $true }

$noteJson = $note | ConvertTo-Json -Compress -Depth 10

# ── Fetch first to avoid conflicts ───────────────────────────────────────────
Log "Fetching notes before write..."
& (Join-Path $PSScriptRoot "fetch.ps1") -Remote $Remote -RepoPath $repo -Merge -Quiet

# ── Check if note already exists on this commit ─────────────────────────────
$existing = git -C $repo notes --ref=$namespace show $Commit 2>&1
$useAppend = ($LASTEXITCODE -eq 0)

if ($useAppend) {
    Log "Note exists on $Commit — appending" DarkYellow
    git -C $repo notes --ref=$namespace append -m $noteJson $Commit
} else {
    git -C $repo notes --ref=$namespace add -m $noteJson $Commit
}

if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to write note to refs/notes/$namespace on $Commit"
    exit 1
}

Log "Note written to refs/notes/$namespace on $($Commit.Substring(0,[Math]::Min(8,$Commit.Length)))" Green

# ── Push with retry ──────────────────────────────────────────────────────────
if (-not $NoPush) {
    $maxRetries = 5
    $nsRef = "refs/notes/$namespace"

    for ($i = 0; $i -lt $maxRetries; $i++) {
        Log "Pushing notes (attempt $($i+1))..."
        $pushOut = git -C $repo push $Remote "${nsRef}:${nsRef}" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Log "Notes pushed successfully." Green
            break
        }

        if ($pushOut -match "non-fast-forward|fetch first|rejected") {
            Log "Push conflict — fetch-first retry..." DarkYellow

            & (Join-Path $PSScriptRoot "fetch.ps1") -Remote $Remote -RepoPath $repo -Merge -Quiet
            if ($i -eq $maxRetries - 1) {
                throw "Notes push still conflicts after $maxRetries attempts. Local notes are preserved in $nsRef."
            }

            $jitter = Get-Random -Minimum 0 -Maximum 1000
            $sleep  = [Math]::Pow(2, $i) + $jitter / 1000
            Start-Sleep -Seconds $sleep

        } else {
            throw "Notes push failed: $pushOut. Local notes are preserved in $nsRef."
        }
    }
}

# ── Show result ───────────────────────────────────────────────────────────────
if (-not $Quiet) {
    Log "Note content:"
    $note | ConvertTo-Json -Depth 5 | Write-Host -ForegroundColor DarkGray
}
