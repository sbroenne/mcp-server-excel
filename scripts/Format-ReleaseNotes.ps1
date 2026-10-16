#Requires -Version 7.4
<#
.SYNOPSIS
    Makes the changesets-generated release-notes body easier to read.

.DESCRIPTION
    Rewrites the release-notes body in place. CHANGELOG.md is not touched.
      - Renames changesets section headings to plain headings that nest under
        the template's "What's New" heading.
      - Puts each entry's description first, followed by its pull request link.
      - Drops commit-hash links.
      - Drops the "Thanks" line for the maintainer and keeps it for other contributors.

.PARAMETER Path
    Release-notes body file to rewrite.

.PARAMETER Maintainer
    GitHub login whose "Thanks" line is dropped.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [string]$Maintainer = 'sbroenne'
)

$ErrorActionPreference = 'Stop'

$headings = @{
    'Major Changes' = 'Breaking changes'
    'Minor Changes' = 'New features'
    'Patch Changes' = 'Fixes and improvements'
}

$authorPattern = '\[@(?<login>[^\]]+)\]\([^)]+\)'
$entryPattern = '^- (?:(?<pr>\[#\d+\]\([^)]+\)) )?(?:\[`[0-9a-f]+`\]\([^)]+\) )?(?:Thanks (?<thanks>' + $authorPattern + '(?:, ' + $authorPattern + ')*)! )?- (?<text>.*)$'

$lines = (Get-Content -LiteralPath $Path -Raw) -split "`r?`n"
$output = foreach ($line in $lines) {
    if ($line -match '^###\s+(?<name>.+?)\s*$' -and $headings.ContainsKey($Matches.name)) {
        "#### $($headings[$Matches.name])"
        continue
    }

    if ($line -match $entryPattern) {
        $entry = "- $($Matches.text)"
        if ($Matches.pr) {
            $entry += " ($($Matches.pr))"
        }
        if ($Matches.thanks) {
            $credits = @([regex]::Matches($Matches.thanks, $authorPattern) |
                Where-Object { $_.Groups['login'].Value -ne $Maintainer } |
                ForEach-Object { $_.Value })
            if ($credits.Count -gt 0) {
                $entry += " Thanks $($credits -join ', ')!"
            }
        }
        $entry
        continue
    }

    $line
}

Set-Content -LiteralPath $Path -Value (($output -join "`n").Trim()) -NoNewline
