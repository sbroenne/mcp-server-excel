param(
    [string]$SkillsDirectory = (Join-Path $PSScriptRoot '..\.github\skills')
)

$ErrorActionPreference = 'Stop'
$repository = 'DietrichGebert/ponytail'
$skillDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    (Join-Path $SkillsDirectory 'ponytail-review'))
if (Test-Path -LiteralPath $skillDirectory) {
    Remove-Item -LiteralPath $skillDirectory -Recurse -Force
}

$installed = $false
try {
    $tag = & gh api "repos/$repository/releases/latest" --jq .tag_name
    if ($LASTEXITCODE -ne 0) {
        throw "Could not resolve the latest Ponytail release (exit code $LASTEXITCODE)."
    }
    if ([string]::IsNullOrWhiteSpace($tag) -or $tag -eq 'null') {
        throw 'The latest Ponytail release has no tag.'
    }

    & gh skill install $repository "ponytail-review@$tag" --dir $SkillsDirectory --force
    if ($LASTEXITCODE -ne 0) {
        throw "Ponytail review skill installation failed (exit code $LASTEXITCODE). GitHub CLI 2.90.0 or later is required."
    }
    if (-not (Test-Path -LiteralPath (Join-Path $skillDirectory 'SKILL.md') -PathType Leaf)) {
        throw 'Ponytail review skill installation did not create SKILL.md.'
    }
    Write-Output "Installed ponytail-review release $tag."
    $installed = $true
}
finally {
    if (-not $installed -and (Test-Path -LiteralPath $skillDirectory)) {
        Remove-Item -LiteralPath $skillDirectory -Recurse -Force
    }
}
