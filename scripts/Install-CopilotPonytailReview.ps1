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
$temporaryDirectory = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcp.Ponytail.$([guid]::NewGuid().ToString('N'))"
try {
    $headers = @{
        Accept = 'application/vnd.github+json'
        'User-Agent' = 'ExcelMcp-Ponytail-Setup'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    if ($env:GH_TOKEN) { $headers.Authorization = "Bearer $env:GH_TOKEN" }
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$repository/releases/latest" `
        -Headers $headers -TimeoutSec 60
    $tag = $release.tag_name
    if ([string]::IsNullOrWhiteSpace($tag) -or $tag -eq 'null') {
        throw 'The latest Ponytail release has no tag.'
    }
    $commit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repository/commits/$([uri]::EscapeDataString($tag))" `
        -Headers $headers -TimeoutSec 60
    if ($commit.sha -notmatch '^[a-f0-9]{40}$') {
        throw 'The latest Ponytail release has no valid commit SHA.'
    }

    New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
    $archive = Join-Path $temporaryDirectory 'release.zip'
    Invoke-WebRequest -Uri "https://api.github.com/repos/$repository/zipball/$($commit.sha)" `
        -Headers $headers -OutFile $archive -TimeoutSec 180
    $extracted = Join-Path $temporaryDirectory 'extracted'
    Expand-Archive -LiteralPath $archive -DestinationPath $extracted
    $roots = @(Get-ChildItem -LiteralPath $extracted -Directory)
    if ($roots.Count -ne 1) {
        throw 'Unexpected Ponytail release archive layout.'
    }
    $sourceSkill = Join-Path $roots[0].FullName 'skills\ponytail-review'
    if (-not (Test-Path -LiteralPath (Join-Path $sourceSkill 'SKILL.md') -PathType Leaf)) {
        throw 'The Ponytail release archive has no review SKILL.md.'
    }

    New-Item -ItemType Directory -Path (Split-Path -Parent $skillDirectory) -Force | Out-Null
    Copy-Item -LiteralPath $sourceSkill -Destination $skillDirectory -Recurse
    Copy-Item -LiteralPath (Join-Path $roots[0].FullName 'LICENSE') -Destination $skillDirectory
    Write-Output "Installed ponytail-review release $tag."
    Write-Output "Ponytail source: $repository@$($commit.sha)"
    $installed = $true
}
finally {
    if (-not $installed -and (Test-Path -LiteralPath $skillDirectory)) {
        Remove-Item -LiteralPath $skillDirectory -Recurse -Force
    }
    if (Test-Path -LiteralPath $temporaryDirectory) {
        Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force
    }
}
