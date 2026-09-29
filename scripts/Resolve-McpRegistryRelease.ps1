[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Tag,

    [string]$ExpectedCommit,

    [string]$ExpectedVersion
)

$ErrorActionPreference = 'Stop'

if ($Tag -notmatch '^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw 'Invalid release tag.'
}

git fetch origin main --no-tags
if ($LASTEXITCODE -ne 0) { throw 'Failed to fetch protected main.' }

$commit = (git rev-parse "refs/tags/$Tag^{commit}")
if ($LASTEXITCODE -ne 0 -or -not $commit) { throw "Release tag '$Tag' was not found." }
$commit = $commit.Trim()

git merge-base --is-ancestor $commit refs/remotes/origin/main
if ($LASTEXITCODE -ne 0) {
    throw "Release tag '$Tag' resolves to a commit that is not reachable from protected main."
}

$version = $Tag.Substring(1)
if ($ExpectedCommit -and $commit -ne $ExpectedCommit) {
    throw 'Release commit mismatch.'
}
if ($ExpectedVersion -and $version -ne $ExpectedVersion) {
    throw 'Release version mismatch.'
}

$release = gh release view $Tag --json isDraft,tagName | ConvertFrom-Json
if ($LASTEXITCODE -ne 0 -or $release.isDraft -or $release.tagName -ne $Tag) {
    throw "Published GitHub release '$Tag' was not found."
}

"commit=$commit" >> $env:GITHUB_OUTPUT
"version=$version" >> $env:GITHUB_OUTPUT
"tag=$Tag" >> $env:GITHUB_OUTPUT
