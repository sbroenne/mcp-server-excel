function Invoke-ExcelMcpBuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Arguments,
        [string]$OptionsParameter,
        [hashtable]$Options
    )
    $root = if ($env:EXCELMCP_BUILD_ROOT) { $env:EXCELMCP_BUILD_ROOT } else { Split-Path -Parent $PSScriptRoot }
    $optionsFile = $null
    try {
        if ($OptionsParameter) {
            $optionsFile = Join-Path ([IO.Path]::GetTempPath()) "ExcelMcpBuildOptions-$([Guid]::NewGuid().ToString('N')).json"
            $typed = @{}
            foreach ($name in $Options.Keys) {
                if ($Options[$name] -is [string] -and $Options[$name].Length -eq 0) { continue }
                $typed[$name] = $Options[$name]
            }
            [IO.File]::WriteAllText($optionsFile, ($typed | ConvertTo-Json -Depth 30), [Text.UTF8Encoding]::new($false))
            $Arguments += @($OptionsParameter, $optionsFile)
        }
        & (Join-Path $root 'build.ps1') @Arguments --root $root
        if ($LASTEXITCODE -ne 0) { throw "Build operation '$($Arguments[0])' failed with exit code $LASTEXITCODE." }
        $global:LASTEXITCODE = 0
    }
    finally {
        if ($optionsFile -and (Test-Path -LiteralPath $optionsFile)) { Remove-Item -LiteralPath $optionsFile -Force }
    }
}
