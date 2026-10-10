function Remove-McpbStagingDirectory {
    param(
        [Parameter(Mandatory)][string]$Path,
        [TimeSpan]$Timeout = [TimeSpan]::FromMinutes(2),
        [TimeSpan]$RetryInterval = [TimeSpan]::FromMilliseconds(500)
    )
    $root = Split-Path $PSScriptRoot -Parent
    . (Join-Path $root 'scripts\PackageHelpers.ps1')
    Invoke-TypedPackage -SourceRoot $root -Options @{
        Operation = 'RemoveStaging'; Source = $Path
        TimeoutMilliseconds = $Timeout.TotalMilliseconds
        RetryMilliseconds = $RetryInterval.TotalMilliseconds
    }
}
