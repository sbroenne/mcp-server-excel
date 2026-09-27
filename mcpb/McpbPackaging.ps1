function Remove-McpbStagingDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [TimeSpan]$Timeout = [TimeSpan]::FromMinutes(2),

        [TimeSpan]$RetryInterval = [TimeSpan]::FromMilliseconds(500),

        [scriptblock]$RemoveDirectory = {
            param([string]$TargetPath)
            [System.IO.Directory]::Delete($TargetPath, $true)
        }
    )

    if ($Timeout -lt [TimeSpan]::Zero) {
        throw [System.ArgumentOutOfRangeException]::new(
            "Timeout",
            $Timeout,
            "The staging cleanup timeout cannot be negative.")
    }

    if ($RetryInterval -lt [TimeSpan]::Zero) {
        throw [System.ArgumentOutOfRangeException]::new(
            "RetryInterval",
            $RetryInterval,
            "The staging cleanup retry interval cannot be negative.")
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $attempts = 0
    $lastFailure = $null

    while (Test-Path -LiteralPath $Path -PathType Container) {
        $attempts++
        try {
            $lastFailure = $null
            & $RemoveDirectory $Path
        }
        catch [System.UnauthorizedAccessException] {
            $lastFailure = $_.Exception
        }
        catch [System.IO.IOException] {
            $lastFailure = $_.Exception
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            return
        }

        $remaining = $Timeout - $stopwatch.Elapsed
        if ($remaining -le [TimeSpan]::Zero) {
            $attemptLabel = if ($attempts -eq 1) { "attempt" } else { "attempts" }
            $timeoutMilliseconds = [Math]::Round($Timeout.TotalMilliseconds)
            $lastError = if ($lastFailure) { $lastFailure.Message } else { "the directory still exists" }
            throw [System.IO.IOException]::new(
                "Failed to remove MCPB staging directory '$Path' after $attempts $attemptLabel within $timeoutMilliseconds ms. " +
                "The verified bundle was preserved, but stale staging remains. Last error: $lastError",
                $lastFailure)
        }

        $delay = if ($RetryInterval -lt $remaining) { $RetryInterval } else { $remaining }
        if ($delay -gt [TimeSpan]::Zero) {
            Start-Sleep -Milliseconds ([Math]::Ceiling($delay.TotalMilliseconds))
        }
    }
}

function New-McpbArchive {
    param(
        [Parameter(Mandatory)]
        [string]$SourceDirectory,

        [Parameter(Mandatory)]
        [string]$DestinationPath,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$MacExecutableRelativePath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $DestinationPath) {
        Remove-Item -LiteralPath $DestinationPath -Force
    }

    $archive = [System.IO.Compression.ZipFile]::Open(
        $DestinationPath,
        [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File) {
            $relativePath = [System.IO.Path]::GetRelativePath($SourceDirectory, $file.FullName).Replace('\', '/')
            $entry = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $file.FullName,
                $relativePath,
                [System.IO.Compression.CompressionLevel]::Optimal)

            if ($relativePath -eq $MacExecutableRelativePath) {
                $unixExecutableMode = [BitConverter]::ToInt32(
                    [BitConverter]::GetBytes([Convert]::ToUInt32("81ED0000", 16)),
                    0)
                $entry.ExternalAttributes = $unixExecutableMode
            }
        }
    }
    finally {
        $archive.Dispose()
    }
}
