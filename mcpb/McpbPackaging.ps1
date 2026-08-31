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
