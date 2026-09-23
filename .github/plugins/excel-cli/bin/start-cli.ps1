[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Windows PowerShell rebuilds a command line when it invokes a native executable, and its
# built-in quoting drops embedded double quotes. That silently corrupts JSON arguments such as
# --values '[["Name","Amount"]]', which is the most common excelcli invocation. Build the command
# line ourselves using the standard MSVCRT quoting rules and hand it to the process verbatim.
function ConvertTo-NativeArgument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    if ($Value.Length -gt 0 -and $Value -notmatch '[ \t\n\v"]') {
        return $Value
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')

    $index = 0
    while ($index -lt $Value.Length) {
        $backslashes = 0
        while ($index -lt $Value.Length -and $Value[$index] -eq '\') {
            $index++
            $backslashes++
        }

        if ($index -eq $Value.Length) {
            # Trailing backslashes must be doubled so they do not escape the closing quote.
            [void]$builder.Append('\' * ($backslashes * 2))
            break
        }

        if ($Value[$index] -eq '"') {
            # Escape the quote and double the backslashes that precede it.
            [void]$builder.Append('\' * ($backslashes * 2 + 1))
            [void]$builder.Append('"')
        } else {
            [void]$builder.Append('\' * $backslashes)
            [void]$builder.Append($Value[$index])
        }

        $index++
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

$downloadScript = Join-Path $PSScriptRoot "download.ps1"
$binaryPath = & $downloadScript -PassThru -Quiet

if ([string]::IsNullOrWhiteSpace($binaryPath) -or -not (Test-Path $binaryPath)) {
    throw "excel-cli bootstrap did not resolve a usable excelcli.exe runtime."
}

if ($null -eq $PassthroughArgs) {
    $PassthroughArgs = @()
}

$startInfo = New-Object System.Diagnostics.ProcessStartInfo
$startInfo.FileName = $binaryPath
$startInfo.Arguments = (($PassthroughArgs | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' ')
$startInfo.UseShellExecute = $false
$startInfo.RedirectStandardOutput = $true
$startInfo.RedirectStandardError = $true

$process = [System.Diagnostics.Process]::Start($startInfo)
$stdoutTask = $process.StandardOutput.ReadToEndAsync()
$stderrTask = $process.StandardError.ReadToEndAsync()
$process.WaitForExit()

$stdout = $stdoutTask.GetAwaiter().GetResult()
$stderr = $stderrTask.GetAwaiter().GetResult()

if (-not [string]::IsNullOrEmpty($stderr)) {
    [Console]::Error.Write($stderr)
}

if (-not [string]::IsNullOrEmpty($stdout)) {
    Write-Output -NoEnumerate $stdout
}

exit $process.ExitCode
