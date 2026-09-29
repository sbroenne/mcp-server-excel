[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Windows PowerShell rebuilds a command line when it invokes a native executable, and its
# built-in quoting drops embedded double quotes. That silently corrupts JSON arguments such as
# --values '[["Name","Amount"]]'. Build the command line using the standard MSVCRT quoting rules
# and hand it directly to Node's npx entry point.
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

if ($null -eq $PassthroughArgs) {
    $PassthroughArgs = @()
}

$npxCommand = Get-Command "npx.cmd" -CommandType Application -ErrorAction SilentlyContinue |
    Select-Object -First 1
$nodeCommand = Get-Command "node.exe" -CommandType Application -ErrorAction SilentlyContinue |
    Select-Object -First 1
$npxCliPath = if ($null -ne $npxCommand) {
    Join-Path (Split-Path -Parent $npxCommand.Source) "node_modules\npm\bin\npx-cli.js"
}

if ($null -ne $nodeCommand -and -not [string]::IsNullOrWhiteSpace($npxCliPath) -and
    (Test-Path -LiteralPath $npxCliPath -PathType Leaf)) {
    $binaryPath = $nodeCommand.Source
    $nativeArguments = @($npxCliPath, "-y", "@sbroenne/excelcli@latest") + @($PassthroughArgs)
} else {
    throw "excel-cli requires Node.js 18 or later with npm/npx available on PATH."
}

$startInfo = New-Object System.Diagnostics.ProcessStartInfo
$startInfo.FileName = $binaryPath
$startInfo.Arguments = (($nativeArguments | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' ')
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
