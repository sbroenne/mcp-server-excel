#Requires -Version 7.0
Set-StrictMode -Version Latest

function Assert-MacAutomationAllowed {
    $info = [Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'pwsh'))
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in @('-NoProfile', '-File',
        (Join-Path $PSScriptRoot 'Test-MacAutomationBoundary.ps1'), '-ManagedProbe')) {
        $info.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw 'Could not start the non-prompting Automation check.' }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(15000)) {
            $process.Kill($true)
            $process.WaitForExit()
            throw 'Automation preflight timed out. No Excel command was dispatched.'
        }
        $output = $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($process.ExitCode -notin @(0, 2)) { throw "Automation preflight failed: $errorText" }
        $result = ConvertFrom-Json $output
        if ($result.status -ne 'Allowed') {
            throw "Mac test prerequisite: $($result.status) (OSStatus $($result.osStatus)). Open licensed Excel and grant Automation consent interactively before testing; tests will not request consent."
        }
    }
    finally { $process.Dispose() }
}
