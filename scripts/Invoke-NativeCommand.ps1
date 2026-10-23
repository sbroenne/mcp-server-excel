#Requires -Version 7.4

function Invoke-NativeCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Executable,
        [string[]]$Arguments = @(),
        [ValidateRange(1, 28800)][int]$DeadlineSeconds = 1200,
        [hashtable]$Environment = @{},
        [string]$OutputPath,
        [string]$ErrorPath,
        [string]$ProgressName,
        [switch]$AllowFailure
    )
    if (-not $IsWindows) { throw 'Bounded build and test execution requires Windows.' }
    if (-not ('ExcelMcpValidation.JobApi' -as [type])) {
        # Assign a suspended process before it can create descendants.
        Add-Type -Namespace ExcelMcpValidation -Name JobApi -MemberDefinition @'
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public struct StartupInfo {
    public int Size;
    public string Reserved, Desktop, Title;
    public int X, Y, XSize, YSize, XChars, YChars, Fill, Flags;
    public short ShowWindow, ReservedLength;
    public System.IntPtr ReservedBytes, Input, Output, Error;
}
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct ProcessInfo {
    public System.IntPtr Process, Thread;
    public uint Id, ThreadId;
}
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern System.IntPtr CreateJobObject(System.IntPtr attributes, string name);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool SetInformationJobObject(System.IntPtr job, int kind, System.IntPtr information, uint length);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool AssignProcessToJobObject(System.IntPtr job, System.IntPtr process);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool QueryInformationJobObject(System.IntPtr job, int kind, System.IntPtr information, uint length, System.IntPtr returned);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool TerminateJobObject(System.IntPtr job, uint exitCode);
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode, SetLastError=true)]
public static extern bool CreateProcess(string application, System.Text.StringBuilder command, System.IntPtr processAttributes,
    System.IntPtr threadAttributes, bool inheritHandles, uint flags, System.IntPtr environment, string directory,
    ref StartupInfo startup, out ProcessInfo process);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern uint ResumeThread(System.IntPtr thread);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool TerminateProcess(System.IntPtr process, uint exitCode);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool GetExitCodeProcess(System.IntPtr process, out uint exitCode);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool CloseHandle(System.IntPtr handle);
public static async System.Threading.Tasks.Task<string> CaptureOutput(
    System.IO.StreamReader reader, System.IO.StreamWriter log) {
    var output = new System.Text.StringBuilder();
    var buffer = new char[4096];
    int count;
    while ((count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) != 0) {
        output.Append(buffer, 0, count);
        if (log != null) {
            await log.WriteAsync(buffer, 0, count).ConfigureAwait(false);
        }
    }
    return output.ToString();
}
'@
    }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $process = $null
    $job = [IntPtr]::Zero
    $information = [IntPtr]::Zero
    $environmentBlock = [IntPtr]::Zero
    $native = [ExcelMcpValidation.JobApi+ProcessInfo]::new()
    $pipes = [Collections.Generic.List[IO.Pipes.AnonymousPipeServerStream]]::new()
    $readers = [Collections.Generic.List[IO.StreamReader]]::new()
    $logs = [Collections.Generic.List[IO.StreamWriter]]::new()
    $primary = $null
    $stdout = $null
    $stderr = $null
    try {
        foreach ($path in @($OutputPath, $ErrorPath)) {
            if ($path) {
                $log = [IO.StreamWriter]::new($path, $false, [Text.UTF8Encoding]::new($false))
                $logs.Add($log)
                $log.AutoFlush = $true
            } else { $logs.Add($null) }
        }
        $job = [ExcelMcpValidation.JobApi]::CreateJobObject([IntPtr]::Zero, $null)
        if ($job -eq [IntPtr]::Zero) { throw [ComponentModel.Win32Exception]::new() }
        $length = if ([IntPtr]::Size -eq 8) { 144 } else { 112 }
        $information = [Runtime.InteropServices.Marshal]::AllocHGlobal($length)
        [Runtime.InteropServices.Marshal]::Copy([byte[]]::new($length), 0, $information, $length)
        [Runtime.InteropServices.Marshal]::WriteInt32($information, 16, 0x2000)
        if (-not [ExcelMcpValidation.JobApi]::SetInformationJobObject($job, 9, $information, $length)) {
            throw [ComponentModel.Win32Exception]::new()
        }
        $application = (Get-Command $Executable -CommandType Application | Select-Object -First 1).Source
        if (-not $application) { throw "Cannot find native executable: $Executable." }
        $info = [Diagnostics.ProcessStartInfo]::new($application)
        foreach ($key in @($info.Environment.Keys)) {
            if ($key.StartsWith('GIT_', [StringComparison]::OrdinalIgnoreCase)) { $null = $info.Environment.Remove($key) }
        }
        foreach ($key in $Environment.Keys) { $info.Environment[$key] = [string]$Environment[$key] }
        if ([IO.Path]::GetFileNameWithoutExtension($application) -eq 'pwsh') {
            $application = Join-Path $PSHOME pwsh.exe
        }
        $environmentText = (($info.Environment.Keys | Sort-Object | ForEach-Object { "$_=$($info.Environment[$_])" }) -join "`0") + "`0`0"
        $environmentBlock = [Runtime.InteropServices.Marshal]::StringToHGlobalUni($environmentText)
        $command = @($application) + @($Arguments) | ForEach-Object {
            '"' + [regex]::Replace([regex]::Replace($_, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
        }
        foreach ($direction in @([IO.Pipes.PipeDirection]::Out, [IO.Pipes.PipeDirection]::In, [IO.Pipes.PipeDirection]::In)) {
            $pipes.Add([IO.Pipes.AnonymousPipeServerStream]::new($direction, [IO.HandleInheritability]::Inheritable))
        }
        $startup = [ExcelMcpValidation.JobApi+StartupInfo]::new()
        $startup.Size = [Runtime.InteropServices.Marshal]::SizeOf($startup)
        $startup.Flags = 0x100
        $startup.Input = $pipes[0].ClientSafePipeHandle.DangerousGetHandle()
        $startup.Output = $pipes[1].ClientSafePipeHandle.DangerousGetHandle()
        $startup.Error = $pipes[2].ClientSafePipeHandle.DangerousGetHandle()
        if (-not [ExcelMcpValidation.JobApi]::CreateProcess($application, [Text.StringBuilder]::new(($command -join ' ')),
                [IntPtr]::Zero, [IntPtr]::Zero, $true, 0x08000404, $environmentBlock, (Get-Location).ProviderPath,
                [ref]$startup, [ref]$native)) { throw [ComponentModel.Win32Exception]::new() }
        if (-not [ExcelMcpValidation.JobApi]::AssignProcessToJobObject($job, $native.Process)) {
            throw [ComponentModel.Win32Exception]::new()
        }
        $process = [Diagnostics.Process]::GetProcessById($native.Id)
        $null = $process.SafeHandle
        foreach ($pipe in $pipes) { $pipe.DisposeLocalCopyOfClientHandle() }
        $pipes[0].Dispose()
        foreach ($pipe in @($pipes[1], $pipes[2])) { $readers.Add([IO.StreamReader]::new($pipe, [Text.Encoding]::UTF8)) }
        $stdout = [ExcelMcpValidation.JobApi]::CaptureOutput($readers[0], $logs[0])
        $stderr = [ExcelMcpValidation.JobApi]::CaptureOutput($readers[1], $logs[1])
        if ([ExcelMcpValidation.JobApi]::ResumeThread($native.Thread) -eq [uint32]::MaxValue) {
            throw [ComponentModel.Win32Exception]::new()
        }
        $nextProgress = 0
        do {
            if ($stdout.IsFaulted) { $null = $stdout.GetAwaiter().GetResult() }
            if ($stderr.IsFaulted) { $null = $stderr.GetAwaiter().GetResult() }
            if (-not [ExcelMcpValidation.JobApi]::QueryInformationJobObject($job, 1, $information, 48, [IntPtr]::Zero)) {
                throw [ComponentModel.Win32Exception]::new()
            }
            $active = [Runtime.InteropServices.Marshal]::ReadInt32($information, 40)
            if ($stdout.IsCompleted -and $stderr.IsCompleted -and $active -eq 0) { break }
            if ($timer.Elapsed.TotalSeconds -ge $DeadlineSeconds) {
                throw [TimeoutException]::new("$Executable exceeded the $DeadlineSeconds-second hard deadline.")
            }
            if ($ProgressName -and $timer.Elapsed.TotalSeconds -ge $nextProgress) {
                Write-Host "${ProgressName}: running; elapsed $($timer.Elapsed.ToString('hh\:mm\:ss'))."
                $nextProgress = $timer.Elapsed.TotalSeconds + 30
            }
            Start-Sleep -Milliseconds 50
        } while ($true)
        $exitCode = [uint32]0
        if (-not [ExcelMcpValidation.JobApi]::GetExitCodeProcess($native.Process, [ref]$exitCode)) {
            throw [ComponentModel.Win32Exception]::new()
        }
        $result = @{ ExitCode = $exitCode; Output = $stdout.GetAwaiter().GetResult(); Error = $stderr.GetAwaiter().GetResult() }
        if (-not $AllowFailure -and $result.ExitCode -ne 0) {
            throw "$Executable $($Arguments[0]) failed with exit code $($result.ExitCode).`n$($result.Output)`n$($result.Error)"
        }
        return $result
    } catch { $primary = $_.Exception; throw }
    finally {
        $cleanup = [Collections.Generic.List[Exception]]::new()
        if ($process -and -not $process.HasExited) {
            # Packaged PowerShell activations can leave a job while remaining descendants.
            try { $process.Kill($true) } catch { $cleanup.Add($_.Exception) }
        } elseif (-not $process -and $native.Process -ne [IntPtr]::Zero) {
            if (-not [ExcelMcpValidation.JobApi]::TerminateProcess($native.Process, 1)) { $cleanup.Add([ComponentModel.Win32Exception]::new()) }
        }
        if ($job -ne [IntPtr]::Zero) {
            if (-not [ExcelMcpValidation.JobApi]::TerminateJobObject($job, 1)) { $cleanup.Add([ComponentModel.Win32Exception]::new()) }
            $stop = [Diagnostics.Stopwatch]::StartNew()
            while ($information -ne [IntPtr]::Zero) {
                if (-not [ExcelMcpValidation.JobApi]::QueryInformationJobObject($job, 1, $information, 48, [IntPtr]::Zero)) {
                    $cleanup.Add([ComponentModel.Win32Exception]::new()); break
                }
                if ([Runtime.InteropServices.Marshal]::ReadInt32($information, 40) -eq 0) { break }
                if ($stop.Elapsed.TotalSeconds -ge 5) { $cleanup.Add([TimeoutException]::new('Owned descendants did not exit during cleanup.')); break }
                Start-Sleep -Milliseconds 25
            }
        }
        foreach ($handle in @($native.Thread, $native.Process, $job)) {
            if ($handle -ne [IntPtr]::Zero -and -not [ExcelMcpValidation.JobApi]::CloseHandle($handle)) {
                $cleanup.Add([ComponentModel.Win32Exception]::new())
            }
            if ($process) { $process.Dispose() }
        }
        foreach ($pointer in @($information, $environmentBlock)) {
            if ($pointer -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($pointer) }
        }
        $timeoutOutput = ''
        if ($stdout -and $stderr) {
            $capture = [Threading.Tasks.Task]::WhenAll([Threading.Tasks.Task[]]@($stdout, $stderr))
            try {
                if ($capture.Wait(5000)) {
                    if ($primary -is [TimeoutException]) {
                        $timeoutOutput = "$($stdout.GetAwaiter().GetResult())`n$($stderr.GetAwaiter().GetResult())"
                    }
                }
                else { $cleanup.Add([TimeoutException]::new('Owned command output did not close.')) }
            } catch { $cleanup.Add($_.Exception) }
        }
        foreach ($reader in $readers) { $reader.Dispose() }
        foreach ($pipe in $pipes) { $pipe.Dispose() }
        foreach ($log in $logs) {
            if ($log) { try { $log.Dispose() } catch { $cleanup.Add($_.Exception) } }
        }
        if ($cleanup.Count) {
            if ($primary) { $cleanup.Insert(0, $primary) }
            throw [AggregateException]::new('Command and owned-process cleanup failed.', $cleanup)
        }
        if ($primary -is [TimeoutException]) { throw [TimeoutException]::new("$($primary.Message)`n$timeoutOutput", $primary) }
    }
}
