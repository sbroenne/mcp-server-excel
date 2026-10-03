. (Join-Path $PSScriptRoot 'AzureRunnerHost.ps1')
. (Join-Path $PSScriptRoot 'ExcelRunnerPolicy.ps1')

function Invoke-ExcelRunnerGithub {
    param([string]$Endpoint, [switch]$Paginate, [ValidateSet('GET', 'POST')][string]$Method = 'GET')
    $remaining = [int]($deadline - [DateTime]::UtcNow).TotalSeconds
    if ($remaining -le 0) { throw 'GitHub runner control exceeded its deadline.' }
    $job = Start-Job -ScriptBlock {
        param($Endpoint, $Paginate, $Method)
        $arguments = @('api', $Endpoint, '--method', $Method)
        if ($Paginate) { $arguments += @('--paginate', '--slurp') }
        $output = & gh @arguments
        if ($LASTEXITCODE -ne 0) { throw 'GitHub runner status query failed; do not assume the machine is idle.' }
        return $output
    } -ArgumentList $Endpoint, $Paginate.IsPresent, $Method
    try {
        if (-not (Wait-Job $job -Timeout ([Math]::Min(120, $remaining)))) {
            Stop-Job $job
            throw 'GitHub runner status query timed out; keep active work undisturbed.'
        }
        $output = Receive-Job $job -ErrorAction Stop
        if ($job.State -ne 'Completed') { throw 'GitHub runner status query did not complete.' }
        $parsed = ($output -join "`n") | ConvertFrom-Json
        return ,$parsed
    }
    finally { Remove-Job $job -Force }
}

function Get-ExcelRunnerActiveJobs {
    param([string]$Repository)
    $jobs = [Collections.Generic.List[object]]::new()
    $repositoryInfo = Invoke-ExcelRunnerGithub "repos/$Repository"
    foreach ($status in @('queued', 'in_progress', 'waiting', 'pending', 'requested')) {
        $pages = Invoke-ExcelRunnerGithub "repos/$Repository/actions/runs?status=$status&per_page=100" -Paginate
        $runs = @($pages | ForEach-Object { $_.workflow_runs })
        if (@($pages | Where-Object { $_.total_count -gt 1000 }).Count) { throw 'GitHub run discovery was truncated.' }
        foreach ($run in $runs) {
            $jobPages = Invoke-ExcelRunnerGithub "repos/$Repository/actions/runs/$($run.id)/jobs?filter=latest&per_page=100" -Paginate
            foreach ($item in @($jobPages | ForEach-Object { $_.jobs })) {
                if ($item.status -eq 'completed') { continue }
                $item | Add-Member NoteProperty excelRunId $run.id -Force
                $item | Add-Member NoteProperty trustedCloudRun (Test-ExcelRunnerCloudRun $run $Repository) -Force
                $validation = $run.event -eq 'workflow_dispatch' -and
                    $run.path -eq '.github/workflows/excel-runner-validation.yml' -and
                    $run.head_repository.full_name -eq $Repository -and
                    $run.head_branch -eq $repositoryInfo.default_branch -and
                    $run.actor.login -eq $repositoryInfo.owner.login
                $item | Add-Member NoteProperty trustedValidationRun $validation -Force
                $jobs.Add($item)
            }
            if ((Test-ExcelRunnerCloudRun $run $Repository) -and
                @($jobPages | ForEach-Object { $_.jobs }).Count -eq 0) {
                throw 'An active cloud-agent run has no job record yet; its ownership is uncertain.'
            }
        }
    }
    return $jobs.ToArray()
}

function Invoke-ExcelRunnerGuest {
    param([string]$Script, [string]$Marker = 'EXCELMCP_CONTROL=', [string]$OperationId)
    $path = Join-Path ([IO.Path]::GetTempPath()) "excel-control-$([Guid]::NewGuid().ToString('N')).ps1"
    try {
        ("`$ErrorActionPreference = 'Stop'`n" + $Script) | Set-Content -LiteralPath $path -Encoding UTF8
        $response = Invoke-RunnerAzure @(
            'vm', 'run-command', 'invoke', '--resource-group', $ResourceGroup, '--name', $VmName,
            '--command-id', 'RunPowerShellScript', '--scripts', "@$path"
        ) 300
        $matches = @(
            foreach ($value in $response.value) {
                foreach ($line in ($value.message -split "`r?`n")) {
                    if ($line.StartsWith($Marker)) { $line.Substring($Marker.Length) }
                }
            }
        )
        if ($matches.Count -ne 1) { throw ('Guest returned no unique control result: ' + (@($response.value.message) -join "`n")) }
        $parsed = $matches[0] | ConvertFrom-Json
        if ($OperationId -and $parsed.operationId -ne $OperationId -and $parsed.passId -ne $OperationId) {
            throw 'Guest returned a result for another operation.'
        }
        return $parsed
    }
    finally { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
}

function Get-ExcelRunnerGuestActivity {
    Invoke-ExcelRunnerGuest @'
$task = Get-ScheduledTask -TaskName 'ExcelMcp-GitHub-Runner' -ErrorAction SilentlyContinue
$state = @{
    listeners = @(Get-Process -Name Runner.Listener -ErrorAction SilentlyContinue).Count
    workers = @(Get-Process -Name Runner.Worker -ErrorAction SilentlyContinue).Count
    excel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
    taskRunning = [bool]($task -and $task.State -eq 'Running')
}
$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ (Get-LocalUser excelrunner).SID.Value)
if ($profile.Count -ne 1) { throw 'Dedicated profile ownership is uncertain.' }
$records = Join-Path $profile[0].LocalPath 'AppData\Local\ExcelMcp\Jobs'
$state.cleanupPending = [bool]((Test-Path -LiteralPath $records) -and @(Get-ChildItem -LiteralPath $records -File -Filter '*.json').Count)
Write-Output ('EXCELMCP_CONTROL=' + ($state | ConvertTo-Json -Compress))
'@
}

function Assert-ExcelRunnerOwnedVm {
    $vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
    if ($vm.tags.project -ne 'mcp-server-excel' -or $vm.tags.purpose -ne 'excel-cloud-agent-devtest') {
        throw 'The VM is not owned by this Excel runner deployment.'
    }
    $vaults = @(Invoke-RunnerAzure @('keyvault', 'list', '--resource-group', $ResourceGroup))
    $owned = @($vaults | Where-Object { $_.tags.purpose -eq 'excel-cloud-agent-devtest' })
    if ($owned.Count -ne 1) { throw 'Expected one dedicated runner password vault.' }
    $vault = Invoke-RunnerAzure @('keyvault', 'show', '--name', $owned[0].name)
    Assert-RunnerIdentityWithoutAzureAccess $vm @($vault.properties.accessPolicies)
    return $vm
}

function Wait-ExcelRunnerGuestAgent {
    $until = [DateTime]::UtcNow.AddMinutes(15)
    do {
        $vm = Invoke-RunnerAzure @('vm', 'get-instance-view', '--resource-group', $ResourceGroup, '--name', $VmName)
        if (@($vm.instanceView.vmAgent.statuses | Where-Object { $_.displayStatus -eq 'Ready' }).Count) { return }
        Start-Sleep -Seconds 15
    } while ([DateTime]::UtcNow -lt $until)
    throw 'The Azure guest agent did not recover within 15 minutes.'
}

function Send-ExcelRunnerFiles {
    param([ValidateSet('Maintenance', 'Desktop')][string]$Directory, [string[]]$Names)
    $builder = [Text.StringBuilder]::new()
    [void]$builder.AppendLine("`$directory = 'C:\ProgramData\ExcelMcp\$Directory'")
    [void]$builder.AppendLine(@'
if (@(Get-Process -Name Runner.Listener, Runner.Worker, EXCEL -ErrorAction SilentlyContinue).Count) {
    throw 'Protected control files must not change while work is running.'
}
foreach ($name in @('ExcelMcp-Update', 'ExcelMcp-Test-Desktop', 'ExcelMcp-Register-Runner', 'ExcelMcp-Recover-Jobs')) {
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    if ($task -and $task.State -eq 'Running') { throw 'Another guest controller is running.' }
}
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$sid = (Get-LocalUser excelrunner).SID.Value
& icacls.exe $directory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' "*${sid}:(OI)(CI)RX" | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Could not protect controller files.' }
'@)
    $root = Split-Path -Parent $PSScriptRoot
    foreach ($name in $Names) {
        if ($name -notmatch '^[a-zA-Z0-9-]+\.(ps1|json)$') { throw 'Invalid control file name.' }
        $source = switch ($name) {
            'ExcelRunnerPolicy.ps1' { Join-Path $PSScriptRoot $name }
            'global.json' { Join-Path $root $name }
            default { Join-Path $root "infrastructure\azure\$name" }
        }
        $encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
        $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        [void]$builder.AppendLine(@"
[IO.File]::WriteAllBytes((Join-Path `$directory '$name'), [Convert]::FromBase64String('$encoded'))
if ((Get-FileHash -LiteralPath (Join-Path `$directory '$name') -Algorithm SHA256).Hash -ne '$hash') { throw 'Controller hash mismatch.' }
"@)
    }
    [void]$builder.AppendLine("Write-Output 'EXCELMCP_CONTROL={`"state`":`"uploaded`"}'")
    $state = Invoke-ExcelRunnerGuest $builder.ToString()
    if ($state.state -ne 'uploaded') { throw 'Control files were not installed.' }
}

function Invoke-ExcelRunnerDesktopHealth {
    $operation = [Guid]::NewGuid().ToString('N')
    $script = @"
`$info = Get-ScheduledTaskInfo -TaskName 'ExcelMcp-Initialize-Desktop'
`$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
if (`$info.LastTaskResult -ne 0 -or `$info.LastRunTime -lt `$boot) { throw 'Profile startup has not succeeded during this boot.' }
`$user = Get-LocalUser excelrunner
`$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ `$user.SID.Value)
if (`$profile.Count -ne 1 -or -not `$profile[0].Loaded) { throw 'The dedicated profile is not loaded.' }
`$report = Join-Path `$profile[0].LocalPath 'AppData\Local\ExcelMcp\Health\$operation.json'
if (Test-Path -LiteralPath `$report) { throw 'Health operation already exists.' }
`$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\ProgramData\ExcelMcp\Desktop\test-excel-desktop.ps1" -OperationId $operation'
`$principal = New-ScheduledTaskPrincipal -UserId "`$env:COMPUTERNAME\excelrunner" -LogonType Interactive -RunLevel Limited
`$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 2)
Register-ScheduledTask -TaskName 'ExcelMcp-Test-Desktop' -Action `$action -Principal `$principal -Settings `$settings -Force | Out-Null
Start-ScheduledTask -TaskName 'ExcelMcp-Test-Desktop'
Write-Output 'EXCELMCP_CONTROL={"state":"running","operationId":"$operation"}'
"@
    $state = Invoke-ExcelRunnerGuest $script -OperationId $operation
    $until = [DateTime]::UtcNow.AddMinutes(4)
    do {
        Start-Sleep -Seconds 10
        $state = Invoke-ExcelRunnerGuest @"
`$task = Get-ScheduledTask -TaskName 'ExcelMcp-Test-Desktop'
`$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ (Get-LocalUser excelrunner).SID.Value)
`$report = Join-Path `$profile[0].LocalPath 'AppData\Local\ExcelMcp\Health\$operation.json'
if (`$task.State -eq 'Running') { `$state = @{ state = 'running'; operationId = '$operation' } }
else {
    if (-not (Test-Path -LiteralPath `$report)) { throw 'The desktop task stopped without a report.' }
    `$state = Get-Content -LiteralPath `$report -Raw | ConvertFrom-Json
    if ((Get-ScheduledTaskInfo -TaskName 'ExcelMcp-Test-Desktop').LastTaskResult -ne 0) {
        throw ('Desktop health task failed: ' + (`$state | ConvertTo-Json -Compress))
    }

    . 'C:\ProgramData\ExcelMcp\Desktop\ExcelRunnerPolicy.ps1'
    Assert-ExcelRunnerReadiness `$state -BootTime (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    `$state | ConvertTo-Json -Compress | Set-Content 'C:\ProgramData\ExcelMcp\Desktop\ready.json' -Encoding UTF8
}
Write-Output ('EXCELMCP_CONTROL=' + (`$state | ConvertTo-Json -Compress))
"@ -OperationId $operation
        if ($state.state -ne 'running') { Assert-ExcelRunnerReadiness $state; return $state }
    } while ([DateTime]::UtcNow -lt $until)
    throw 'The non-admin Excel desktop health task exceeded its deadline.'
}

function Invoke-ExcelRunnerJobRecovery {
    $operation = [Guid]::NewGuid().ToString('N')
    $null = Invoke-ExcelRunnerGuest @"
if (@(Get-Process -Name Runner.Listener, Runner.Worker -ErrorAction SilentlyContinue).Count) { throw 'Work is still active.' }
`$action = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument '-NoProfile -NonInteractive -File "C:\ProgramData\ExcelMcp\Desktop\recover-excel-jobs.ps1" -OperationId $operation'
`$principal = New-ScheduledTaskPrincipal -UserId "`$env:COMPUTERNAME\excelrunner" -LogonType Interactive -RunLevel Limited
`$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 3)
Register-ScheduledTask -TaskName 'ExcelMcp-Recover-Jobs' -Action `$action -Principal `$principal -Settings `$settings -Force | Out-Null
Start-ScheduledTask -TaskName 'ExcelMcp-Recover-Jobs'
Write-Output 'EXCELMCP_CONTROL={"state":"started"}'
"@
    $until = [DateTime]::UtcNow.AddMinutes(4)
    do {
        Start-Sleep -Seconds 10
        $state = Invoke-ExcelRunnerGuest @"
`$task = Get-ScheduledTask -TaskName 'ExcelMcp-Recover-Jobs'
`$state = @{ state = 'running'; operationId = '$operation' }
if (`$task.State -ne 'Running') {
    if ((Get-ScheduledTaskInfo -TaskName 'ExcelMcp-Recover-Jobs').LastTaskResult -ne 0) { throw 'Limited job recovery failed; retain quarantine.' }
    `$profile = @(Get-CimInstance Win32_UserProfile | Where-Object SID -EQ (Get-LocalUser excelrunner).SID.Value)
    `$path = Join-Path `$profile[0].LocalPath 'AppData\Local\ExcelMcp\Recovery\$operation.json'
    `$state = Get-Content `$path -Raw | ConvertFrom-Json
}
Write-Output ('EXCELMCP_CONTROL=' + (`$state | ConvertTo-Json -Compress))
"@ -OperationId $operation
        if ($state.state -eq 'recovered') { return }
        if ($state.state -ne 'running') { throw 'Job recovery did not establish cleanup.' }
    } while ([DateTime]::UtcNow -lt $until)
    throw 'Limited job recovery exceeded its deadline.'
}

function Stop-ExcelRunnerExpiredListener {
    param([object[]]$Jobs)
    if (@($Jobs | Where-Object { (Test-ExcelRunnerJobTarget $_) -and $_.status -ne 'queued' -and $_.status -ne 'completed' }).Count) { return }
    $null = Invoke-ExcelRunnerGuest @'
. 'C:\ProgramData\ExcelMcp\Desktop\ExcelRunnerPolicy.ps1'
$permit = Get-Content 'C:\ProgramData\ExcelMcp\Desktop\permit.json' -Raw | ConvertFrom-Json
$state = @{ state = 'waiting' }
if ((ConvertTo-ExcelRunnerUtc $permit.expiresAt) -lt [DateTime]::UtcNow) {
    if (@(Get-Process -Name Runner.Worker -ErrorAction SilentlyContinue).Count) { throw 'A job started; do not drain.' }
    $listeners = @(Get-Process -Name Runner.Listener -IncludeUserName -ErrorAction SilentlyContinue)
    foreach ($listener in $listeners) {
        try {
            $null = $listener.Handle
            $start = $listener.StartTime
            if ($listener.UserName -ine "$env:COMPUTERNAME\excelrunner" -or
                $listener.Path -ine 'C:\actions-runner\bin\Runner.Listener.exe') { throw 'Unexpected listener ownership.' }
            $current = Get-Process -Id $listener.Id -ErrorAction SilentlyContinue
            try {
                if ($current -and $current.StartTime -eq $start) { Stop-Process -Id $current.Id -Force }
            }
            finally { if ($current) { $current.Dispose() } }
            if (-not $listener.WaitForExit(10000)) { throw 'Expired listener did not stop.' }
        }
        finally { $listener.Dispose() }
    }
    $state.state = 'drained'
}
Write-Output ('EXCELMCP_CONTROL=' + ($state | ConvertTo-Json -Compress))
'@
}
