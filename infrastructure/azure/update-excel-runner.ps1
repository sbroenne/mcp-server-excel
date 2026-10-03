<#
.SYNOPSIS
Runs bounded Windows and Click-to-Run update passes, without restarting.
.DESCRIPTION
Adapted from mcp-windows update-runner.ps1 at
a76b3eeb075419986d0b00721bae5cdc896c2f81 (MIT, Copyright (c) 2025 Sbroenne).
The hosted controller owns restarts, clean rescans and desktop qualification.
#>
param(
    [ValidateSet('Start', 'Status', 'Worker')]
    [string]$Action,
    [ValidateSet('Windows', 'Office')]
    [string]$Component = 'Windows',
    [ValidatePattern('^[a-f0-9]{32}$')]
    [string]$PassId
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Release-ExcelRunnerCom {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Test-ExcelRunnerUpdate {
    param($Update)
    $collection = $null
    $categories = [Collections.Generic.List[string]]::new()
    try {
        $collection = $Update.Categories
        foreach ($category in $collection) {
            try { $categories.Add($category.CategoryID) }
            finally { Release-ExcelRunnerCom $category }
        }
    }
    finally { Release-ExcelRunnerCom $collection }
    $approved = @(
        '0fa1201d-4330-4fa8-8ae9-b877473b6441',
        'e6cf1350-c01b-414d-a61f-263d14d133b4',
        '28bc880e-0592-4cbf-8f95-c79b17911d5f',
        'e0789628-ce08-4437-be74-2495b842f43b',
        'cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83'
    )
    return $Update.Type -eq 1 -and -not $Update.BrowseOnly -and
        -not $categories.Contains('3689bdc8-b205-4af4-8d4a-a63924c5e9d5') -and
        @($categories | Where-Object { $_ -in $approved }).Count -gt 0
}

function Assert-ExcelRunnerUpdateResult {
    param($Result, [string]$Operation)
    if ([Convert]::ToInt32($Result.ResultCode) -ne 2) {
        throw "$Operation failed: result=$($Result.ResultCode), HRESULT=$($Result.HResult)."
    }
}

function Test-ExcelRunnerRestart {
    $info = New-Object -ComObject Microsoft.Update.SystemInfo
    try {
        return $info.RebootRequired -or
            (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
            (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    }
    finally { Release-ExcelRunnerCom $info }
}

function Invoke-ExcelRunnerWindowsUpdate {
    if (Test-ExcelRunnerRestart) { return @{ installed = 0; rebootRequired = $true } }
    $references = [Collections.Generic.List[object]]::new()
    function Own-UpdateCom { param($Value) $references.Add($Value); return ,$Value }
    try {
        $session = Own-UpdateCom (New-Object -ComObject Microsoft.Update.Session)
        $session.ClientApplicationID = 'ExcelMcp-Runner-Maintenance'
        $searcher = Own-UpdateCom ($session.CreateUpdateSearcher())
        $searcher.Online = $true
        $search = Own-UpdateCom ($searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software' and BrowseOnly=0"))
        Assert-ExcelRunnerUpdateResult $search 'Search'
        $found = Own-UpdateCom $search.Updates
        $updates = Own-UpdateCom (New-Object -ComObject Microsoft.Update.UpdateColl)
        $selected = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $found.Count; $i++) {
            $update = Own-UpdateCom ($found.Item($i))
            if (Test-ExcelRunnerUpdate $update) { $selected.Add($update) }
        }
        $exclusive = $null
        foreach ($update in $selected) {
            $behavior = Own-UpdateCom $update.InstallationBehavior
            if ($behavior.CanRequestUserInput) { throw 'A selected Windows update requires interaction.' }
            if ($behavior.Impact -eq 2 -and -not $exclusive) { $exclusive = $update }
        }
        if ($exclusive) { $selected.Clear(); $selected.Add($exclusive) }
        foreach ($update in $selected) {
            if (-not $update.EulaAccepted) { $update.AcceptEula() }
            [void]$updates.Add($update)
        }
        if ($updates.Count -eq 0) { return @{ installed = 0; rebootRequired = (Test-ExcelRunnerRestart) } }
        $downloader = Own-UpdateCom ($session.CreateUpdateDownloader())
        $downloader.Updates = $updates
        $download = Own-UpdateCom ($downloader.Download())
        Assert-ExcelRunnerUpdateResult $download 'Download'
        for ($i = 0; $i -lt $updates.Count; $i++) {
            $itemResult = Own-UpdateCom ($download.GetUpdateResult($i))
            Assert-ExcelRunnerUpdateResult $itemResult "Download item $i"
        }
        $installer = Own-UpdateCom ($session.CreateUpdateInstaller())
        $installer.Updates = $updates
        $installer.AllowSourcePrompts = $false
        $installer.ForceQuiet = $true
        $installed = Own-UpdateCom ($installer.Install())
        Assert-ExcelRunnerUpdateResult $installed 'Install'
        for ($i = 0; $i -lt $updates.Count; $i++) {
            $itemResult = Own-UpdateCom ($installed.GetUpdateResult($i))
            Assert-ExcelRunnerUpdateResult $itemResult "Install item $i"
        }
        return @{ installed = $updates.Count; rebootRequired = ($installed.RebootRequired -or (Test-ExcelRunnerRestart)) }
    }
    finally {
        for ($i = $references.Count - 1; $i -ge 0; $i--) { Release-ExcelRunnerCom $references[$i] }
    }
}

function Get-ExcelRunnerOfficeTarget {
    param([object[]]$Metadata)
    $current = @($Metadata | Where-Object channelId -eq 'Current')
    if ($current.Count -ne 1 -or $current[0].latestVersion -notmatch '^16\.0\.\d+\.\d+$') {
        throw 'Microsoft did not return exactly one valid Current channel version.'
    }
    return $current[0].latestVersion
}

function Invoke-ExcelRunnerOfficeUpdate {
    . (Join-Path $PSScriptRoot 'install-excel-office.ps1')
    $before = Get-VerifiedExcelInstallation
    $configuration = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if ($configuration.CDNBaseUrl -notmatch '^https?://officecdn\.microsoft\.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/?$') {
        throw 'Excel must remain on the Current Click-to-Run channel.'
    }
    $metadata = Invoke-RestMethod -Uri 'https://clients.config.office.net/releases/v1.0/OfficeReleases' -TimeoutSec 60
    $target = Get-ExcelRunnerOfficeTarget $metadata
    $client = Join-Path $env:ProgramFiles 'Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    Assert-MicrosoftOfficeSignature $client
    if ([version]$before.version -lt [version]$target) {
        $exitCode = Invoke-OfficeSetupProcess $client "/update user displaylevel=false forceappshutdown=false updatepromptuser=false updatetoversion=$target" 300
        if ($exitCode -ne 0) { throw "Click-to-Run request failed with exit code $exitCode." }
        # The client starts asynchronous servicing; its exit code is not completion.
        $until = [DateTime]::UtcNow.AddMinutes(40)
        do {
            Start-Sleep -Seconds 20
            $after = Get-VerifiedExcelInstallation
            if ($after.version -eq $target) { break }
        } while ([DateTime]::UtcNow -lt $until)
        if ($after.version -ne $target) { throw 'Excel did not reach the requested Current version before its deadline.' }
    }
    $after = Get-VerifiedExcelInstallation
    return @{ officeVersion = $after.version; officeTarget = $target; rebootRequired = (Test-ExcelRunnerRestart) }
}

function Write-ExcelRunnerUpdateState {
    param([hashtable]$State)
    $State.passId = $PassId
    $State.component = $Component
    $State.checkedAt = [DateTime]::UtcNow.ToString('o')
    $State | ConvertTo-Json -Compress | Set-Content -LiteralPath "$resultPath.tmp" -Encoding UTF8
    if (Test-Path -LiteralPath $resultPath) { [IO.File]::Replace("$resultPath.tmp", $resultPath, [NullString]::Value) }
    else { [IO.File]::Move("$resultPath.tmp", $resultPath) }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Action -or -not $PassId) { throw 'An explicit action and unique pass ID are required.' }
$directory = 'C:\ProgramData\ExcelMcp\Maintenance'
$resultPath = Join-Path $directory "$PassId.json"
$taskName = 'ExcelMcp-Update'
switch ($Action) {
    'Start' {
        if (Test-Path -LiteralPath $resultPath) { throw 'This update pass already exists.' }
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task -and $task.State -eq 'Running') { throw 'An update pass is already running.' }
        if (@(Get-Process -Name Runner.Listener, Runner.Worker, EXCEL -ErrorAction SilentlyContinue).Count) {
            throw 'Updates require the listener, coding worker and Excel to be stopped.'
        }
        Write-ExcelRunnerUpdateState @{ state = 'running' }
        $policy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
        New-Item $policy -Force | Out-Null
        New-ItemProperty $policy -Name ExcludeWUDriversInQualityUpdate -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty $policy -Name TargetReleaseVersion -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty $policy -Name ProductVersion -Value 'Windows 11' -PropertyType String -Force | Out-Null
        New-ItemProperty $policy -Name TargetReleaseVersionInfo -Value '25H2' -PropertyType String -Force | Out-Null
        New-Item "$policy\AU" -Force | Out-Null
        New-ItemProperty "$policy\AU" -Name NoAutoUpdate -Value 0 -PropertyType DWord -Force | Out-Null
        New-ItemProperty "$policy\AU" -Name AUOptions -Value 2 -PropertyType DWord -Force | Out-Null
        New-ItemProperty "$policy\AU" -Name NoAutoRebootWithLoggedOnUsers -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' `
            -Name UpdatesEnabled -Value 'False' -PropertyType String -Force | Out-Null
        $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
            "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Worker -Component $Component -PassId $PassId"
        )
        $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 60)
        Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $principal -Settings $settings -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
    }
    'Worker' {
        try {
            if (@(Get-Process -Name Runner.Listener, Runner.Worker, EXCEL -ErrorAction SilentlyContinue).Count) {
                throw 'Work started during maintenance; refusing to install updates.'
            }
            $state = if ($Component -eq 'Windows') { Invoke-ExcelRunnerWindowsUpdate } else { Invoke-ExcelRunnerOfficeUpdate }
            $state.state = 'complete'
            Write-ExcelRunnerUpdateState $state
        }
        catch {
            Write-ExcelRunnerUpdateState @{ state = 'failed'; error = $_.Exception.Message }
            throw
        }
        return
    }
    'Status' {
        $state = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        if ($state.passId -ne $PassId -or $state.component -ne $Component) { throw 'Wrong maintenance result identity.' }
        if ($state.state -eq 'running' -and (Get-ScheduledTask -TaskName $taskName).State -ne 'Running') {
            $info = Get-ScheduledTaskInfo -TaskName $taskName
            Write-ExcelRunnerUpdateState @{ state = 'failed'; error = "Update task stopped without a result: $($info.LastTaskResult)." }
        }
    }
}
Write-Output ('EXCELMCP_UPDATE=' + (Get-Content -LiteralPath $resultPath -Raw).Trim())
