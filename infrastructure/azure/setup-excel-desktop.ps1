<#
.SYNOPSIS
Prepares a dedicated non-admin interactive Excel desktop.
.DESCRIPTION
Adapts mcp-windows secure automatic logon and interactive scheduled-task setup.
The temporary VM identity reads only the dedicated runner password from Key Vault.
The host must remove that vault grant before accepting agent work. It removes only
identities it created; a guarded policy-owned identity is preserved without access.
This script does not register or start a GitHub runner.
#>
param(
    [ValidateSet('Setup', 'Health', 'Activation')]
    [string]$Action,
    [ValidatePattern('^[a-zA-Z][a-zA-Z0-9-]{1,22}[a-zA-Z0-9]$')]
    [string]$KeyVaultName,
    [ValidatePattern('^[a-zA-Z][a-zA-Z0-9-]{1,18}$')]
    [string]$WindowsUser = 'excelrunner'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Assert-NonAdminDesktopAccount {
    param([string]$Name)
    $user = Get-LocalUser -Name $Name
    if (-not $user.Enabled) { throw 'The desktop account must be enabled.' }
    $administrators = @(Get-LocalGroupMember -SID 'S-1-5-32-544')
    if (@($administrators | Where-Object { $_.SID.Value -eq $user.SID.Value }).Count -gt 0) {
        throw 'The desktop account must not be a local administrator.'
    }
    return $user
}

function Get-DesktopBootstrapPassword {
    $tokenUri = 'http://169.254.169.254/metadata/identity/oauth2/token' +
        '?api-version=2018-02-01&resource=https%3A%2F%2Fvault.azure.net'
    $token = Invoke-RestMethod -Headers @{ Metadata = 'true' } -Uri $tokenUri
    if (-not $token.access_token) { throw 'The temporary VM identity returned no access token.' }
    $headers = @{ Authorization = "Bearer $($token.access_token)" }
    $secret = Invoke-RestMethod -Headers $headers -Uri (
        "https://$KeyVaultName.vault.azure.net/secrets/runner-account-password?api-version=7.4"
    )
    if (-not $secret.value) { throw 'The dedicated desktop account password is missing.' }
    return $secret.value
}

function Assert-DesktopProfileRun {
    param($TaskInfo, [DateTime]$BootTime)
    if ($TaskInfo.LastTaskResult -ne 0 -or $TaskInfo.LastRunTime -lt $BootTime) {
        throw 'Desktop profile initialization must succeed during the current boot.'
    }
}

function Enable-PrivateActivationDesktop {
    param([string]$Name)
    if (@(Get-Process -Name Runner.Listener -ErrorAction SilentlyContinue).Count -gt 0) {
        throw 'Stop the coding runner before interactive activation.'
    }
    $user = Assert-NonAdminDesktopAccount -Name $Name
    $members = @(Get-LocalGroupMember -SID 'S-1-5-32-555')
    if (@($members | Where-Object { $_.SID.Value -eq $user.SID.Value }).Count -eq 0) {
        Add-LocalGroupMember -SID 'S-1-5-32-555' -Member "$env:COMPUTERNAME\$Name"
    }
    Set-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' `
        -Name fDenyTSConnections -Value 0
    Enable-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-TCP'
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $Action) { throw 'An explicit desktop setup action is required.' }

$account = "$env:COMPUTERNAME\$WindowsUser"
$directory = 'C:\ProgramData\ExcelMcp\Desktop'
$taskName = 'ExcelMcp-Initialize-Desktop'

switch ($Action) {
    'Activation' {
        Enable-PrivateActivationDesktop -Name $WindowsUser
        Write-Output 'EXCELMCP_DESKTOP={"state":"activation-access-enabled","runnerRegistered":false}'
    }
    'Setup' {
        if (-not $KeyVaultName) { throw 'KeyVaultName is required for desktop setup.' }
        if (@(Get-Process -Name Runner.Listener -ErrorAction SilentlyContinue).Count -gt 0) {
            throw 'Stop the coding runner before changing desktop setup.'
        }
        $password = Get-DesktopBootstrapPassword
        try {
            $user = Get-LocalUser -Name $WindowsUser -ErrorAction SilentlyContinue
            if (-not $user) {
                $user = New-LocalUser -Name $WindowsUser `
                    -Password (ConvertTo-SecureString $password -AsPlainText -Force) `
                    -PasswordNeverExpires -AccountNeverExpires `
                    -Description 'Dedicated non-admin Excel development runner'
            }
            $user = Assert-NonAdminDesktopAccount -Name $WindowsUser
            $members = @(Get-LocalGroupMember -SID 'S-1-5-32-545')
            if (@($members | Where-Object { $_.SID.Value -eq $user.SID.Value }).Count -eq 0) {
                Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $account
            }
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            & icacls.exe $directory /inheritance:r /grant:r `
                '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' "*$($user.SID.Value):(OI)(CI)RX" | Out-Null
            if ($LASTEXITCODE -ne 0) { throw 'Could not protect desktop startup scripts.' }
            $archive = Join-Path $directory 'Autologon.zip'
            $extract = Join-Path $directory 'Autologon'
            try {
                Invoke-WebRequest -Uri 'https://download.sysinternals.com/files/AutoLogon.zip' `
                    -OutFile $archive -UseBasicParsing
                Expand-Archive -LiteralPath $archive -DestinationPath $extract -Force
                $autologon = Join-Path $extract 'Autologon64.exe'
                $signature = Get-AuthenticodeSignature -LiteralPath $autologon
                if ($signature.Status -ne 'Valid' -or
                    $signature.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,') {
                    throw 'Sysinternals Autologon does not have a valid Microsoft signature.'
                }
                & $autologon $WindowsUser $env:COMPUTERNAME $password '/accepteula' | Out-Null
                if ($LASTEXITCODE -ne 0) { throw 'Secure automatic logon setup failed.' }
            }
            finally {
                if (Test-Path -LiteralPath $archive) { Remove-Item -LiteralPath $archive -Force }
                if (Test-Path -LiteralPath $extract) { Remove-Item -LiteralPath $extract -Recurse -Force }
            }
        }
        finally { $password = $null }

        Set-WinSystemLocale -SystemLocale 'en-US'
        foreach ($setting in @('standby-timeout-ac', 'monitor-timeout-ac')) {
            & powercfg.exe /change $setting 0
            if ($LASTEXITCODE -ne 0) { throw "Power policy failed: $setting" }
        }
        $profilePath = Join-Path $directory 'initialize-profile.ps1'
        @'
$ErrorActionPreference = 'Stop'
Set-Culture -CultureInfo 'en-US'
Set-WinUserLanguageList -LanguageList 'en-US' -Force
$desktop = 'HKCU:\Control Panel\Desktop'
Set-ItemProperty -LiteralPath $desktop -Name ScreenSaveActive -Value '0'
Set-ItemProperty -LiteralPath $desktop -Name ScreenSaverIsSecure -Value '0'
$security = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security'
New-Item -Path $security -Force | Out-Null
New-ItemProperty -Path $security -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
'@ | Set-Content -LiteralPath $profilePath -Encoding UTF8
        $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
            "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$profilePath`""
        )
        $principal = New-ScheduledTaskPrincipal -UserId $account -LogonType Interactive -RunLevel Limited
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $account
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $principal `
            -Trigger $trigger -Settings $settings -Force | Out-Null
        Write-Output 'EXCELMCP_DESKTOP={"state":"configured","runnerRegistered":false}'
    }
    'Health' {
        $user = Assert-NonAdminDesktopAccount -Name $WindowsUser
        $explorers = @(Get-Process -Name explorer -IncludeUserName -ErrorAction SilentlyContinue |
            Where-Object { $_.UserName -ieq $account -and $_.SessionId -gt 0 })
        $task = Get-ScheduledTask -TaskName $taskName
        $info = Get-ScheduledTaskInfo -TaskName $taskName
        if ($explorers.Count -ne 1 -or $task.State -eq 'Running') {
            throw 'The dedicated desktop or profile initialization is not ready.'
        }
        $bootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
        Assert-DesktopProfileRun -TaskInfo $info -BootTime $bootTime
        Write-Output ('EXCELMCP_DESKTOP=' + (@{
            state = 'ready'
            account = $WindowsUser
            sessionId = $explorers[0].SessionId
            runnerRegistered = $false
        } | ConvertTo-Json -Compress))
    }
}
