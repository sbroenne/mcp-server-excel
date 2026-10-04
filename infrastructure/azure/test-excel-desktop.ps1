<#
.SYNOPSIS
Qualifies licensed Excel and saved workbook behavior in the non-admin desktop.
.DESCRIPTION
The administrator-protected controller starts this as an interactive limited task.
It never removes account data, changes activation or reports account identifiers.
#>
param([ValidatePattern('^[a-f0-9]{32}$')][string]$OperationId)
$ErrorActionPreference = 'Stop'

function Test-ExcelRunnerDeviceLicense {
    param([object[]]$Records)
    $candidate = @{}
    foreach ($record in $Records) {
        if ($record -is [Management.Automation.ErrorRecord]) { throw 'Office licence diagnostic returned an error.' }
        if ($record.Type -eq 'Device|Perpetual' -and $record.Product -eq 'Excel2024Retail' -and $record.LicenseState -eq 'Licensed') {
            return $true
        }
        $text = if ($record -is [Management.Automation.InformationRecord] -and $record.MessageData.Message) {
            $record.MessageData.Message
        } else { $record.ToString() }
        foreach ($line in ($text -split "`r?`n")) {
            $match = [regex]::Match($line, '^\s*"?(Type|Product|LicenseState)"?\s*[:=]\s*"?([a-zA-Z0-9_| -]{1,80})"?[,]?\s*$')
            if (-not $match.Success) { continue }
            $field = $match.Groups[1].Value
            if ($field -eq 'Type' -or ($field -eq 'Product' -and $candidate.ContainsKey('Product'))) { $candidate = @{} }
            $candidate[$field] = $match.Groups[2].Value.Trim()
            if ($candidate.Type -eq 'Device|Perpetual' -and $candidate.Product -eq 'Excel2024Retail' -and
                $candidate.LicenseState -eq 'Licensed') { return $true }
        }
    }
    return $false
}

function Assert-ExcelRunnerAccountPrivacy {
    foreach ($path in @(
        'HKCU:\Software\Microsoft\IdentityCRL\StoredIdentities',
        'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities',
        'HKCU:\Software\Microsoft\IdentityCRL\UserExtendedProperties'
    )) {
        if ((Test-Path -LiteralPath $path) -and @(Get-ChildItem -LiteralPath $path).Count) {
            throw 'Personal account registration prevents runner readiness.'
        }
    }
    $oneDrive = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path -LiteralPath $oneDrive) {
        foreach ($entry in Get-ChildItem -LiteralPath $oneDrive) {
            if ((Get-ItemProperty -LiteralPath $entry.PSPath).UserEmail) { throw 'OneDrive remains signed in.' }
        }
    }
    foreach ($relative in @(
        'Microsoft\TokenBroker\Accounts',
        'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts',
        'Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts'
    )) {
        $path = Join-Path $env:LOCALAPPDATA $relative
        if ((Test-Path -LiteralPath $path) -and @(Get-ChildItem -LiteralPath $path -File -Recurse).Count) {
            throw 'A broker account remains registered.'
        }
    }
    foreach ($relative in @('Microsoft\Edge\User Data', 'Google\Chrome\User Data')) {
        $path = Join-Path $env:LOCALAPPDATA $relative
        if (-not (Test-Path -LiteralPath $path)) { continue }
        foreach ($profile in Get-ChildItem -LiteralPath $path -Directory | Where-Object Name -Match '^(Default|Profile \d+)$') {
            $preferences = Join-Path $profile.FullName 'Preferences'
            if (Test-Path -LiteralPath $preferences) {
                $settings = Get-Content -LiteralPath $preferences -Raw | ConvertFrom-Json
                if (@($settings.account_info | Where-Object { $null -ne $_ }).Count) { throw 'A browser remains signed in.' }
            }
        }
    }
    if ((Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive').DisableFileSyncNGSC -ne 1 -or
        @(Get-Process -Name OneDrive -ErrorAction SilentlyContinue).Count) {
        throw 'OneDrive synchronization must remain disabled.'
    }
    if ([ExcelRunnerDesktopNative]::HasAccountCredentials()) { throw 'Personal or unclassified saved credentials remain.' }
}

if ($MyInvocation.InvocationName -eq '.') { return }
if (-not $OperationId) { throw 'A unique health operation ID is required.' }
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class ExcelRunnerDesktopNative {
    [StructLayout(LayoutKind.Sequential)]
    private struct Credential {
        public uint Flags, Type;
        public IntPtr Target, Comment;
        public long Written;
        public uint BlobSize;
        public IntPtr Blob;
        public uint Persist, AttributeCount;
        public IntPtr Attributes, Alias, User;
    }
    [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern bool CredEnumerateW(string filter, uint flags, out int count, out IntPtr buffer);
    [DllImport("advapi32.dll")] private static extern void CredFree(IntPtr buffer);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    public static bool HasAccountCredentials() {
        int count; IntPtr buffer;
        if (!CredEnumerateW(null, 1, out count, out buffer)) {
            int error = Marshal.GetLastWin32Error();
            if (error == 1168) return false;
            throw new Win32Exception(error);
        }
        try {
            for (int i=0; i<count; i++) {
                var credential = (Credential)Marshal.PtrToStructure(Marshal.ReadIntPtr(buffer,i*IntPtr.Size),typeof(Credential));
                string target = Marshal.PtrToStringUni(credential.Target) ?? "";
                string user = Marshal.PtrToStringUni(credential.User) ?? "";
                foreach (string prefix in new [] {"LegacyGeneric:target=", "Generic:target=", "Domain:target="})
                    if (target.StartsWith(prefix,StringComparison.OrdinalIgnoreCase)) target=target.Substring(prefix.Length);
                bool device = target == "MicrosoftAccount:target=SSO_POP_Device" || target == "SSO_POP_Device" ||
                    target == "WindowsLive:target=virtualapp/didlogical";
                if (!device || target.Contains("@") || user.Contains("@") || credential.BlobSize != 0) return true;
            }
            return false;
        } finally { CredFree(buffer); }
    }
}
'@
$directory = Join-Path $env:LOCALAPPDATA 'ExcelMcp\Health'
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$reportPath = Join-Path $directory "$OperationId.json"
$scratch = Join-Path $directory "$OperationId.xlsx"
$state = @{ operationId = $OperationId; state = 'failed' }
$references = [Collections.Generic.List[object]]::new()
$failures = [Collections.Generic.List[Exception]]::new()
$excel = $null
$activeBook = $null
$ownedProcess = $null
$ownedStartTime = $null
function Own-HealthCom { param($Value) $references.Add($Value); return ,$Value }
try {
    $state.stage = 'account'
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        if ($identity.Name -ine "$env:COMPUTERNAME\excelrunner" -or
            ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -or
            (Get-LocalUser excelrunner).PrincipalSource.ToString() -ne 'Local') {
            throw 'Health requires the dedicated local non-admin account.'
        }
    }
    finally { $identity.Dispose() }
    $session = (Get-Process -Id $PID).SessionId
    if (-not [Environment]::UserInteractive -or $session -le 0 -or
        @(Get-Process -Name explorer | Where-Object SessionId -EQ $session).Count -ne 1) {
        throw 'Health requires the actual interactive desktop.'
    }
    if (@(Get-Process -Name Runner.Listener, Runner.Worker, EXCEL -ErrorAction SilentlyContinue).Count) {
        throw 'Health must not overlap coding work or an existing workbook.'
    }
    $state.bootTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().ToString('o')
    Assert-ExcelRunnerAccountPrivacy
    $state.stage = 'licence'
    $configuration = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if ($configuration.Platform -ne 'x64' -or @($configuration.ProductReleaseIds -split ',') -notcontains 'Excel2024Retail') {
        throw 'Unexpected Excel product or architecture.'
    }
    $diag = Join-Path $configuration.InstallationPath 'root\Office16\vNextDiag.ps1'
    if (-not (Test-Path -LiteralPath $diag)) { throw 'Installed retail licence diagnostic is missing.' }
    Push-Location (Split-Path -Parent $diag)
    try { $records = @(& $diag -Action list *>&1) }
    finally { Pop-Location }
    if (@($records | Where-Object { $_ -is [Management.Automation.ErrorRecord] }).Count -or
        -not (Test-ExcelRunnerDeviceLicense $records)) {
        throw 'Microsoft did not establish a licensed Excel2024Retail perpetual device.'
    }
    $state.license = 'Licensed'
    $state.officeVersion = $configuration.VersionToReport
    if (Test-Path -LiteralPath $scratch) { throw 'Health must use a new owned workbook.' }
    $state.stage = 'workbook'
    $excel = Own-HealthCom (New-Object -ComObject Excel.Application)
    $excel.DisplayAlerts = $false
    [uint32]$excelPid = 0
    [void][ExcelRunnerDesktopNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    if ($excelPid -le 0) { throw 'Could not identify the health-owned Excel process.' }
    $ownedProcess = Get-Process -Id $excelPid
    if ($ownedProcess.ProcessName -ne 'EXCEL' -or $ownedProcess.SessionId -ne $session) {
        throw 'Health returned an unexpected Excel process identity.'
    }
    $ownedStartTime = $ownedProcess.StartTime
    $null = $ownedProcess.Handle
    $books = Own-HealthCom $excel.Workbooks
    $activeBook = Own-HealthCom ($books.Add())
    $sheets = Own-HealthCom $activeBook.Worksheets
    $sheet = Own-HealthCom ($sheets.Item(1))
    $cell = Own-HealthCom ($sheet.Range('A1'))
    $cell.Formula = '=SUM(20,22)'
    $excel.Calculate()
    $state.formula = [Convert]::ToInt32($cell.Value2)
    if ($state.formula -ne 42) { throw 'Excel calculation failed.' }
    $activeBook.SaveAs($scratch, 51)
    $activeBook.Close($false)
    $activeBook = $null
    $activeBook = Own-HealthCom ($books.Open($scratch))
    $sheets = Own-HealthCom $activeBook.Worksheets
    $sheet = Own-HealthCom ($sheets.Item(1))
    $cell = Own-HealthCom ($sheet.Range('A1'))
    $state.persistedFormula = [Convert]::ToInt32($cell.Value2)
    if ($state.persistedFormula -ne 42 -or $cell.Formula -ne '=SUM(20,22)') { throw 'Saved Excel readback failed.' }
}
catch { $failures.Add($_.Exception) }
finally {
    if ($activeBook) { try { $activeBook.Close($false) } catch { $failures.Add($_.Exception) } }
    if ($excel) { try { $excel.Quit() } catch { $failures.Add($_.Exception) } }
    for ($i = $references.Count - 1; $i -ge 0; $i--) {
        try {
            if ([Runtime.InteropServices.Marshal]::IsComObject($references[$i])) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($references[$i])
            }
        }
        catch { $failures.Add($_.Exception) }
    }
    if ($ownedProcess) {
        try {
            if (-not $ownedProcess.WaitForExit(15000)) {
                $current = Get-Process -Id $ownedProcess.Id -ErrorAction SilentlyContinue
                if ($current -and $current.StartTime -eq $ownedStartTime) { Stop-Process -Id $current.Id -Force }
                $failures.Add([Exception]::new('The health-owned Excel process did not exit cleanly.'))
            }
        }
        catch { $failures.Add($_.Exception) }
        finally { $ownedProcess.Dispose() }
    }
    try { if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Force } }
    catch { $failures.Add($_.Exception) }
    $state.checkedAt = [DateTime]::UtcNow.ToString('o')
    if ($failures.Count) { $state.errorTypes = @($failures | ForEach-Object { $_.GetType().FullName }) }
    else { $state.state = 'ready'; $state.stage = 'complete' }
    $state | ConvertTo-Json -Compress | Set-Content -LiteralPath "$reportPath.tmp" -Encoding UTF8
    [IO.File]::Move("$reportPath.tmp", $reportPath)
}
if ($failures.Count) { exit 1 }
