import { execFile } from 'node:child_process';
import { constants } from 'node:fs';
import { access } from 'node:fs/promises';
import { promisify } from 'node:util';

const run = promisify(execFile);
const registrationCheck = `
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'7.4') { exit 3 }
$key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey('Excel.Application\\CLSID')
if ($null -eq $key) { exit 2 }
try {
    $clsid = [guid]::Empty
    if (-not [guid]::TryParse([string]$key.GetValue(''), [ref]$clsid)) { exit 2 }
} finally { $key.Dispose() }
`;

export class LaunchSetupError extends Error {}

function checkCancellation(signal: AbortSignal) {
	if (signal.aborted) {
		throw new DOMException('ExcelMcp startup canceled.', 'AbortError');
	}
}

export async function checkLaunchPrerequisites(executable: string, signal: AbortSignal) {
	checkCancellation(signal);
	if (process.platform !== 'win32') {
		throw new LaunchSetupError('ExcelMcp requires a local Windows desktop with Microsoft Excel.');
	}
	try {
		await access(executable, constants.R_OK);
	} catch {
		throw new LaunchSetupError('The bundled ExcelMcp server is missing or unreadable. Check access permissions or reinstall the extension.');
	}
	checkCancellation(signal);

	try {
		await run('pwsh', ['-NoProfile', '-NonInteractive', '-Command', registrationCheck], {
			timeout: 10000,
			windowsHide: true,
			maxBuffer: 65536,
			signal
		});
	} catch (error) {
		checkCancellation(signal);
		if (error instanceof Error && 'code' in error && (error.code === 'ENOENT' || error.code === 3)) {
			throw new LaunchSetupError('Install PowerShell 7.4 or later and make pwsh available on PATH, then restart VS Code.');
		}
		if (error instanceof Error && 'code' in error && error.code === 2) {
			throw new LaunchSetupError('Microsoft Excel is not registered. Install or repair desktop Excel, open it once, and retry.');
		}
		if (error instanceof Error && 'killed' in error && error.killed === true) {
			throw new LaunchSetupError('Checking Excel registration timed out. Verify PowerShell 7.4 or later and desktop Excel work, then retry.');
		}
		throw new LaunchSetupError('Could not check Excel registration. Verify PowerShell 7.4 or later works, then retry or repair Microsoft Office.');
	}
	checkCancellation(signal);
}
