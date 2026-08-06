// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Reports the architecture of a specific Windows process without inferring it from the caller.
/// </summary>
public static class ProcessArchitectureDetector
{
    private const uint ProcessQueryLimitedInformation = 0x1000;
    private const ushort ImageFileMachineUnknown = 0x0000;
    private const ushort ImageFileMachineI386 = 0x014c;
    private const ushort ImageFileMachineAmd64 = 0x8664;
    private const ushort ImageFileMachineArm64 = 0xaa64;

    /// <summary>
    /// Returns x86, x64, arm64, or unknown for the target process.
    /// </summary>
    public static string GetBitness(int? processId)
    {
        if (!OperatingSystem.IsWindows() || processId is null or <= 0)
        {
            return "unknown";
        }

        var processHandle = OpenProcess(ProcessQueryLimitedInformation, inheritHandle: false, processId.Value);
        if (processHandle == IntPtr.Zero)
        {
            return "unknown";
        }

        try
        {
            try
            {
                if (IsWow64Process2(processHandle, out ushort processMachine, out ushort nativeMachine))
                {
                    return MachineToBitness(processMachine == ImageFileMachineUnknown ? nativeMachine : processMachine);
                }
            }
            catch (EntryPointNotFoundException)
            {
                // Older supported Windows versions expose only IsWow64Process.
            }

            if (!IsWow64Process(processHandle, out bool isWow64))
            {
                return "unknown";
            }

            if (isWow64)
            {
                return "x86";
            }

            return RuntimeInformation.OSArchitecture switch
            {
                Architecture.X64 => "x64",
                Architecture.X86 => "x86",
                Architecture.Arm64 => "arm64",
                _ => "unknown"
            };
        }
        finally
        {
            _ = CloseHandle(processHandle);
        }
    }

    private static string MachineToBitness(ushort machine) => machine switch
    {
        ImageFileMachineI386 => "x86",
        ImageFileMachineAmd64 => "x64",
        ImageFileMachineArm64 => "arm64",
        _ => "unknown"
    };

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint desiredAccess, bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWow64Process2(
        IntPtr processHandle,
        out ushort processMachine,
        out ushort nativeMachine);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWow64Process(IntPtr processHandle, [MarshalAs(UnmanagedType.Bool)] out bool isWow64);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);
}
