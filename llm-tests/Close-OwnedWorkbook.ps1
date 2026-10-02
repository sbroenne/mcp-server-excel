param(
    [Parameter(Mandatory)][string]$Path,
    [switch]$InspectOnly
)

$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class OwnedWorkbookRot {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx context);
    public static object Find(string path) {
        IRunningObjectTable table = null;
        IBindCtx context = null;
        IEnumMoniker enumerator = null;
        try {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out table));
            Marshal.ThrowExceptionForHR(CreateBindCtx(0, out context));
            table.EnumRunning(out enumerator);
            var monikers = new IMoniker[1];
            while (enumerator.Next(1, monikers, IntPtr.Zero) == 0) {
                var moniker = monikers[0];
                try {
                    string name;
                    moniker.GetDisplayName(context, null, out name);
                    if (String.Equals(name.TrimStart('!'), path, StringComparison.OrdinalIgnoreCase)) {
                        object book;
                        Marshal.ThrowExceptionForHR(table.GetObject(moniker, out book));
                        return book;
                    }
                } finally { Marshal.FinalReleaseComObject(moniker); }
            }
            return null;
        } finally {
            if (enumerator != null) Marshal.FinalReleaseComObject(enumerator);
            if (context != null) Marshal.FinalReleaseComObject(context);
            if (table != null) Marshal.FinalReleaseComObject(table);
        }
    }
}
'@

$resolved = [IO.Path]::GetFullPath($Path)
$book = $null
$excel = $null
$books = $null
$matched = 0
$closed = 0
try {
    $book = [OwnedWorkbookRot]::Find($resolved)
    if ($null -ne $book) {
        if (-not [string]::Equals($book.FullName, $resolved, [StringComparison]::OrdinalIgnoreCase)) {
            throw 'Running workbook identity does not match the owned fixture.'
        }
        $matched++
        if (-not $InspectOnly) {
            $excel = $book.Application
            $books = $excel.Workbooks
            $soleWorkbook = $books.Count -eq 1
            $book.Close($false)
            $closed++
            if ($soleWorkbook) { $excel.Quit() }
        }
    }
    @{ matched = $matched; closed = $closed } | ConvertTo-Json -Compress
} finally {
    foreach ($owned in @($books, $excel, $book)) {
        if ($null -ne $owned -and [Runtime.InteropServices.Marshal]::IsComObject($owned)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned)
        }
    }
}
