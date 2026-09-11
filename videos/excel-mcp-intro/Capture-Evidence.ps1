param([string]$Workbook = "$PSScriptRoot\assets\sales-workbook.xlsx")
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ExcelCaptureNative {
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
}
'@
$objects = [System.Collections.Generic.List[object]]::new()
function Keep($value) { $objects.Add($value); return ,$value }
$excel = $null
$book = $null
try {
    $excel = Keep (New-Object -ComObject Excel.Application)
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $books = Keep $excel.Workbooks
    $book = Keep ($books.Open($Workbook, 0, $false))
    $sheets = Keep $book.Worksheets
    $window = Keep $excel.ActiveWindow
    $excel.WindowState = -4143
    [ExcelCaptureNative]::MoveWindow([IntPtr]$excel.Hwnd, 0, 0, 1700, 1080, $true) | Out-Null
    foreach ($shot in @(@('Sales','sales-data.png',110), @('Report','sales-report.png',80), @('CleanSales','sales-query.png',110))) {
        $sheet = Keep ($sheets.Item($shot[0]))
        $sheet.Activate()
        if ($shot[0] -eq 'Report') {
            $column = Keep ($sheet.Range('B:B'))
            $column.ColumnWidth = 24
        }
        $window.Zoom = [int]$shot[2]
        $window.DisplayGridlines = $false
        $cell = Keep ($sheet.Range('A1'))
        $cell.Select()
        Start-Sleep -Milliseconds 900
        $bitmap = [System.Drawing.Bitmap]::new(1700,1080)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $hdc = $graphics.GetHdc()
        try {
            if (-not [ExcelCaptureNative]::PrintWindow([IntPtr]$excel.Hwnd,$hdc,2)) {
                throw 'Excel window capture failed.'
            }
        } finally { $graphics.ReleaseHdc($hdc); $graphics.Dispose() }
        # Crop the title/account chrome; remaining pixels are unaltered Excel UI.
        $crop = $bitmap.Clone([System.Drawing.Rectangle]::new(8,80,1684,946),$bitmap.PixelFormat)
        try { $crop.Save("$PSScriptRoot\assets\$($shot[1])",[System.Drawing.Imaging.ImageFormat]::Png) }
        finally { $crop.Dispose(); $bitmap.Dispose() }
    }
    $report = Keep ($sheets.Item('Report'))
    $total = Keep ($report.Range('B3'))
    if ([double]$total.Value2 -ne 584000) { throw 'Persisted revenue differs from expected total.' }
    $book.Save()
    Write-Output 'Captured Sales, Report, CleanSales. Persisted formula total verified: 584000.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    for ($i=$objects.Count-1; $i -ge 0; $i--) {
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($objects[$i])) {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($objects[$i]) | Out-Null
        }
    }
}
