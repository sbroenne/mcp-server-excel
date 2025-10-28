# Extract test sections into partial files

$testFile = "tests/ExcelMcp.Core.Tests/Integration/RangeCommandsTests.cs"
$outputDir = "tests/ExcelMcp.Core.Tests/Integration/Range"
$lines = Get-Content $testFile

# Section definitions (line numbers are 0-indexed)
$sections = @(
    @{Name="Values"; Start=73; End=151},      # Lines 74-152
    @{Name="Formulas"; Start=152; End=215},   # Lines 153-216
    @{Name="Editing"; Start=216; End=314},    # Lines 217-315 (Clear + Copy + Insert/Delete)
    @{Name="Search"; Start=372; End=461},     # Lines 373-462 (Find + Replace + Sort)
    @{Name="Discovery"; Start=462; End=530},  # Lines 463-531 (UsedRange + CurrentRegion + RangeInfo)
    @{Name="Hyperlinks"; Start=531; End=599}  # Lines 532-600 (Add + Remove + List + Get)
)

foreach ($section in $sections) {
    $name = $section.Name
    $start = $section.Start
    $end = $section.End

    $sectionLines = $lines[$start..$end]

    $content = @"
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Tests for range $($name.ToLower()) operations
/// </summary>
public partial class RangeCommandsTests
{
$($sectionLines -join "`n")
}
"@

    $outputFile = Join-Path $outputDir "RangeCommandsTests.$name.cs"
    Set-Content -Path $outputFile -Value $content -Encoding UTF8
    Write-Output "Created RangeCommandsTests.$name.cs ($($end - $start + 1) lines)"
}

Write-Output "`nDone! Created 6 partial test files."
