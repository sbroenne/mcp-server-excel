<#
.SYNOPSIS
    Updates aggregate GitHub star history and generates its SVG chart.

.DESCRIPTION
    Reads date/count aggregates from a CSV file, optionally records a current
    daily snapshot, and writes a deterministic, theme-aware SVG. The script
    never reads or stores stargazer identities.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[^/]+/[^/]+$")]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$HistoryPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [int]$CurrentCount,

    [DateTimeOffset]$SnapshotDate = [DateTimeOffset]::UtcNow
)

$ErrorActionPreference = "Stop"

function ConvertTo-SvgText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-SvgNumber {
    param(
        [Parameter(Mandatory = $true)]
        [double]$Value
    )

    return $Value.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

$resolvedHistoryPath = [System.IO.Path]::GetFullPath($HistoryPath)

if (-not (Test-Path -LiteralPath $resolvedHistoryPath -PathType Leaf)) {
    throw "Star-history aggregate file '$resolvedHistoryPath' does not exist."
}

$csvText = [System.IO.File]::ReadAllText($resolvedHistoryPath).TrimStart([char]0xFEFF)
$csvLines = @($csvText -split "\r?\n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

if ($csvLines.Count -eq 0 -or $csvLines[0].Trim().ToLowerInvariant() -ne "date,count") {
    throw "Star-history aggregate file '$resolvedHistoryPath' must start with the header 'date,count'."
}

if ($csvLines.Count -eq 1) {
    throw "Star-history aggregate file '$resolvedHistoryPath' does not contain any aggregate rows."
}

$rawRows = @($csvLines | ConvertFrom-Csv)
$history = [System.Collections.Generic.List[object]]::new()
$previousDate = [DateTime]::MinValue

foreach ($row in $rawRows) {
    $date = [DateTime]::MinValue
    $parsedDate = [DateTime]::TryParseExact(
        [string]$row.date,
        "yyyy-MM-dd",
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$date)

    if (-not $parsedDate) {
        throw "Star-history aggregate row has invalid date '$($row.date)'; expected yyyy-MM-dd."
    }

    $count = 0
    $parsedCount = [int]::TryParse(
        [string]$row.count,
        [System.Globalization.NumberStyles]::None,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$count)

    if (-not $parsedCount -or $count -lt 0) {
        throw "Star-history aggregate row for '$($row.date)' has invalid count '$($row.count)'."
    }

    if ($history.Count -gt 0 -and $date -le $previousDate) {
        throw "Star-history aggregate dates must be strictly increasing; found '$($row.date)' after '$($previousDate.ToString("yyyy-MM-dd"))'."
    }

    $history.Add([pscustomobject]@{
        Date = $date
        Count = $count
    })
    $previousDate = $date
}

if ($PSBoundParameters.ContainsKey("CurrentCount")) {
    if ($CurrentCount -lt 0) {
        throw "Current star count cannot be negative."
    }

    $snapshotDay = $SnapshotDate.UtcDateTime.Date
    $latestDay = $history[-1].Date

    if ($snapshotDay -lt $latestDay) {
        throw "Snapshot date '$($snapshotDay.ToString("yyyy-MM-dd"))' precedes the latest history date '$($latestDay.ToString("yyyy-MM-dd"))'."
    }

    if ($snapshotDay -eq $latestDay) {
        $history[-1].Count = $CurrentCount
    }
    else {
        $history.Add([pscustomobject]@{
            Date = $snapshotDay
            Count = $CurrentCount
        })
    }

    $historyRows = $history | ForEach-Object {
        "$($_.Date.ToString("yyyy-MM-dd")),$($_.Count)"
    }
    $updatedCsv = "date,count`n$($historyRows -join "`n")`n"
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($resolvedHistoryPath, $updatedCsv, $utf8NoBom)
}

$firstSnapshot = $history[0]
$latestSnapshot = $history[-1]
$chartEnd = $latestSnapshot.Date

if ($chartEnd -eq $firstSnapshot.Date) {
    $chartEnd = $firstSnapshot.Date.AddDays(1)
}

$width = 900
$height = 480
$left = 72
$right = 24
$top = 76
$bottom = 62
$plotWidth = $width - $left - $right
$plotHeight = $height - $top - $bottom
$durationTicks = ($chartEnd - $firstSnapshot.Date).Ticks
$maxStars = ($history | Measure-Object -Property Count -Maximum).Maximum

if ($maxStars -le 0) {
    throw "Star-history aggregate file '$resolvedHistoryPath' must contain at least one positive count."
}

$points = foreach ($snapshot in $history) {
    $elapsedTicks = ($snapshot.Date - $firstSnapshot.Date).Ticks
    $x = $left + (($elapsedTicks / $durationTicks) * $plotWidth)
    $y = $top + $plotHeight - (($snapshot.Count / $maxStars) * $plotHeight)

    [pscustomobject]@{
        X = $x
        Y = $y
    }
}

$lineCoordinates = ($points | ForEach-Object {
    "$(ConvertTo-SvgNumber $_.X) $(ConvertTo-SvgNumber $_.Y)"
}) -join " L "
$linePath = "M $lineCoordinates"

$firstX = ConvertTo-SvgNumber $points[0].X
$lastX = ConvertTo-SvgNumber $points[-1].X
$baselineY = ConvertTo-SvgNumber ($top + $plotHeight)
$areaPath = "M $firstX $baselineY L $lineCoordinates L $lastX $baselineY Z"

$repositoryText = ConvertTo-SvgText $Repository
$dateRange = "{0:MMM yyyy} - {1:MMM yyyy}" -f $firstSnapshot.Date, $latestSnapshot.Date
$subtitle = ConvertTo-SvgText "$Repository - $($latestSnapshot.Count) stars - $dateRange"
$description = ConvertTo-SvgText (
    "Daily aggregate GitHub star counts for $Repository from " +
    "$($firstSnapshot.Date.ToString("yyyy-MM-dd")) to $($latestSnapshot.Date.ToString("yyyy-MM-dd")); " +
    "latest count $($latestSnapshot.Count).")

$svg = [System.Text.StringBuilder]::new()
[void]$svg.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$svg.AppendLine("<svg xmlns=`"http://www.w3.org/2000/svg`" width=`"$width`" height=`"$height`" viewBox=`"0 0 $width $height`" role=`"img`" aria-labelledby=`"title description`">")
[void]$svg.AppendLine("  <title id=`"title`">GitHub stars over time for $repositoryText</title>")
[void]$svg.AppendLine("  <desc id=`"description`">$description</desc>")
[void]$svg.AppendLine("  <style>")
[void]$svg.AppendLine("    .background { fill: #ffffff; }")
[void]$svg.AppendLine("    .grid { stroke: #d0d7de; stroke-width: 1; }")
[void]$svg.AppendLine("    .axis-text { fill: #57606a; font: 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .title { fill: #1f2328; font: 600 22px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .subtitle { fill: #57606a; font: 14px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }")
[void]$svg.AppendLine("    .area { fill: #2da44e; opacity: 0.14; }")
[void]$svg.AppendLine("    .line { fill: none; stroke: #1f883d; stroke-linecap: round; stroke-linejoin: round; stroke-width: 3; }")
[void]$svg.AppendLine("    @media (prefers-color-scheme: dark) {")
[void]$svg.AppendLine("      .background { fill: #0d1117; }")
[void]$svg.AppendLine("      .grid { stroke: #30363d; }")
[void]$svg.AppendLine("      .axis-text, .subtitle { fill: #8b949e; }")
[void]$svg.AppendLine("      .title { fill: #f0f6fc; }")
[void]$svg.AppendLine("      .area { fill: #3fb950; opacity: 0.18; }")
[void]$svg.AppendLine("      .line { stroke: #3fb950; }")
[void]$svg.AppendLine("    }")
[void]$svg.AppendLine("  </style>")
[void]$svg.AppendLine("  <rect class=`"background`" width=`"$width`" height=`"$height`" rx=`"8`" />")
[void]$svg.AppendLine("  <text class=`"title`" x=`"$left`" y=`"34`">GitHub stars over time</text>")
[void]$svg.AppendLine("  <text class=`"subtitle`" x=`"$left`" y=`"57`">$subtitle</text>")

for ($index = 0; $index -le 4; $index++) {
    $value = [Math]::Round(($maxStars * $index) / 4)
    $y = $top + $plotHeight - (($value / $maxStars) * $plotHeight)
    $yText = ConvertTo-SvgNumber $y

    [void]$svg.AppendLine("  <line class=`"grid`" x1=`"$left`" y1=`"$yText`" x2=`"$($left + $plotWidth)`" y2=`"$yText`" />")
    [void]$svg.AppendLine("  <text class=`"axis-text`" x=`"$($left - 12)`" y=`"$yText`" text-anchor=`"end`" dominant-baseline=`"middle`">$value</text>")
}

for ($index = 0; $index -le 4; $index++) {
    $x = $left + (($plotWidth * $index) / 4)
    $tickDate = $firstSnapshot.Date.AddTicks([long](($durationTicks * $index) / 4))
    $anchor = if ($index -eq 0) { "start" } elseif ($index -eq 4) { "end" } else { "middle" }

    [void]$svg.AppendLine("  <text class=`"axis-text`" x=`"$(ConvertTo-SvgNumber $x)`" y=`"$($top + $plotHeight + 28)`" text-anchor=`"$anchor`">$($tickDate.ToString("MMM yyyy"))</text>")
}

[void]$svg.AppendLine("  <path class=`"area`" d=`"$areaPath`" />")
[void]$svg.AppendLine("  <path class=`"line`" d=`"$linePath`" />")
[void]$svg.AppendLine("</svg>")

$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath

if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($resolvedOutputPath, $svg.ToString(), $utf8NoBom)

Write-Host "Generated $resolvedOutputPath from $($history.Count) aggregate snapshots; latest count $($latestSnapshot.Count)." -ForegroundColor Green
