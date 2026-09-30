param(
    [Parameter(Mandatory)][string]$Path,
    [string]$SourceRange = 'A1:E9'
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'Stop'
$references = [System.Collections.Generic.List[object]]::new()
$excel = $null
$book = $null

function Track-Com {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $references.Add($Value)
    }
    return ,$Value
}

function Read-Values {
    param([object]$Range)
    $values = $Range.Value2
    $rows = [System.Collections.Generic.List[object]]::new()
    if ($null -ne $values) {
        if ($values -is [Array] -and $values.Rank -eq 2) {
            for ($r = $values.GetLowerBound(0); $r -le $values.GetUpperBound(0); $r++) {
                $row = [System.Collections.Generic.List[object]]::new()
                for ($c = $values.GetLowerBound(1); $c -le $values.GetUpperBound(1); $c++) {
                    $row.Add($values[$r, $c])
                }
                $rows.Add($row.ToArray())
            }
        } else {
            $rows.Add(@($values))
        }
    }
    return ,$rows.ToArray()
}

try {
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $excel = Track-Com (New-Object -ComObject Excel.Application)
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $books = Track-Com $excel.Workbooks
    $book = Track-Com ($books.Open($resolved, 0, $true))
    $sheets = Track-Com $book.Worksheets
    $sheetResults = @()
    for ($s = 1; $s -le $sheets.Count; $s++) {
        $sheet = Track-Com ($sheets.Item($s))
        $source = Track-Com ($sheet.Range($SourceRange))
        $positions = @{}
        foreach ($address in @('D2', 'E2', 'F2', 'G2', 'H2')) {
            $cell = Track-Com ($sheet.Range($address))
            $positions[$address] = @{ left = $cell.Left; top = $cell.Top }
        }
        $charts = Track-Com ($sheet.ChartObjects())
        $chartResults = @()
        for ($i = 1; $i -le $charts.Count; $i++) {
            $object = Track-Com ($charts.Item($i))
            $chart = Track-Com $object.Chart
            $seriesCollection = Track-Com ($chart.SeriesCollection())
            $seriesResults = @()
            for ($j = 1; $j -le $seriesCollection.Count; $j++) {
                $series = Track-Com ($seriesCollection.Item($j))
                $seriesResults += @{
                    name = $series.Name
                    values = @($series.Values)
                    categories = @($series.XValues)
                }
            }
            $chartResults += @{
                name = $object.Name
                type = [int]$chart.ChartType
                left = $object.Left
                top = $object.Top
                width = $object.Width
                height = $object.Height
                series = $seriesResults
            }
        }
        $tables = Track-Com $sheet.ListObjects
        $tableResults = @()
        for ($i = 1; $i -le $tables.Count; $i++) {
            $table = Track-Com ($tables.Item($i))
            $body = Track-Com $table.DataBodyRange
            $visibleRows = @()
            $data = @()
            if ($null -ne $body) {
                $data = Read-Values $body
                $rows = Track-Com $body.Rows
                for ($r = 1; $r -le $rows.Count; $r++) {
                    $row = Track-Com ($rows.Item($r))
                    $entireRow = Track-Com $row.EntireRow
                    if (-not $entireRow.Hidden) { $visibleRows += ,$data[$r - 1] }
                }
            }
            $tableResults += @{ name = $table.Name; rows = $data; visibleRows = $visibleRows }
        }
        $pivots = Track-Com ($sheet.PivotTables())
        $pivotResults = @()
        for ($i = 1; $i -le $pivots.Count; $i++) {
            $pivot = Track-Com ($pivots.Item($i))
            $range = Track-Com $pivot.TableRange2
            $pivotResults += @{ name = $pivot.Name; values = (Read-Values $range) }
        }
        $sheetResults += @{
            name = $sheet.Name
            sourceValues = (Read-Values $source)
            bounds = @{ left = $source.Left; top = $source.Top; width = $source.Width; height = $source.Height }
            positions = $positions
            charts = $chartResults
            tables = $tableResults
            pivots = $pivotResults
        }
    }
    $caches = Track-Com $book.SlicerCaches
    $slicerResults = @()
    for ($i = 1; $i -le $caches.Count; $i++) {
        $cache = Track-Com ($caches.Item($i))
        $items = Track-Com $cache.SlicerItems
        $selected = @()
        for ($j = 1; $j -le $items.Count; $j++) {
            $item = Track-Com ($items.Item($j))
            if ($item.Selected) { $selected += $item.Name }
        }
        $linkedPivots = Track-Com $cache.PivotTables
        $pivotNames = @()
        for ($j = 1; $j -le $linkedPivots.Count; $j++) {
            $pivot = Track-Com ($linkedPivots.Item($j))
            $pivotNames += $pivot.Name
        }
        $tableName = $null
        if ($linkedPivots.Count -eq 0) {
            $table = Track-Com $cache.ListObject
            if ($null -ne $table) { $tableName = $table.Name }
        }
        $slicers = Track-Com $cache.Slicers
        for ($j = 1; $j -le $slicers.Count; $j++) {
            $slicer = Track-Com ($slicers.Item($j))
            $shape = Track-Com $slicer.Shape
            $anchor = Track-Com $shape.TopLeftCell
            $sheet = Track-Com $shape.Parent
            $slicerResults += @{
                name = $slicer.Name
                field = $cache.SourceName
                sheet = $sheet.Name
                position = $anchor.Address($true, $true)
                left = $shape.Left
                top = $shape.Top
                selected = $selected
                pivots = $pivotNames
                table = $tableName
            }
        }
    }
    @{ sheets = $sheetResults; slicers = $slicerResults } | ConvertTo-Json -Depth 15 -Compress
}
finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
    }
    finally {
        try {
            if ($null -ne $excel) { $excel.Quit() }
        }
        finally {
            for ($i = $references.Count - 1; $i -ge 0; $i--) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($references[$i])
            }
        }
    }
}
