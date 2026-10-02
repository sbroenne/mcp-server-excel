param(
    [Parameter(Mandatory)][string]$Path,
    [string]$SourceRange = 'A1:E9',
    [switch]$IncludeAnalysis,
    [switch]$IncludePresentation,
    [switch]$UseUsedRange,
    [switch]$Recalculate,
    [string]$ProbeChangesJson
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
    param([object]$Range, [switch]$Formulas)
    $values = $Range.Value2
    if ($Formulas) { $values = $Range.Formula2 }
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
    if ($ProbeChangesJson) {
        if (-not $Recalculate) { throw 'A formula probe requires recalculation.' }
        $changes = ConvertFrom-Json -InputObject $ProbeChangesJson -NoEnumerate
        if ($changes -isnot [array] -or $changes.Count -eq 0 -or $changes.Count -gt 10) {
            throw 'A formula probe requires between one and ten numeric cell changes.'
        }
        $probeSheets = Track-Com $book.Worksheets
        foreach ($change in $changes) {
            if ($change.sheet -isnot [string] -or $change.cell -notmatch '^[A-Z]+[1-9][0-9]*$' -or
                ($change.value -isnot [long] -and $change.value -isnot [double])) {
                throw 'A formula probe requires a sheet, single cell, and numeric value.'
            }
            $probeSheet = Track-Com ($probeSheets.Item($change.sheet))
            $probeCell = Track-Com ($probeSheet.Range($change.cell))
            $probeCell.Value2 = [double]$change.value
        }
    }
    if ($Recalculate) { $excel.CalculateFull() }
    $sheets = Track-Com $book.Worksheets
    $sheetResults = @()
    for ($s = 1; $s -le $sheets.Count; $s++) {
        $sheet = Track-Com ($sheets.Item($s))
        if ($UseUsedRange) {
            $source = Track-Com $sheet.UsedRange
        } else {
            $source = Track-Com ($sheet.Range($SourceRange))
        }
        $sourceRows = Track-Com $source.Rows
        $sourceColumns = Track-Com $source.Columns
        if ($UseUsedRange -and ([long]$sourceRows.Count * $sourceColumns.Count) -gt 10000) {
            throw "InspectionLimitExceeded: Worksheet '$($sheet.Name)' exceeds the 10000-cell independent inspection limit."
        }
        $sourceCells = Track-Com $source.Cells
        $formats = @()
        $presentation = @()
        for ($r = 1; $r -le $sourceRows.Count; $r++) {
            $formatRow = @()
            $presentationRow = @()
            for ($c = 1; $c -le $sourceColumns.Count; $c++) {
                $cell = Track-Com ($sourceCells.Item($r, $c))
                $formatRow += $cell.NumberFormat
                if ($IncludePresentation) {
                    $font = Track-Com $cell.Font
                    $interior = Track-Com $cell.Interior
                    $presentationRow += @{
                        bold = [bool]$font.Bold
                        fontColor = [long]$font.Color
                        fillColor = [long]$interior.Color
                        text = [string]$cell.Text
                    }
                }
            }
            $formats += ,$formatRow
            if ($IncludePresentation) { $presentation += ,$presentationRow }
        }
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
            $pivotName = $null
            if ($IncludeAnalysis) {
                $layout = Track-Com $chart.PivotLayout
                if ($null -ne $layout) {
                    $linkedPivot = Track-Com $layout.PivotTable
                    $pivotName = $linkedPivot.Name
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
                pivot = $pivotName
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
            $connectionName = $null
            $queryName = $null
            if ($table.SourceType -in @(0, 3)) {
                $queryTable = Track-Com $table.QueryTable
                $connection = Track-Com $queryTable.WorkbookConnection
                $connectionName = $connection.Name
                $oleDb = Track-Com $connection.OLEDBConnection
                $connectionString = [string]$oleDb.Connection
                if ($connectionString -match '(?i)Microsoft\.Mashup\.OleDb\.1' -and
                    $connectionString -match '(?i)(?:^|;)\s*Location=([^;]+)') {
                    $queryName = $Matches[1].Trim('"')
                }
            }
            $tableRange = Track-Com $table.Range
            $tableStyle = Track-Com $table.TableStyle
            $tableResults += @{
                name = $table.Name; rows = $data; visibleRows = $visibleRows
                address = $tableRange.Address($false, $false)
                connection = $connectionName; style = $tableStyle.Name
                query = $queryName
            }
        }
        $pivots = Track-Com ($sheet.PivotTables())
        $pivotResults = @()
        for ($i = 1; $i -le $pivots.Count; $i++) {
            $pivot = Track-Com ($pivots.Item($i))
            $range = Track-Com $pivot.TableRange2
            $cache = Track-Com ($pivot.PivotCache())
            $pivotResults += @{
                name = $pivot.Name; values = (Read-Values $range)
                layout = [int]$pivot.LayoutRowDefault; olap = [bool]$cache.OLAP
            }
        }
        $sheetResults += @{
            name = $sheet.Name
            sourceRow = [int]$source.Row
            sourceColumn = [int]$source.Column
            sourceValues = (Read-Values $source)
            sourceFormulas = (Read-Values $source -Formulas)
            sourceFormats = $formats
            sourcePresentation = $presentation
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
    $queries = Track-Com $book.Queries
    $queryResults = @()
    for ($i = 1; $i -le $queries.Count; $i++) {
        $query = Track-Com ($queries.Item($i))
        $queryResults += @{ name = $query.Name; formula = $query.Formula }
    }
    $modelResult = $null
    if ($IncludeAnalysis) {
        $model = Track-Com $book.Model
        $modelTables = Track-Com $model.ModelTables
        $modelNames = @()
        for ($i = 1; $i -le $modelTables.Count; $i++) {
            $modelTable = Track-Com ($modelTables.Item($i))
            $modelNames += $modelTable.Name
        }
        $relationships = Track-Com $model.ModelRelationships
        $relationshipResults = @()
        for ($i = 1; $i -le $relationships.Count; $i++) {
            $relationship = Track-Com ($relationships.Item($i))
            $foreignTable = Track-Com $relationship.ForeignKeyTable
            $foreignColumn = Track-Com $relationship.ForeignKeyColumn
            $primaryTable = Track-Com $relationship.PrimaryKeyTable
            $primaryColumn = Track-Com $relationship.PrimaryKeyColumn
            $relationshipResults += @{
                fromTable = $foreignTable.Name; fromColumn = $foreignColumn.Name
                toTable = $primaryTable.Name; toColumn = $primaryColumn.Name
                active = $relationship.Active
            }
        }
        $measures = Track-Com $model.ModelMeasures
        $measureResults = @()
        for ($i = 1; $i -le $measures.Count; $i++) {
            $measure = Track-Com ($measures.Item($i))
            $measureResults += @{ name = $measure.Name; formula = $measure.Formula }
        }
        $modelResult = @{ tables = $modelNames; relationships = $relationshipResults; measures = $measureResults }
    }
    @{
        sheets = $sheetResults; slicers = $slicerResults; queries = $queryResults; model = $modelResult
        calculationMode = [int]$excel.Calculation
    } |
        ConvertTo-Json -Depth 15 -Compress
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
