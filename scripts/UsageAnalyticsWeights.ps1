<#
.SYNOPSIS
    Loads the hand-picked usage analytics work weights and builds matching query lookups.
.DESCRIPTION
    Dot-source this file. The weights file lists every MCP tool action with a work level
    (light, medium, or heavy) and maps excelcli command categories to MCP tool names so
    both entry points report the same action name.
#>

function Assert-UsageAnalyticsJsonHasUniqueKeys {
    param([System.Text.Json.JsonElement]$Element, [string]$Location)

    if ($Element.ValueKind -eq [System.Text.Json.JsonValueKind]::Object) {
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($property in $Element.EnumerateObject()) {
            if (-not $seen.Add($property.Name)) {
                throw "Usage analytics weights repeat '$Location/$($property.Name)'."
            }
            Assert-UsageAnalyticsJsonHasUniqueKeys $property.Value "$Location/$($property.Name)"
        }
    }
    elseif ($Element.ValueKind -eq [System.Text.Json.JsonValueKind]::Array) {
        foreach ($item in $Element.EnumerateArray()) {
            Assert-UsageAnalyticsJsonHasUniqueKeys $item $Location
        }
    }
}

function Read-UsageAnalyticsWeights {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolvedPath = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        throw "Usage analytics weights file '$resolvedPath' does not exist."
    }
    $text = [IO.File]::ReadAllText($resolvedPath)
    $document = [System.Text.Json.JsonDocument]::Parse($text)
    try {
        Assert-UsageAnalyticsJsonHasUniqueKeys $document.RootElement ""
    }
    finally {
        $document.Dispose()
    }
    $source = $text | ConvertFrom-Json -AsHashtable

    $levels = [ordered]@{}
    foreach ($level in @($source.levels.Keys)) {
        $value = $source.levels[$level]
        if ($level -notmatch '^[a-z]+$' -or
            -not ($value -is [int] -or $value -is [long]) -or
            $value -le 0) {
            throw "Usage analytics weights level '$level' must be a positive whole number."
        }
        $levels[$level] = [int]$value
    }
    if ($levels.Count -eq 0 -or
        @($levels.Values | Sort-Object -Unique).Count -ne $levels.Count) {
        throw "Usage analytics weights levels must be non-empty and distinct."
    }

    $features = @($source.features)
    if ($features.Count -eq 0 -or
        @($features | Sort-Object -Unique).Count -ne $features.Count -or
        $features -notcontains "other" -or
        @($features | Where-Object { $_ -notmatch '^[a-z0-9-]+$' }).Count -gt 0) {
        throw "Usage analytics weights features must be unique, safe, and include 'other'."
    }

    $weightByName = @{}
    $levelByName = @{}
    $featureByTool = @{}
    $cliOnlyTools = @{}
    foreach ($tool in @($source.tools.Keys)) {
        $entry = $source.tools[$tool]
        if ($tool -notmatch '^[a-z0-9_]+$' -or $tool.EndsWith("_read")) {
            throw "Usage analytics weights tool '$tool' is not a base MCP tool name."
        }
        if ($features -notcontains $entry.feature) {
            throw "Usage analytics weights tool '$tool' uses unknown feature '$($entry.feature)'."
        }
        if ($null -eq $entry.actions -or $entry.actions.Count -eq 0) {
            throw "Usage analytics weights tool '$tool' lists no actions."
        }
        $featureByTool[$tool] = $entry.feature
        if ($entry.cliOnly -eq $true) {
            $cliOnlyTools[$tool] = $true
        }
        foreach ($action in @($entry.actions.Keys)) {
            $level = $entry.actions[$action]
            if ($action -notmatch '^[a-z0-9-]+$') {
                throw "Usage analytics weights action '$tool/$action' is not a safe action name."
            }
            if (-not $levels.Contains([string]$level)) {
                throw "Usage analytics weights action '$tool/$action' uses unknown level '$level'."
            }
            $weightByName["$tool/$action"] = $levels[[string]$level]
            $levelByName["$tool/$action"] = [string]$level
        }
    }

    $excludedActions = @($source.excludedActions)
    foreach ($excluded in $excludedActions) {
        if (-not $weightByName.ContainsKey($excluded)) {
            throw "Usage analytics weights exclude unknown action '$excluded'."
        }
    }

    $toolMap = [ordered]@{}
    function Add-ToolMapping([string]$RawTool, [string]$Tool) {
        if ($toolMap.Contains($RawTool) -and $toolMap[$RawTool] -ne $Tool) {
            throw "Usage analytics weights map '$RawTool' to both '$($toolMap[$RawTool])' and '$Tool'."
        }
        $toolMap[$RawTool] = $Tool
    }
    foreach ($tool in ($featureByTool.Keys | Sort-Object)) {
        if (-not $cliOnlyTools.ContainsKey($tool)) {
            Add-ToolMapping $tool $tool
            Add-ToolMapping "${tool}_read" $tool
        }
    }

    $splitMap = [ordered]@{}
    foreach ($category in (@($source.cliCategories.Keys) | Sort-Object)) {
        $targets = @($source.cliCategories[$category])
        if ($category -notmatch '^[a-z0-9_]+$' -or $targets.Count -eq 0) {
            throw "Usage analytics weights CLI category '$category' is invalid."
        }
        foreach ($target in $targets) {
            if (-not $featureByTool.ContainsKey($target)) {
                throw "Usage analytics weights CLI category '$category' targets unknown tool '$target'."
            }
        }
        if ($targets.Count -eq 1) {
            Add-ToolMapping $category $targets[0]
            continue
        }
        foreach ($target in $targets) {
            foreach ($action in @($source.tools[$target].actions.Keys)) {
                $rawName = "$category/$action"
                if ($splitMap.Contains($rawName)) {
                    throw "Usage analytics weights CLI action '$rawName' matches more than one tool."
                }
                $splitMap[$rawName] = "$target/$action"
            }
        }
    }

    $heaviestLevel = ($levels.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    $heavyNames = @($levelByName.Keys | Where-Object { $levelByName[$_] -eq $heaviestLevel } | Sort-Object)

    return [pscustomobject]@{
        Levels = $levels
        HeaviestLevel = $heaviestLevel
        Features = $features
        ExcludedActions = $excludedActions
        WeightByName = $weightByName
        LevelByName = $levelByName
        FeatureByTool = $featureByTool
        ToolMap = $toolMap
        SplitMap = $splitMap
        HeavyNames = $heavyNames
    }
}

function Get-UsageAnalyticsFeature {
    param([Parameter(Mandatory = $true)]$Weights, [Parameter(Mandatory = $true)][string]$Name)
    $tool = $Name.Split('/')[0]
    if ($Weights.FeatureByTool.ContainsKey($tool)) {
        return $Weights.FeatureByTool[$tool]
    }
    return "other"
}

function New-UsageAnalyticsQueryPrelude {
    <#
    .SYNOPSIS
        Returns KQL let statements that normalize CLI and read-only MCP names to one action name.
    .DESCRIPTION
        Defines normalizedRequests and normalizedEvents with Name rewritten to the MCP base
        tool name, Feature set from the weights file, EntryPoint defaulted to mcp-server for
        rows recorded before the label existed, and excluded actions removed.
    #>
    param([Parameter(Mandatory = $true)]$Weights)

    $toolRows = ($Weights.ToolMap.GetEnumerator() | ForEach-Object {
        "    '$($_.Key)', '$($_.Value)', '$($Weights.FeatureByTool[$_.Value])'"
    }) -join ",`n"
    $splitRows = ($Weights.SplitMap.GetEnumerator() | ForEach-Object {
        $tool = $_.Value.Split('/')[0]
        "    '$($_.Key)', '$($_.Value)', '$($Weights.FeatureByTool[$tool])'"
    }) -join ",`n"
    $excluded = ($Weights.ExcludedActions | ForEach-Object { "'$_'" }) -join ", "
    $heavy = ($Weights.HeavyNames | ForEach-Object { "'$_'" }) -join ", "
    $normalize = @'
| extend MapRawTool=tostring(split(Name, '/')[0]),
         MapAction=tostring(split(Name, '/')[1])
| lookup kind=leftouter toolMap on MapRawTool
| lookup kind=leftouter splitMap on $left.Name == $right.MapRawName
| extend Name=case(
             isnotempty(MapSplitName), MapSplitName,
             isnotempty(MapTool), strcat(MapTool, '/', MapAction),
             Name),
         Feature=coalesce(MapSplitFeature, MapFeature, 'other'),
         EntryPoint=coalesce(tostring(Properties['EntryPoint']), 'mcp-server')
| project-away Map*
| where Name !in (excludedNames)
'@

    return @"
let toolMap = datatable(MapRawTool:string, MapTool:string, MapFeature:string)[
$toolRows
];
let splitMap = datatable(MapRawName:string, MapSplitName:string, MapSplitFeature:string)[
$splitRows
];
let excludedNames = dynamic([$excluded]);
let heavyNames = dynamic([$heavy]);
let normalizedRequests = AppRequests
$normalize;
let normalizedEvents = AppEvents
| where Name != 'SessionStart'
$normalize;

"@
}
