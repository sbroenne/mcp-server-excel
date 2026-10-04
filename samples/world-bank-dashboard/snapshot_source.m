let
    Source = Excel.CurrentWorkbook(){[Name="SourceObservations"]}[Content],
    Typed = Table.TransformColumnTypes(Source, {{"CountryCode", type text}, {"Year", Int64.Type},
        {"IndicatorCode", type text}, {"Metric", type text}, {"Value", type number}}, "en-US"),
    Keys = Table.Group(Typed, {"CountryCode", "Year", "IndicatorCode"},
        {{"Count", each Table.RowCount(_), Int64.Type}}),
    Checked = if Table.RowCount(Table.SelectRows(Keys, each [Count] <> 1)) > 0
        then error "Duplicate country/year/indicator observations."
        else if Table.RowCount(Typed) <> 5625
        then error "Expected 25 countries, 25 years and 9 indicators, including explicit missing observations."
        else Typed,
    Mode = Table.AddColumn(Checked, "SourceMode", each "Snapshot", type text),
    Refreshed = Table.AddColumn(Mode, "LoadedAt", each DateTime.FixedLocalNow(), type datetime)
in Refreshed
