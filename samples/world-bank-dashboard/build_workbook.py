"""Build the sample through excelcli and installed Excel, not an XLSX library."""

import argparse
import csv
import json
from pathlib import Path

from excel_bridge import Excel

ROOT = Path(__file__).parent
NAVY = "#0B1426"
PANEL = "#15233A"
WHITE = "#F1F5FB"
MUTED = "#A3B4CB"
TEAL = "#41D9C3"
BLUE = "#5EA9FF"
GOLD = "#F3C969"
COLORS = [TEAL, BLUE, GOLD, "#ED8C96", "#B19CFF", "#81C784"]
REGION_COLORS = {
    "East Asia & Pacific": TEAL,
    "Europe & Central Asia": BLUE,
    "Latin America & Caribbean": GOLD,
    "Middle East & North Africa": "#F7A77A",
    "North America": "#C6A6FF",
    "South Asia": "#F58BAC",
    "Sub-Saharan Africa": "#A8CE7A",
}
AGENT_LIVE_REFRESH = (
    "Ask your ExcelMCP-connected agent to switch Observations from EmbeddedSnapshot "
    "to WorldBankLive, refresh it, then refresh the PivotTables and verify the loaded "
    "results. This downloads about 283 MB. The status below reports the last successful load."
)
AGENT_OFFLINE = (
    "Ask your agent to keep EmbeddedSnapshot for offline use, or restore that source "
    "and refresh Observations and the PivotTables. Follow Excel privacy protections."
)


def cmd(command, **args):
    return {"command": command, "args": args}


def col(number):
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def read_data(name):
    with (ROOT / "data" / name).open(encoding="utf-8", newline="") as source:
        rows = list(csv.reader(source))
    return rows


def write(sheet, address, values):
    return cmd("range.set-values", sheetName=sheet, rangeAddress=address, values=values, overwritePolicy="allow")


def formulas(sheet, address, values):
    return cmd("range.set-formulas", sheetName=sheet, rangeAddress=address, formulas=values, overwritePolicy="allow")


def style(sheet, addresses, **options):
    return cmd("rangeformat.format", sheetName=sheet, rangeAddresses=addresses, formatOptions=options)


def table_commands(sheet, name, rows):
    address = f"A1:{col(len(rows[0]))}{len(rows)}"
    return [
        write(sheet, address, rows),
        cmd("table.create", sheetName=sheet, tableName=name, rangeAddress=address, tableStyle="TableStyleMedium2"),
        cmd("rangeformat.auto-fit-columns", sheetName=sheet, rangeAddress=f"A:{col(len(rows[0]))}"),
        cmd("window.freeze-panes", sheetName=sheet, frozenRows=1),
    ]


def build_model(excel):
    sheets = ["World Overview", "Growth & Resilience", "Prosperity & Progress",
              "Start Here", "Snapshot", "Countries", "Indicators", "Years",
              "Model Data", "Pivot Snapshot", "Pivot Growth", "Pivot Progress", "Chart Data"]
    existing = excel.call("sheet.list")
    print("Initial sheets:", json.dumps(existing)[:600], flush=True)
    # A new workbook starts with Sheet1 in this installation; discover its actual name.
    payload = existing.get("result", existing)
    if isinstance(payload, str):
        payload = json.loads(payload)
    first = payload["worksheets"][0]["name"]
    excel.batch([cmd("sheet.rename", oldName=first, newName=sheets[0])] +
                [cmd("sheet.create", sheetName=name) for name in sheets[1:]])
    observations = read_data("observations.csv")
    for row in observations[1:]:
        row[1] = int(row[1])
        row[4] = float(row[4]) if row[4] else None
    excel.batch(table_commands("Snapshot", "SourceObservations", observations))
    countries = read_data("countries.csv")
    indicators = read_data("indicators.csv")
    excel.batch(table_commands("Countries", "Countries", countries) +
                table_commands("Indicators", "Indicators", [row[:5] for row in indicators]) +
                table_commands("Years", "Years", [["Year"]] + [[year] for year in range(2000, 2025)]))
    query = (ROOT / "snapshot_source.m").read_text(encoding="utf-8")
    excel.call("powerquery.evaluate", mCode="Table.FirstN((" + query + "), 3)")
    excel.call("powerquery.create", queryName="Observations", mCode=query,
               loadDestination="both", targetSheet="Model Data", targetCellAddress="A1")
    excel.batch([cmd("table.add-to-data-model", tableName=name) for name in ["Countries", "Indicators", "Years"]])
    excel.batch([
        cmd("datamodelrel.create-relationship", fromTable="Observations", fromColumn=source,
            toTable=target, toColumn=key)
        for source, target, key in [("CountryCode", "Countries", "Code"),
                                   ("IndicatorCode", "Indicators", "Code"), ("Year", "Years", "Year")]
    ])
    measures = {}
    for metric in ("Income", "GDP", "Growth", "Inflation", "Population", "Life", "Electricity", "Internet", "Trade"):
        aggregation = "SUM" if metric in ("GDP", "Population") else "AVERAGE"
        measures[metric] = f'IF(HASONEVALUE(Years[Year]), CALCULATE({aggregation}(Observations[Value]), Indicators[Metric] = "{metric}"), BLANK())'
    measures.update({
        "Population millions": "DIVIDE([Population], 1000000)",
        "GDP index 2000": "DIVIDE([GDP], CALCULATE([GDP], ALL(Years), Years[Year] = 2000)) * 100",
        "Median income": "MEDIANX(VALUES(Countries[Code]), [Income])",
        "Median growth": "MEDIANX(VALUES(Countries[Code]), [Growth])",
        "Median life": "MEDIANX(VALUES(Countries[Code]), [Life])",
        "Selected year": "IF(HASONEVALUE(Years[Year]), MAX(Years[Year]), BLANK())",
        "Selected countries": "DISTINCTCOUNT(Countries[Code])",
        "Available observations": "COUNT(Observations[Value])",
    })
    excel.batch([
        cmd("datamodel.create-measure", tableName="Observations", measureName=name,
            daxFormula=formula, formatType="Decimal",
            description="Sample calculation; definitions and aggregation limits are on Start Here.")
        for name, formula in measures.items()
    ])
    print("Model built", flush=True)


def pivot(excel, name, sheet, rows, columns, measures):
    excel.call("pivottable.create-from-datamodel", tableName="Observations",
               destinationSheet=sheet, destinationCell="A1", pivotTableName=name)
    operations = []
    for field in rows:
        operations.append(cmd("pivottablefield.add-row-field", pivotTableName=name, fieldName=field))
    for field in columns:
        operations.append(cmd("pivottablefield.add-column-field", pivotTableName=name, fieldName=field))
    for measure in measures:
        operations.append(cmd("pivottablefield.add-value-field", pivotTableName=name, fieldName=f"[Measures].[{measure}]"))
    operations.extend([
        cmd("pivottablecalc.set-grand-totals", pivotTableName=name, showRowGrandTotals=False, showColumnGrandTotals=False),
        cmd("pivottablecalc.set-layout-options", pivotTableName=name, layoutOptions={"rowLayout": 1, "repeatLabels": True, "styleName": "PivotStyleMedium9"}),
        cmd("pivottable.refresh", pivotTableName=name),
    ])
    excel.batch(operations)


def slicer(excel, name, pivot_name, field, sheet, cell, left, top, width, height, columns=1, selected=None):
    # Excel's native OLAP slicer creation can crash on an inactive destination.
    excel.call("window.set-zoom", sheetName=sheet, zoom=100)
    excel.call("slicer.create-slicer", pivotTableName=pivot_name, fieldName=field,
               slicerName=name, destinationSheet=sheet, position=cell)
    excel.call("slicer.update-slicer", slicerName=name, slicerOptions={
        "left": left, "top": top, "width": width, "height": height,
        "columnCount": columns, "style": "SlicerStyleDark2",
    })
    if selected is not None:
        excel.call("slicer.set-slicer-selection", slicerName=name, selectedItems=selected)


def build_pivots(excel):
    pivot(excel, "SnapshotPivot", "Pivot Snapshot", ["[Countries].[Country]"], [],
          ["Income", "Life", "Population millions", "Growth", "Electricity", "Internet"])
    excel.batch([
        cmd("pivottablefield.add-filter-field", pivotTableName="SnapshotPivot", fieldName="[Years].[Year]"),
        cmd("pivottablefield.set-field-filter", pivotTableName="SnapshotPivot",
            fieldName="[Years].[Year]", selectedValues=["2024"]),
        cmd("pivottable.refresh", pivotTableName="SnapshotPivot"),
    ])
    slicer(excel, "OverviewYear", "SnapshotPivot", "[Years].[Year]", "World Overview", "AC8",
           810, 120, 240, 130, columns=5, selected=["2024"])
    slicer(excel, "OverviewRegion", "SnapshotPivot", "[Countries].[Region]", "World Overview", "AC17",
           810, 270, 240, 170)
    pivot(excel, "GrowthPivot", "Pivot Growth", ["[Years].[Year]"], ["[Countries].[Country]"], ["GDP index 2000"])
    slicer(excel, "GrowthCountries", "GrowthPivot", "[Countries].[Country]", "Growth & Resilience", "AC8",
           810, 120, 240, 320, columns=2,
           selected=["United States", "China", "Germany", "India", "Brazil", "South Africa"])
    pivot(excel, "ProgressPivot", "Pivot Progress", ["[Years].[Year]"], ["[Countries].[Country]"], ["Internet"])
    slicer(excel, "ProgressCountries", "ProgressPivot", "[Countries].[Country]", "Prosperity & Progress", "AC8",
           810, 120, 240, 320, columns=2,
           selected=["United States", "China", "Germany", "India", "Brazil", "South Africa"])
    for name in ("SnapshotPivot", "GrowthPivot", "ProgressPivot"):
        print(name, json.dumps(excel.call("pivottablecalc.get-data", pivotTableName=name))[:1600], flush=True)


def panel(excel, sheet, title, subtitle, page):
    excel.batch([
        cmd("rangeformat.set-column-width", sheetName=sheet, rangeAddress="A:AN", columnWidth=5),
        cmd("rangeformat.set-row-height", sheetName=sheet, rangeAddress="1:39", rowHeight=17),
        style(sheet, ["A1:AN39"], fillColor=NAVY, fontName="Aptos", fontSize=10, fontColor=WHITE),
        style(sheet, ["B2:AL3"], fontSize=27, bold=True),
        write(sheet, "B2", [[title]]),
        write(sheet, "B4", [[subtitle]]),
        style(sheet, ["B4:AL4"], fontSize=10, fontColor=MUTED),
        write(sheet, "B38", [["WORLD IN MOTION  /  Excel MCP  /  World Bank WDI · CC BY 4.0"]]),
        style(sheet, ["B38:AM38"], fontSize=9, fontColor=MUTED),
        cmd("window.set-display-options", sheetName=sheet, showGridlines=False, showHeadings=False),
        cmd("window.set-zoom", sheetName=sheet, zoom=90),
    ])
    links = [
        ("B36", "World Overview", "01   WORLD OVERVIEW"),
        ("K36", "Growth & Resilience", "02   GROWTH & RESILIENCE"),
        ("U36", "Prosperity & Progress", "03   PROSPERITY & PROGRESS"),
        ("AF36", "Start Here", "HOW TO USE"),
    ]
    excel.batch([
        formulas(sheet, cell, [[f'=HYPERLINK("#\'{target}\'!A1","{caption}")']])
        for cell, target, caption in links
    ] + [style(sheet, ["B36:AM36"], fontColor=TEAL, bold=True, fontSize=9)])


def kpi(sheet, left, caption, formula, number_format, accent=TEAL):
    right = col(ord(left[0]) - 64 + 6) if len(left) == 1 else "AM"
    # Each KPI value has its own wide merged cell; source data never uses merges.
    return [
        write(sheet, f"{left}7", [[caption]]),
        style(sheet, [f"{left}7:{right}7"], fontColor=MUTED, fontSize=9, bold=True),
        cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress=f"{left}8:{right}10"),
        formulas(sheet, f"{left}8", [[formula]]),
        style(sheet, [f"{left}8:{right}10"], fontColor=accent, fontSize=29, bold=True,
              numberFormat=number_format, verticalAlignment="center"),
    ]


def chart_style(excel, name, title, value_format=None, colors=COLORS):
    commands = [
        cmd("chartconfig.set-title", chartName=name, title=title),
        cmd("chartconfig.set-style", chartName=name, styleId=10),
        cmd("chartconfig.set-area-format", chartName=name, area="Chart", fillColor="#FFFFFF", lineColor="#FFFFFF"),
        cmd("chartconfig.set-area-format", chartName=name, area="Plot", fillColor="#FFFFFF", lineColor="#FFFFFF"),
        cmd("chartconfig.set-placement", chartName=name, placement=3, roundedCorners=True),
        cmd("chartconfig.show-legend", chartName=name, visible=True, legendPosition="Bottom"),
    ]
    if value_format:
        commands.append(cmd("chartconfig.set-axis-number-format", chartName=name, axis="Value", numberFormat=value_format))
    excel.batch(commands)
    info = excel.call("chart.read", chartName=name)["result"]
    count = len(info.get("series", []))
    if not count:
        print("Chart result:", json.dumps(info)[:1000], flush=True)
        count = info.get("seriesCount", 0)
    excel.batch([
        cmd("chartconfig.set-series-format", chartName=name, seriesIndex=index + 1,
            lineColor=colors[index % len(colors)], lineWeight=2.5,
            fillColor=colors[index % len(colors)])
        for index in range(count)
    ]) if count else None


def build_dashboards(excel):
    panel(excel, "World Overview", "THE WORLD IN MOTION",
          "Income, longevity and scale  /  25 selected economies  /  Observation years 2000-2024", 1)
    panel(excel, "Growth & Resilience", "GROWTH IS NOT A STRAIGHT LINE",
          "Real economic output  /  Each country's GDP indexed to 100 in 2000  /  Not stock-market returns", 2)
    panel(excel, "Prosperity & Progress", "PROGRESS YOU CAN MEASURE",
          "Digital access and longer lives  /  Country comparisons, not a causal model", 3)
    excel.batch([
        write("Chart Data", "A1:D1", [["Country", "Income per person", "Life expectancy", "Population millions"]]),
        formulas("Chart Data", "A2:D26", [
            [f'=IF(\'Pivot Snapshot\'!A{row}="","",\'Pivot Snapshot\'!A{row})'] +
            [f'=IF(AND($A{row-2}<>"",ISNUMBER(\'Pivot Snapshot\'!{column}{row})),\'Pivot Snapshot\'!{column}{row},NA())'
             for column in "BCD"] for row in range(4, 29)
        ]),
        style("Chart Data", ["B2:B26"], numberFormat="#,##0"),
        style("Chart Data", ["C2:D26"], numberFormat="0.0"),
        cmd("rangeformat.auto-fit-columns", sheetName="Chart Data", rangeAddress="A:D"),
    ])
    excel.batch(
        kpi("World Overview", "B", "OBSERVATION YEAR", '=IF(ISNUMBER(\'Pivot Snapshot\'!B1*1),\'Pivot Snapshot\'!B1*1,"Choose one year")', "0") +
        kpi("World Overview", "K", "MEDIAN COUNTRY INCOME", '=IF(COUNT(\'Pivot Snapshot\'!B4:B28)=0,"Select one year",MEDIAN(\'Pivot Snapshot\'!B4:B28))', '#,##0" intl $"', BLUE) +
        kpi("World Overview", "T", "PEOPLE IN SELECTION", '=IF(COUNT(\'Pivot Snapshot\'!D4:D28)=0,"Select one year",SUM(\'Pivot Snapshot\'!D4:D28)/1000)', '0.00" bn"', GOLD)
    )
    excel.call("chart.create-from-range", sheetName="World Overview",
               sourceRangeAddress="'Chart Data'!B1:D26", chartType="Bubble",
               chartName="IncomeAndLongevity", left=24, top=185, width=760, height=345)
    chart_style(excel, "IncomeAndLongevity", "Income and longevity · bubble area represents population", "0")
    excel.batch([
        cmd("chartconfig.show-legend", chartName="IncomeAndLongevity", visible=False),
        cmd("chartconfig.set-axis-title", chartName="IncomeAndLongevity", axis="Category",
            title="GDP per person · constant 2021 international $ (PPP)"),
        cmd("chartconfig.set-axis-title", chartName="IncomeAndLongevity", axis="Value", title="Life expectancy · years"),
        cmd("chartconfig.set-axis-number-format", chartName="IncomeAndLongevity", axis="Category", numberFormat='#,##0'),
        cmd("chartconfig.set-axis-scale", chartName="IncomeAndLongevity", axis="Value", minimumScale=45, maximumScale=90),
        cmd("chartconfig.set-axis-scale", chartName="IncomeAndLongevity", axis="Category", minimumScale=0, maximumScale=110000),
        cmd("chartconfig.set-series-format", chartName="IncomeAndLongevity", seriesIndex=1,
            fillColor=TEAL, fillTransparency=0.32, lineColor="#087F8C", lineWeight=0.8),
        write("World Overview", "B33", [["A comparison, not causation. Select ONE year; use Region to narrow the view. Missing observations are not plotted."]]),
        style("World Overview", ["B33:AB34"], fontColor=MUTED, fontSize=9),
        write("World Overview", "AD29", [["EXPLORE THE EVIDENCE"]]),
        formulas("World Overview", "AD31", [['=HYPERLINK("#\'Chart Data\'!A1","Open country-level chart data")']]),
        style("World Overview", ["AD29:AN33"], fontColor=TEAL, fontSize=10),
    ])
    for name, pivot_name, sheet, title, value_format in [
        ("GrowthLines", "GrowthPivot", "Growth & Resilience", "Real GDP · 2000 = 100", '0"x"'),
        ("InternetLines", "ProgressPivot", "Prosperity & Progress", "Individuals using the Internet · % of population", '0"%"'),
    ]:
        excel.call("chart.create-from-pivottable", pivotTableName=pivot_name, sheetName=sheet,
                   chartType="Line", chartName=name, left=24, top=185, width=760, height=280)
        chart_style(excel, name, title, value_format)
    excel.call("chartconfig.set-axis-number-format", chartName="GrowthLines", axis="Value", numberFormat="0")
    excel.call("chartconfig.set-axis-scale", chartName="InternetLines", axis="Value", minimumScale=0, maximumScale=100)
    excel.batch(
        kpi("Growth & Resilience", "B", "BASELINE", '=2000', "0") +
        kpi("Growth & Resilience", "K", "YEARS OF EVIDENCE", '=25', "0", BLUE) +
        kpi("Growth & Resilience", "T", "SELECTED ECONOMIES", '=COUNTA(\'Pivot Growth\'!B2:Z2)', "0", GOLD) +
        kpi("Prosperity & Progress", "B", "START OF SERIES", '=2000', "0") +
        kpi("Prosperity & Progress", "K", "END OF SERIES", '=2024', "0", BLUE) +
        kpi("Prosperity & Progress", "T", "SELECTED ECONOMIES", '=COUNTA(\'Pivot Progress\'!B2:Z2)', "0", GOLD)
    )
    heat_years = [2008, 2009, 2019, 2020, 2021, 2022, 2023, 2024]
    excel.batch([
        write("Growth & Resilience", "B29", [["ANNUAL GDP GROWTH (%) · FIRST FOUR SELECTED COUNTRIES"]]),
        style("Growth & Resilience", ["B29:AB29"], fontColor=MUTED, fontSize=9, bold=True),
    ])
    # Compact native heatmap: country labels in B, eight year columns spread across the chart width.
    heat_columns = ["I", "L", "O", "R", "U", "X", "AA", "AD"]
    heat_commands = []
    for column, year in zip(heat_columns, heat_years):
        heat_commands += [write("Growth & Resilience", f"{column}30", [[year]])]
    for index in range(4):
        r = 31 + index
        country_ref = f"'Pivot Growth'!{col(index+2)}2"
        heat_commands += [formulas("Growth & Resilience", f"B{r}", [[f'=IF({country_ref}="","",{country_ref})']])]
        for column, year in zip(heat_columns, heat_years):
            formula = observation_formula(r, year, "Growth")
            heat_commands += [formulas("Growth & Resilience", f"{column}{r}", [[formula]])]
    heat_commands += [
        style("Growth & Resilience", ["I31:AD34"], numberFormat='0.0', fontSize=11, bold=True),
        style("Growth & Resilience", ["B30:AD34"], fillColor=PANEL),
        write("Prosperity & Progress", "B30", [["READ THIS CORRECTLY"]]),
        write("Prosperity & Progress", "B32", [["Internet use is a population share. It is not a measure of connection quality or affordability."]]),
        write("Prosperity & Progress", "B33", [["The same country can make progress while remaining far behind its peers. Inspect the trend, not just the rank."]]),
        style("Prosperity & Progress", ["B30:AB34"], fontColor=MUTED, fontSize=10),
    ]
    excel.batch(heat_commands)
    print("Dashboard visuals built", flush=True)


def observation_formula(row, year, metric):
    criteria = (
        f'Observations[CountryCode],INDEX(Countries[Code],MATCH($B{row},Countries[Country],0)),'
        f'Observations[Year],{year},Observations[Metric],"{metric}"'
    )
    return (
        f'=IF($B{row}="","",IF(COUNTIFS({criteria},Observations[Value],"<>")=0,'
        f'"",SUMIFS(Observations[Value],{criteria})))'
    )


def polish(excel):
    operations = []
    for sheet in ("World Overview", "Growth & Resilience", "Prosperity & Progress"):
        operations += [
            formulas("World Overview", "T8", [['=IF(COUNT(\'Pivot Snapshot\'!D4:D28)=0,"Select one year",SUM(\'Pivot Snapshot\'!D4:D28)/1000)']]),
            cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress="B2:AM3"),
            style(sheet, ["B2:AM3"], verticalAlignment="center", fontSize=26),
            style(sheet, ["B8:H10", "K8:Q10", "T8:Z10"], horizontalAlignment="left"),
        ]
    for name, caption, top, height, columns in [
        ("OverviewYear", "Observation year - choose one", 115, 160, 5),
        ("OverviewRegion", "Region", 285, 200, 1),
        ("GrowthCountries", "Compare economies", 115, 350, 2),
        ("ProgressCountries", "Compare economies", 115, 350, 2),
    ]:
        operations.append(cmd("slicer.update-slicer", slicerName=name, slicerOptions={
            "caption": caption, "left": 840, "top": top, "width": 365,
            "height": height, "columnCount": columns, "style": "SlicerStyleDark1",
        }))
    operations += [
        cmd("chartconfig.set-plot-options", chartName="IncomeAndLongevity", plotBy="Columns"),
        cmd("chartconfig.set-series-format", chartName="IncomeAndLongevity", seriesIndex=1,
            fillColor=TEAL, fillTransparency=0.32, lineColor="#087F8C", lineWeight=0.8),
        write("Growth & Resilience", "B29", [["ANNUAL GDP GROWTH (%) · FIRST FOUR SELECTED COUNTRIES"]]),
        write("World Overview", "AD29:AN33", [[""] * 11 for _ in range(5)]),
        write("World Overview", "AD31", [["EXPLORE THE EVIDENCE"]]),
        formulas("World Overview", "AD33", [['=HYPERLINK("#\'Chart Data\'!A1","Open country-level chart data")']]),
    ]
    for column, year in zip(["I", "L", "O", "R", "U", "X", "AA", "AD"],
                            [2008, 2009, 2019, 2020, 2021, 2022, 2023, 2024]):
        operations += [
            formulas("Growth & Resilience", f"{column}31:{column}34",
                     [[observation_formula(row, year, "Growth")] for row in range(31, 35)]),
            cmd("conditionalformat.clear-rules", sheetName="Growth & Resilience", rangeAddress=f"{column}31:{column}34"),
            cmd("conditionalformat.add-rule", sheetName="Growth & Resilience", rangeAddress=f"{column}31:{column}34",
                ruleType="colorScale", colorScaleMinType="number", colorScaleMinValue="-10", colorScaleMinColor="#B43D58",
                colorScaleMidType="number", colorScaleMidValue="0", colorScaleMidColor=PANEL,
                colorScaleMaxType="number", colorScaleMaxValue="10", colorScaleMaxColor="#147D72"),
        ]
    operations += [
        cmd("rangeformat.unmerge-cells", sheetName="Prosperity & Progress", rangeAddress="B30:AN34"),
        write("Prosperity & Progress", "B30:AN34", [[""] * 39 for _ in range(5)]),
        write("Prosperity & Progress", "B29", [["BEYOND THE CONNECTION · FIRST FOUR SELECTED COUNTRIES"]]),
        write("Prosperity & Progress", "B30", [["COUNTRY"]]),
        write("Prosperity & Progress", "K30", [["LIFE EXPECTANCY (2024)"]]),
        write("Prosperity & Progress", "V30", [["ELECTRICITY ACCESS (2024)"]]),
        style("Prosperity & Progress", ["B29:AN30"], fontColor=MUTED, fontSize=9, bold=True),
        style("Prosperity & Progress", ["B31:AN34"], fillColor=PANEL, fontColor=WHITE, fontSize=11),
    ]
    for index in range(4):
        row = 31 + index
        country = f"'Pivot Progress'!{col(index + 2)}2"
        operations += [
            cmd("rangeformat.merge-cells", sheetName="Prosperity & Progress", rangeAddress=f"K{row}:P{row}"),
            cmd("rangeformat.merge-cells", sheetName="Prosperity & Progress", rangeAddress=f"V{row}:AB{row}"),
            formulas("Prosperity & Progress", f"B{row}", [[f'=IF({country}="","",{country})']]),
            formulas("Prosperity & Progress", f"K{row}", [[observation_formula(row, 2024, "Life")]]),
            formulas("Prosperity & Progress", f"V{row}", [[observation_formula(row, 2024, "Electricity")]]),
        ]
    operations += [
        style("Prosperity & Progress", ["K31:K34"], numberFormat='0.0" years"', fontColor=TEAL),
        style("Prosperity & Progress", ["V31:V34"], numberFormat='0.0"%"', fontColor=GOLD),
        cmd("window.set-zoom", sheetName="World Overview", zoom=90),
    ]
    excel.batch(operations)


def configure_sources(excel):
    sheets = {sheet["name"] for sheet in excel.call("sheet.list")["result"]["worksheets"]}
    if "Indicator Notes" not in sheets:
        excel.call("sheet.create", sheetName="Indicator Notes")
        excel.batch(table_commands("Indicator Notes", "IndicatorNotes", read_data("indicators.csv")) + [
            cmd("rangeformat.set-column-width", sheetName="Indicator Notes", rangeAddress="A:B", columnWidth=24),
            cmd("rangeformat.set-column-width", sheetName="Indicator Notes", rangeAddress="C:I", columnWidth=45),
            style("Indicator Notes", ["A1:I10"], wrapText=True, verticalAlignment="top"),
            cmd("rangeformat.auto-fit-rows", sheetName="Indicator Notes", rangeAddress="A1:I10"),
        ])
    excel.batch([
        cmd("workbook.set-document-property", propertyName="Title", value="The World in Motion", scope="built-in"),
        cmd("workbook.set-document-property", propertyName="Author", value="Excel MCP sample", scope="built-in"),
        cmd("workbook.set-document-property", propertyName="Subject",
            value="25 selected countries; World Bank WDI; CC BY 4.0", scope="built-in"),
    ])
    existing = {query["name"] for query in excel.call("powerquery.list")["result"]["queries"]}
    for name, filename in [("EmbeddedSnapshot", "snapshot_source.m"), ("WorldBankLive", "world_bank_source.m")]:
        query = (ROOT / filename).read_text(encoding="utf-8")
        if name in existing:
            excel.call("powerquery.update", queryName=name, mCode=query, refresh=False)
        else:
            excel.call("powerquery.create", queryName=name, mCode=query, loadDestination="connection-only")
    excel.call("powerquery.evaluate", mCode="Table.FirstN(EmbeddedSnapshot, 3)")
    excel.call("powerquery.update", queryName="Observations", mCode="EmbeddedSnapshot", refresh=True)
    if "SourceMode" in existing:
        excel.call("powerquery.delete", queryName="SourceMode")


def redesign_overview(excel):
    sheet = "World Overview"
    countries = read_data("countries.csv")[1:]
    excel.batch([
        cmd("rangeformat.unmerge-cells", sheetName=sheet, rangeAddress="A1:AN39"),
        write(sheet, "A1:AN39", [[""] * 40 for _ in range(39)]),
    ])
    panel(excel, sheet, "THE WORLD IN MOTION",
          "25 selected economies. 25 years of change. One connected Excel model.", 1)
    operations = [
        cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress="B2:Z3"),
        style(sheet, ["B2:Z3"], fontSize=28, verticalAlignment="center"),
        cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress="AC2:AM3"),
        formulas(sheet, "AC2", [['=HYPERLINK("#\'Start Here\'!A34","ASK YOUR AGENT TO ADAPT THIS")']]),
        style(sheet, ["AC2:AM3"], fillColor=PANEL, fontColor=TEAL, fontSize=10,
              bold=True, horizontalAlignment="center", verticalAlignment="center"),
        write("Chart Data", "A1:L1", [[
            "Country", "GDP per person", "Life expectancy", "Population millions",
            "Code", "Region", "Income 2000", "Population millions 2000", "Life 2000",
            "Matched income", "Matched population", "Matched life"]]),
        write("Chart Data", "A2:A26", [[r[1]] for r in countries]),
        write("Chart Data", "E2:F26", [[r[0], r[2]] for r in countries]),
    ]
    for row in range(2, 27):
        lookup = f'MATCH($A{row},\'Pivot Snapshot\'!$A$4:$A$28,0)'
        current = []
        for column in "BCD":
            cell = f"INDEX('Pivot Snapshot'!{column}$4:{column}$28,{lookup})"
            current.append(f'=IF(COUNTIF(\'Pivot Snapshot\'!$A$4:$A$28,$A{row})=0,NA(),'
                           f'IF(COUNT({cell})=0,NA(),{cell}))')
        operations.append(formulas("Chart Data", f"B{row}:D{row}", [current]))
        baselines = []
        for metric, current_column, divisor in [("Income", "B", 1), ("Population", "D", 1000000), ("Life", "C", 1)]:
            criteria = f'Observations[CountryCode],$E{row},Observations[Year],2000,Observations[Metric],"{metric}"'
            baselines.append(f'=IF(AND(ISNUMBER({current_column}{row}),'
                             f'COUNTIFS({criteria},Observations[Value],"<>")>0),'
                             f'SUMIFS(Observations[Value],{criteria})/{divisor},"")')
        operations.append(formulas("Chart Data", f"G{row}:I{row}", [baselines]))
        operations.append(formulas("Chart Data", f"J{row}:L{row}", [[
            f'=IF(ISNUMBER({baseline}{row}),{current_column}{row},"")'
            for baseline, current_column in [("G", "B"), ("H", "D"), ("I", "C")]
        ]]))
    excel.batch(operations)
    tables = excel.call("table.list")["result"]["tables"]
    if not any(table["name"] == "OverviewData" for table in tables):
        excel.call("table.create", sheetName="Chart Data", tableName="OverviewData",
                   rangeAddress="A1:L26", tableStyle="TableStyleMedium2")
    excel.batch([
        style("Chart Data", ["B2:B26", "G2:G26", "J2:J26"], numberFormat="#,##0"),
        style("Chart Data", ["C2:D26", "H2:I26", "K2:L26"], numberFormat="0.0"),
        cmd("rangeformat.auto-fit-columns", sheetName="Chart Data", rangeAddress="A:L"),
    ])
    # Bubble charts share one X row, then pair each country's Y row with its size row.
    names = [["Country"]]
    plot = [[f"=B{row+2}" for row in range(25)]]
    for index, country in enumerate(countries):
        names.extend([[country[1]], ["Population"]])
        for column in "CD":
            plot.append([f"={column}{index+2}" if point == index else "=NA()" for point in range(25)])
    excel.batch([
        write("Chart Data", "N1:N51", names),
        formulas("Chart Data", "O1:AM51", plot),
        cmd("chartconfig.set-source-range", chartName="IncomeAndLongevity", sourceRange="'Chart Data'!N1:AM51"),
        cmd("chartconfig.set-plot-options", chartName="IncomeAndLongevity", plotBy="Rows"),
    ])
    operations = []
    for address in ("B7:I11", "K7:R11", "T7:AA11", "AC7:AM11"):
        operations.append(style(sheet, [address], fillColor=PANEL))
    operations += (
        kpi(sheet, "B", "OBSERVATION YEAR", '=IF(ISNUMBER(\'Pivot Snapshot\'!B1*1),\'Pivot Snapshot\'!B1*1,"Choose one year")', "0") +
        kpi(sheet, "K", "GDP PER PERSON / MEDIAN", '=IF(COUNT(\'Pivot Snapshot\'!B4:B28)=0,"Choose one year",MEDIAN(\'Pivot Snapshot\'!B4:B28))', '#,##0" intl $"', BLUE) +
        kpi(sheet, "T", "PEOPLE IN SELECTION", '=IF(COUNT(\'Pivot Snapshot\'!D4:D28)=0,"Choose one year",SUM(\'Pivot Snapshot\'!D4:D28)/1000)', '0.00" bn"', GOLD) +
        kpi(sheet, "AC", "MEDIAN LIFE EXPECTANCY", '=IF(COUNT(\'Pivot Snapshot\'!C4:C28)=0,"Choose one year",MEDIAN(\'Pivot Snapshot\'!C4:C28))', '0.0" years"', "#C6A6FF")
    )
    for left, right, expression, number_format, color in [
        ("B", "I", '=IF(ISNUMBER(B8),COUNTA(\'Pivot Snapshot\'!A4:A28),"Choose one year")', '0" selected countries"', MUTED),
        ("K", "R", '=IF(COUNT(\'Chart Data\'!J2:J26)=0,"Choose one year",IF(MEDIAN(\'Chart Data\'!G2:G26)=0,"No baseline",MEDIAN(\'Chart Data\'!J2:J26)/MEDIAN(\'Chart Data\'!G2:G26)-1))', '+0.0%" vs 2000";-0.0%" vs 2000";0.0%" vs 2000"', BLUE),
        ("T", "AA", '=IF(COUNT(\'Chart Data\'!K2:K26)=0,"Choose one year",IF(SUM(\'Chart Data\'!H2:H26)=0,"No baseline",SUM(\'Chart Data\'!K2:K26)/SUM(\'Chart Data\'!H2:H26)-1))', '+0.0%" vs 2000";-0.0%" vs 2000";0.0%" vs 2000"', GOLD),
        ("AC", "AM", '=IF(COUNT(\'Chart Data\'!L2:L26)=0,"Choose one year",MEDIAN(\'Chart Data\'!L2:L26)-MEDIAN(\'Chart Data\'!I2:I26))', '+0.0" years vs 2000";-0.0" years vs 2000";0.0" years vs 2000"', "#C6A6FF"),
    ]:
        operations += [
            cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress=f"{left}11:{right}11"),
            formulas(sheet, f"{left}11", [[expression]]),
            style(sheet, [f"{left}11:{right}11"], fontColor=color, fontSize=10, numberFormat=number_format),
        ]
    operations += [
        style(sheet, ["B8:H10", "K8:Q10", "T8:Z10", "AC8:AM10"], horizontalAlignment="left", fontSize=27),
        write(sheet, "B13", [["PROSPERITY HAS MORE THAN ONE DIMENSION"]]),
        style(sheet, ["B13:AB13"], fontSize=15, bold=True),
        write(sheet, "AC13", [["CHANGE THE VIEW"]]),
        style(sheet, ["AC13:AM13"], fontColor=TEAL, bold=True, fontSize=11),
    ]
    legends = [
        ("B14", "East Asia / Pacific", "East Asia & Pacific"),
        ("I14", "Europe / Central Asia", "Europe & Central Asia"),
        ("Q14", "Latin America / Caribbean", "Latin America & Caribbean"),
        ("B15", "Middle East / N. Africa", "Middle East & North Africa"),
        ("I15", "North America", "North America"),
        ("P15", "South Asia", "South Asia"),
        ("V15", "Sub-Saharan Africa", "Sub-Saharan Africa"),
    ]
    for cell, text, region in legends:
        operations += [write(sheet, cell, [[text]]), style(sheet, [cell], fontColor=REGION_COLORS[region], fontSize=9, bold=True)]
    operations += [
        cmd("rangeformat.merge-cells", sheetName=sheet, rangeAddress="B34:AB35"),
        formulas(sheet, "B34", [['=IF(AND(ISNUMBER(K11),ISNUMBER(AC11)),'
                               '"SINCE 2000  /  Median GDP per person "&TEXT(K11,"+0%;-0%;0%")&'
                               '"  /  Life expectancy "&IF(AC11>0,"+","")&ROUND(AC11,1)&" years. Matched countries.",'
                               '"Choose one year to compare countries. Missing observations are not plotted.")']]),
        style(sheet, ["B34:AB35"], fillColor=PANEL, fontColor=WHITE, fontSize=11,
              wrapText=True, verticalAlignment="center"),
        formulas(sheet, "AC38", [['=HYPERLINK("#\'Chart Data\'!A1","Inspect the supporting numbers")']]),
        style(sheet, ["AC38:AM38"], fontColor=MUTED, fontSize=9),
        write(sheet, "B38", [["World Bank WDI / CC BY 4.0   |   Bubble area = population. Association is not causation."]]),
        cmd("chart.move", chartName="IncomeAndLongevity", left=22, top=258, width=838, height=296),
        cmd("chartconfig.set-style", chartName="IncomeAndLongevity", styleId=42),
        cmd("chartconfig.set-title", chartName="IncomeAndLongevity", title=""),
        cmd("chartconfig.set-area-format", chartName="IncomeAndLongevity", area="Chart", fillColor=NAVY, lineColor=NAVY),
        cmd("chartconfig.set-area-format", chartName="IncomeAndLongevity", area="Plot", fillColor=NAVY, lineColor=NAVY),
        cmd("chartconfig.show-legend", chartName="IncomeAndLongevity", visible=False),
        cmd("chartconfig.set-axis-scale", chartName="IncomeAndLongevity", axis="Value", minimumScale=45, maximumScale=90, majorUnit=15),
        cmd("chartconfig.set-gridlines", chartName="IncomeAndLongevity", axis="Value", showMajor=False, showMinor=False),
        cmd("chartconfig.set-axis-title", chartName="IncomeAndLongevity", axis="Category", title="GDP per person / constant 2021 international $ (PPP)"),
        cmd("chartconfig.set-axis-title", chartName="IncomeAndLongevity", axis="Value", title="Life expectancy / years"),
        cmd("chartconfig.set-axis-number-format", chartName="IncomeAndLongevity", axis="Category", numberFormat='0,"k"'),
    ]
    for index, country in enumerate(countries, 1):
        color = REGION_COLORS[country[2]]
        operations.append(cmd("chartconfig.set-series-format", chartName="IncomeAndLongevity",
                              seriesIndex=index, fillColor=color, fillTransparency=.22,
                              lineColor=color, lineWeight=1))
        if country[0] in ("CHN", "IND", "DEU", "NGA", "USA"):
            operations.append(cmd("chartconfig.set-data-labels", chartName="IncomeAndLongevity",
                                  seriesIndex=index, showValue=False, showCategoryName=False,
                                  showSeriesName=True, showBubbleSize=False,
                                  labelPosition={"IND": "Below", "DEU": "Above"}.get(country[0], "Right")))
    for name, caption, top, height, columns in [
        ("OverviewYear", "Year / choose one", 222, 150, 5),
        ("OverviewRegion", "Regions", 384, 190, 1),
    ]:
        operations.append(cmd("slicer.update-slicer", slicerName=name, slicerOptions={
            "caption": caption, "left": 890, "top": top, "width": 315,
            "height": height, "columnCount": columns, "style": "SlicerStyleDark1",
        }))
    excel.batch(operations)
    excel.call("window.set-zoom", sheetName=sheet, zoom=90)


def agent_instructions(excel):
    excel.batch([
        write("Start Here", "B2", [["Ask your agent to explore this workbook with ExcelMCP. You can also use its native Excel filters. Year and Region affect World Overview; each country selector affects only its own trend dashboard."]]),
        write("Start Here", "B24:B25", [
            [AGENT_LIVE_REFRESH], [AGENT_OFFLINE]]),
        write("Start Here", "A34:B39", [
            ["MAKE IT YOURS WITH YOUR AGENT", "Download the workbook, then ask your agent to adapt it through ExcelMCP. You do not need to edit formulas, queries or the Data Model yourself."],
            ["Requirements for adaptation", "Use an agent connected to ExcelMCP on Windows with desktop Excel installed. Saved charts and filters work without an agent."],
            ["Example prompt", "Open this workbook with ExcelMCP. Compare Germany, France and Italy from 2000 to 2024. Adapt the dashboard for that comparison, retain the original data sources and missing-value handling, and save the changes."],
            ["New reporting needs", "Tell your agent which countries, questions and measures you need. Ask it to explain data limitations and verify its results in Excel."],
            ["Overview comparisons", "Changes since 2000 compare the same selected countries with observations in both years, separately for each metric. GDP per person and life expectancy use unweighted country medians. Population uses summed people."],
            ["Region colors", "Each country's bubble keeps its region color as the year or region filters change. Country identities are fixed in OverviewData; filtered-out values are not plotted."],
        ]),
        style("Start Here", ["A34:B39"], fontName="Aptos", fontSize=11, wrapText=True),
        style("Start Here", ["A34:B34"], fillColor=NAVY, fontColor=WHITE, bold=True),
        cmd("rangeformat.auto-fit-rows", sheetName="Start Here", rangeAddress="A1:B39"),
        cmd("rangeformat.validate-range", sheetName="Start Here", rangeAddress="B32",
            validationType="custom", formula1="=FALSE()", showInputMessage=True,
            inputTitle="Load status",
            inputMessage="Ask your agent to change the Observations query source. This cell reports the last successful load."),
    ])


def documentation(excel):
    manifest = json.loads((ROOT / "data" / "sources.json").read_text(encoding="utf-8"))
    lines = [
        ["WORLD IN MOTION", "A working Excel MCP sample, using real World Bank data."],
        ["Explore", "Use the three dashboard tabs. Year and Region affect World Overview; country selectors affect their own trend chart."],
        ["Year selection", "Select exactly one year on World Overview. Multiple years produce blank measures, not misleading sums."],
        ["Slicer controls", "Use Ctrl-click or Excel's multi-select button for multiple countries. The clear-filter button restores all countries."],
        ["Requirements", "Full experience tested in Microsoft 365 desktop Excel for Windows. No MCP, CLI, Python, API key, or internet required to explore saved results."],
        ["Coverage", "25 selected countries; 2000-2024; 9 annual indicators. This is not the whole world or a representative global aggregate."],
        ["Source", manifest["source"]],
        ["Retrieved", manifest["retrieved"]],
        ["Licence", "Each selected indicator is CC BY-4.0 in the downloaded WDISeries.csv metadata."],
        ["Terms", manifest["terms"]],
        ["Attribution", manifest["attribution"]],
        ["Changes", "Selected countries/years/indicators, reshaped the observations, and added sample calculations and visualisations."],
        ["Income", "GDP per capita at purchasing-power parity, in constant 2021 international dollars. This is NOT salary or household disposable income."],
        ["Median income", "Unweighted median of the selected countries' GDP per capita, not population-weighted and not an official World Bank regional estimate."],
        ["GDP index", "Each country's constant-2015-US-dollar GDP divided by its own 2000 GDP, multiplied by 100. Measures growth, not economic size."],
        ["Growth and inflation", "Annual percent change is stored in percentage-point units: 5 means 5%, not 500%."],
        ["Population", "Country populations are summed only within one year. The chart's bubble areas use population in millions."],
        ["Life expectancy", "Period life expectancy at birth, in years. Cross-country association with income does not establish causation."],
        ["Internet and electricity", "Population shares in percent. Availability and definitions vary; inspect indicator metadata."],
        ["Missing values", f"{manifest['missing']} of {manifest['observations']} snapshot observations are missing. They remain blank; no interpolation or zero filling."],
        ["Classifications", "Country names, regions and income groups reflect the downloaded metadata, not historical classifications."],
        ["Revisions", "WDI revises historical data. A future live refresh may change earlier years; the published sample is a fixed, attributed snapshot."],
        ["Refresh", "Snapshot mode reprocesses the embedded real observations. Live mode downloads the official WDI bulk ZIP (about 283 MB); internet is required."],
        ["Live refresh steps", AGENT_LIVE_REFRESH],
        ["Offline mode", AGENT_OFFLINE],
        ["Architecture", "SourceObservations -> Power Query Observations -> Data Model. Countries, Indicators and Years provide related lookup tables. DAX measures feed native PivotTables and linked PivotCharts."],
        ["Scope of controls", "The growth heatmap shows the first four alphabetically selected countries. It is tied to the country selection, but its displayed crisis years are fixed."],
        ["Interpretation", "Educational demonstration, not investment, policy, or business advice. World Bank has not endorsed this sample."],
    ]
    excel.batch([
        write("Start Here", "A1:B28", lines),
        cmd("rangeformat.set-column-width", sheetName="Start Here", rangeAddress="A:A", columnWidth=25),
        cmd("rangeformat.set-column-width", sheetName="Start Here", rangeAddress="B:B", columnWidth=115),
        style("Start Here", ["A1:B28"], fontName="Aptos", fontSize=11, wrapText=True),
        style("Start Here", ["A1:B1"], fillColor=NAVY, fontColor=WHITE, bold=True, fontSize=18),
        cmd("rangeformat.auto-fit-rows", sheetName="Start Here", rangeAddress="A1:B28"),
        write("Start Here", "A31:B32", [["Setting", "Value"], ["Last successful source", "Snapshot"]]),
        cmd("table.create", sheetName="Start Here", tableName="RefreshSettings", rangeAddress="A31:B32", tableStyle="TableStyleMedium2"),
        formulas("Start Here", "B32", [['=INDEX(Observations[SourceMode],1)']]),
    ])
    configure_sources(excel)
    excel.batch([cmd("pivottable.refresh", pivotTableName=name)
                 for name in ("SnapshotPivot", "GrowthPivot", "ProgressPivot")])
    excel.call("window.set-zoom", sheetName="World Overview", zoom=90)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cli", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--logs", type=Path, required=True)
    parser.add_argument("--model-only", action="store_true")
    parser.add_argument("--resume-model", action="store_true")
    parser.add_argument("--resume-visuals", action="store_true")
    parser.add_argument("--polish-only", action="store_true")
    parser.add_argument("--overview-only", action="store_true",
                        help="Redesign only World Overview and agent instructions in an existing workbook")
    args = parser.parse_args()
    excel = Excel(args.cli, args.logs)
    if args.resume_model or args.resume_visuals or args.polish_only or args.overview_only:
        excel.open(args.output, show=True)
    else:
        excel.create(args.output)
    save = False
    try:
        if args.overview_only:
            redesign_overview(excel)
            agent_instructions(excel)
            save = True
            return
        if args.polish_only:
            polish(excel)
            configure_sources(excel)
            excel.batch([
                write("Start Here", "B24:B25", [
                    [AGENT_LIVE_REFRESH], [AGENT_OFFLINE]]),
                write("Start Here", "A32", [["Last successful source"]]),
                formulas("Start Here", "B32", [['=INDEX(Observations[SourceMode],1)']]),
                cmd("rangeformat.validate-range", sheetName="Start Here", rangeAddress="B32",
                    validationType="custom", formula1="=FALSE()",
                    showInputMessage=True, inputTitle="Load status",
                    inputMessage="Ask your agent to change the Observations query source. This cell reports the last successful load."),
            ])
            redesign_overview(excel)
            agent_instructions(excel)
            save = True
            return
        if not args.resume_model and not args.resume_visuals:
            build_model(excel)
            excel.close(True)
            excel.open(args.output, show=True)
        if not args.resume_visuals:
            build_pivots(excel)
        if not args.model_only:
            build_dashboards(excel)
            documentation(excel)
            polish(excel)
            redesign_overview(excel)
            agent_instructions(excel)
        save = True
    finally:
        excel.close(save)
    print(f"Saved {args.output}", flush=True)


if __name__ == "__main__":
    main()
