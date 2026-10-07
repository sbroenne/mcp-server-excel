"""Render the public usage-analytics page from ``.github/usage-analytics.json``.

Charts are plain HTML/CSS (styled in ``docs/assets/stylesheets/extra.css``) so the
page needs no JavaScript and every bar carries an accessible label.
"""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from html import escape

from sitegen.sources import read


def _analytics_cell(value: object) -> str:
    """Format a validated aggregate value for a Markdown table."""
    if isinstance(value, float):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
    elif isinstance(value, int):
        text = f"{value:,}"
    else:
        text = str(value)
    return text.replace("|", r"\|").replace("\r", " ").replace("\n", " ")


_ANALYTICS_FAMILY_NAMES = {
    "range": "Reading and writing cells",
    "file": "Managing workbooks",
    "range_format": "Formatting cells",
    "vba": "Running and editing macros",
    "worksheet": "Working with worksheets",
    "powerquery": "Refreshing and checking data",
    "range_edit": "Finding, sorting, and editing cells",
    "calculation_mode": "Calculating formulas",
    "screenshot": "Taking screenshots",
    "table": "Working with Excel tables",
    "datamodel": "Working with the Data Model",
}

_ANALYTICS_HERO_FEATURE_NAMES = {
    "power-query": "Power Query & M code",
    "power-pivot-dax": "Power Pivot & DAX",
    "pivottables-charts": "PivotTables & charts",
    "tables-ranges": "Tables & ranges",
    "vba": "VBA macros",
    "worksheets-connections": "Worksheets & connections",
    "agent-mode": "Agent mode",
    "python-in-excel": "Python in Excel",
    "other": "Other features",
}

_ANALYTICS_OPERATION_NAMES = {
    "range/get-values": "Read cell values",
    "range/set-values": "Write cell values",
    "file/open": "Open a workbook",
    "file/close": "Close a workbook",
    "range/set-formulas": "Write formulas",
    "range/get-used-range": "Find the used area",
    "range_format/format-range": "Format cells",
    "range/get-formulas": "Read formulas",
    "file/list": "List open workbooks",
    "worksheet/list": "List worksheets",
    "screenshot/capture": "Take a screenshot",
    "range_format/set-column-width": "Set column width",
    "vba/run": "Run a macro",
    "range_edit/find": "Find cells",
    "range/set-number-format": "Set number format",
    "powerquery/refresh": "Refresh a Power Query",
    "powerquery/evaluate": "Run Power Query code",
    "powerquery/create": "Create a Power Query",
    "powerquery/update": "Update a Power Query",
    "datamodel/evaluate": "Run a Data Model query",
    "datamodel/refresh": "Refresh the Data Model",
    "pivottable/create-from-range": "Create a PivotTable",
    "pivottable/refresh": "Refresh a PivotTable",
    "connection/refresh": "Refresh a data connection",
}

_ANALYTICS_LEVEL_NAMES = {
    "light": "Quick read",
    "medium": "Everyday edit",
    "heavy": "Heavy data work",
}

_ANALYTICS_ENTRY_POINT_NAMES = {
    "mcp-server": "AI assistant (MCP Server)",
    "cli": "Command line (excelcli)",
}


def _analytics_name(value: object, names: dict[str, str]) -> str:
    """Replace an internal action name with a reader-friendly label."""
    raw = str(value)
    if raw in names:
        return names[raw]
    return raw.replace("/", " ").replace("_", " ").replace("-", " ").title()


def _analytics_bar_chart(
    rows: list[dict[str, object]],
    *,
    label_field: str,
    value_field: str,
    value_suffix: str = "",
    display_field: str | None = None,
    work: bool = False,
) -> str:
    """Render an accessible horizontal comparison chart."""
    maximum = max((float(row[value_field]) for row in rows), default=0)
    modifier = " analytics-bars__track--work" if work else ""
    lines = ['<div class="analytics-bars" role="list">']
    for row in rows:
        value = float(row[value_field])
        width = 0 if maximum == 0 else max(2, value / maximum * 100)
        label = escape(str(row[label_field]))
        display_value = (
            str(row[display_field])
            if display_field
            else f"{_analytics_cell(row[value_field])}{value_suffix}"
        )
        lines.extend(
            [
                '  <div class="analytics-bars__row" role="listitem">',
                '    <div class="analytics-bars__label">',
                f"      <span>{label}</span><strong>{escape(display_value)}</strong>",
                "    </div>",
                f'    <div class="analytics-bars__track{modifier}" aria-hidden="true">',
                f'      <span style="width: {width:.2f}%"></span>',
                "    </div>",
                "  </div>",
            ]
        )
    lines.append("</div>")
    return "\n".join(lines)


def _analytics_paired_bar_chart(
    rows: list[dict[str, object]],
    *,
    label_field: str,
    first_field: str,
    second_field: str,
    first_name: str,
    second_name: str,
    value_suffix: str = "%",
    display_field: str | None = None,
    scale_each_row: bool = False,
) -> str:
    """Render two bars per row, on one shared scale or scaled within each row."""
    shared_maximum = max(
        (
            max(float(row[first_field]), float(row[second_field]))
            for row in rows
        ),
        default=0,
    )
    lines = [
        '<div class="analytics-bars" role="list">',
        '  <div class="analytics-bars__legend" aria-hidden="true">',
        f"    <span><i></i>{escape(first_name)}</span>",
        '    <span><i class="analytics-bars__swatch--work"></i>'
        f"{escape(second_name)}</span>",
        "  </div>",
    ]
    for row in rows:
        label = escape(str(row[label_field]))
        first = f"{_analytics_cell(row[first_field])}{value_suffix}"
        second = f"{_analytics_cell(row[second_field])}{value_suffix}"
        shown = str(row[display_field]) if display_field else f"{first} / {second}"
        maximum = (
            max(float(row[first_field]), float(row[second_field]))
            if scale_each_row
            else shared_maximum
        )
        lines.extend(
            [
                '  <div class="analytics-bars__row" role="listitem" '
                f'aria-label="{label}: {escape(first_name)} {escape(first)}, '
                f'{escape(second_name)} {escape(second)}">',
                '    <div class="analytics-bars__label" aria-hidden="true">',
                f"      <span>{label}</span>"
                f"<strong>{escape(shown)}</strong>",
                "    </div>",
            ]
        )
        for field, modifier in ((first_field, ""), (second_field, " analytics-bars__track--work")):
            value = float(row[field])
            width = 0 if maximum == 0 else max(2, value / maximum * 100)
            lines.extend(
                [
                    f'    <div class="analytics-bars__track{modifier}" aria-hidden="true">',
                    f'      <span style="width: {width:.2f}%"></span>',
                    "    </div>",
                ]
            )
        lines.append("  </div>")
    lines.append("</div>")
    return "\n".join(lines)


def _analytics_week_chart(
    rows: list[dict[str, object]],
    *,
    value_field: str,
    title: str,
) -> str:
    """Render weekly values as an accessible compact bar chart."""
    maximum = float(max((float(row[value_field]) for row in rows), default=0))
    midpoint = maximum / 2
    if maximum.is_integer() and maximum >= 10:
        midpoint = round(midpoint)
    lines = [
        '<div class="analytics-week-chart" role="group" '
        f'aria-label="{escape(title)}">',
        f"  <strong>{escape(title)}</strong>",
        '  <div class="analytics-week-chart__body">',
        '    <div class="analytics-week-chart__y-axis" aria-hidden="true">',
        f"      <span>{escape(_analytics_cell(maximum))}</span>",
        f"      <span>{escape(_analytics_cell(midpoint))}</span>",
        "      <span>0</span>",
        "    </div>",
        '  <div class="analytics-week-chart__plot" role="list">',
    ]
    for row in rows:
        value = float(row[value_field])
        height = 0 if maximum == 0 else max(2, value / maximum * 100)
        week = datetime.fromisoformat(str(row["week"]))
        label = week.strftime("%b %d")
        display_value = _analytics_cell(row[value_field])
        lines.extend(
            [
                '    <div class="analytics-week-chart__week" role="listitem" '
                f'aria-label="Week of {escape(label)}: {escape(display_value)}">',
                f'      <span style="height: {height:.2f}%" aria-hidden="true"></span>',
                f"      <small>{escape(label)}</small>",
                "    </div>",
            ]
        )
    lines.extend(["    </div>", "  </div>", "</div>"])
    return "\n".join(lines)


def _analytics_version_chart(rows: list[dict[str, object]]) -> str:
    """Render weekly release adoption as a 100% stacked column chart."""
    palette = (
        "#4051b5",
        "#008b8b",
        "#d97706",
        "#db2777",
        "#7c3aed",
        "#15803d",
        "#dc2626",
        "#64748b",
        "#0891b2",
    )
    weeks: dict[str, dict[str, dict[str, object]]] = {}
    totals: dict[str, int] = {}
    for row in rows:
        week = str(row["week"])
        version = str(row["version"])
        weeks.setdefault(week, {})[version] = row
        totals[version] = totals.get(version, 0) + int(row["users"])

    versions = sorted(
        totals,
        key=lambda version: (version == "Other", -totals[version], version),
    )
    colors = {
        version: palette[index % len(palette)]
        for index, version in enumerate(versions)
    }
    lines = [
        '<div class="analytics-version-chart" role="group" '
        'aria-label="Share of users by release each week">',
        '  <div class="analytics-version-chart__legend" aria-hidden="true">',
    ]
    for version in versions:
        lines.append(
            '    <span><i style="background: '
            f'{colors[version]}"></i>{escape(version)}</span>'
        )
    lines.extend(
        [
            "  </div>",
            '  <div class="analytics-version-chart__body">',
            '    <div class="analytics-version-chart__y-axis" aria-hidden="true">',
            "      <span>100%</span>",
            "      <span>50%</span>",
            "      <span>0%</span>",
            "    </div>",
            '    <div class="analytics-version-chart__plot" role="list">',
        ]
    )
    for week_value in sorted(weeks):
        week = datetime.fromisoformat(week_value)
        label = week.strftime("%b %d")
        entries = weeks[week_value]
        summary = ", ".join(
            f"{version}: {_analytics_cell(entries[version]['sharePct'])}%"
            for version in versions
            if version in entries
        )
        lines.extend(
            [
                '      <div class="analytics-version-chart__week" role="listitem" '
                f'aria-label="Week of {escape(label)}. {escape(summary)}">',
                '        <div class="analytics-version-chart__stack" aria-hidden="true">',
            ]
        )
        for version in versions:
            if version not in entries:
                continue
            row = entries[version]
            share = float(row["sharePct"])
            title = (
                f"{version}: {_analytics_cell(row['sharePct'])}% "
                f"({_analytics_cell(row['users'])} users)"
            )
            lines.append(
                f'          <span title="{escape(title)}" '
                f'style="height: {share:.2f}%; background: {colors[version]}"></span>'
            )
        lines.extend(
            [
                "        </div>",
                f"        <small>{escape(label)}</small>",
                "      </div>",
            ]
        )
    lines.extend(["    </div>", "  </div>", "</div>"])
    return "\n".join(lines)


def _analytics_weighted_feature_section(
    report: dict[str, object], hero_rows: list[dict[str, object]]
) -> list[str]:
    return [
        "The bars group actions by the main features highlighted on the Excel MCP "
        "homepage. Each feature has two bars. **Share of actions** counts every "
        "action once. **Share of work** gives heavier actions more weight, so "
        "refreshing a Power Query counts for more than reading a few cells. See "
        "[how share of work is calculated](#how-share-of-work-is-calculated). "
        "Smaller capabilities are grouped as **Other features**.",
        "",
        _analytics_paired_bar_chart(
            hero_rows,
            label_field="friendlyName",
            first_field="sharePct",
            second_field="workSharePct",
            first_name="Share of actions",
            second_name="Share of work",
        ),
        "",
    ]


def _analytics_work_sections(report: dict[str, object]) -> list[str]:
    levels = report["weights"]
    summary = report["summary"]
    heavy = report["heavyWork"]
    sections = [
        "## How share of work is calculated",
        "",
        "!!! info \"An estimate, not a measurement\"\n"
        f"    Every action is given one of three fixed effort levels. Quick reads, "
        f"such as reading cells or listing worksheets, count "
        f"**{levels['light']}**. Everyday edits, such as writing values or "
        f"formatting cells, count **{levels['medium']}**. Heavy data work, such "
        "as refreshing Power Query or the Data Model, running macros, or "
        f"building PivotTables, counts **{levels['heavy']}**. The levels are "
        "chosen by the maintainers and checked automatically whenever an action "
        "is added. They describe the kind of work, not how long it took. See the "
        "[full list of levels](https://github.com/sbroenne/mcp-server-excel/"
        "blob/main/.github/usage-analytics-weights.json).",
        "",
    ]
    if int(summary["unweightedActions"]) > 0:
        sections.extend(
            [
                f"**{_analytics_cell(summary['unweightedActions'])} actions** came "
                "from older releases that used action names which no longer "
                "exist. They are counted as actions but left out of share of work.",
                "",
            ]
        )
    sections.extend(
        [
            f"**{_analytics_cell(heavy['heavyUserSharePct'])}%** of people who "
            "used Excel MCP in this period did at least one piece of heavy data "
            "work.",
            "",
            "## Where most of the work goes",
            "",
            "These actions add up to the largest share of estimated work. Common "
            "light actions can still appear here when they are used very often.",
            "",
            _analytics_bar_chart(
                [
                    {
                        **row,
                        "friendlyName": _analytics_name(
                            row["name"], _ANALYTICS_OPERATION_NAMES
                        )
                        + " ("
                        + _ANALYTICS_LEVEL_NAMES.get(
                            str(row["level"]), str(row["level"]).title()
                        ).lower()
                        + ")",
                        "shown": f"{_analytics_cell(row['workSharePct'])}% of work, "
                        f"used {_analytics_cell(row['actions'])} times",
                    }
                    for row in report["operationsByWork"][:10]
                ],
                label_field="friendlyName",
                value_field="workSharePct",
                display_field="shown",
                work=True,
            ),
            "",
        ]
    )
    return sections


def _analytics_entry_point_sections(
    report: dict[str, object], date_format: str
) -> list[str]:
    windows = report["windows"]
    since = datetime.fromisoformat(
        str(windows["entryPointSinceUtc"]).replace("Z", "+00:00")
    )
    minimum = windows["entryPointMinimumUsers"]
    sections = [
        "## Command line and AI assistant",
        "",
        "Excel MCP can be used through an AI assistant, which talks to the MCP "
        "Server, or directly from the command line with `excelcli`. Each action "
        f"has recorded which of the two was used since **{since.strftime(date_format)}**, "
        "so this comparison covers a shorter period than the rest of the page. "
        "The two groups are mostly different people doing different jobs, so "
        "differences describe how each is used, not which is better. A group is "
        f"shown only when it has at least **{_analytics_cell(minimum)} users**.",
        "",
    ]
    shown = [row for row in report["entryPoints"] if row.get("enoughData")]
    hidden = [row for row in report["entryPoints"] if not row.get("enoughData")]
    if shown:
        sections.extend(
            [
                _analytics_paired_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": _analytics_name(
                                row["name"], _ANALYTICS_ENTRY_POINT_NAMES
                            )
                            + f" ({_analytics_cell(row['users'])} users)",
                            "shown": f"{_analytics_cell(row['actionsPerUser'])} actions / "
                            f"{_analytics_cell(row['workUnitsPerUser'])} work per user",
                        }
                        for row in shown
                    ],
                    label_field="friendlyName",
                    first_field="actionsPerUser",
                    second_field="workUnitsPerUser",
                    first_name="Actions per user",
                    second_name="Estimated work per user",
                    value_suffix="",
                    display_field="shown",
                ),
                "",
            ]
        )
    for row in hidden:
        name = _analytics_name(row["name"], _ANALYTICS_ENTRY_POINT_NAMES)
        sections.extend(
            [f"There is not enough data yet to show **{name}** on its own.", ""]
        )
    if len(shown) < 2:
        return sections

    entry_points = [str(row["name"]) for row in shown]
    first_point, second_point = entry_points[0], entry_points[1]
    features: dict[str, dict[str, object]] = {}
    for row in report["entryPointFeatures"]:
        feature = features.setdefault(
            str(row["name"]),
            {
                "friendlyName": _analytics_name(
                    row["name"], _ANALYTICS_HERO_FEATURE_NAMES
                ),
                first_point: 0.0,
                second_point: 0.0,
            },
        )
        if str(row["entryPoint"]) in (first_point, second_point):
            feature[str(row["entryPoint"])] = float(row["workSharePct"])
    feature_rows = sorted(
        features.values(),
        key=lambda item: -max(float(item[first_point]), float(item[second_point])),
    )
    short_names = {"mcp-server": "AI assistant", "cli": "Command line"}
    sections.extend(
        [
            "### What each group works on",
            "",
            "Each feature's share of estimated work within each group.",
            "",
            _analytics_paired_bar_chart(
                feature_rows,
                label_field="friendlyName",
                first_field=first_point,
                second_field=second_point,
                first_name=short_names.get(first_point, first_point),
                second_name=short_names.get(second_point, second_point),
            ),
            "",
            "### Most common actions in each group",
            "",
        ]
    )
    for entry_point in entry_points:
        sections.extend(
            [
                f"**{_analytics_name(entry_point, _ANALYTICS_ENTRY_POINT_NAMES)}**",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": _analytics_name(
                                row["name"], _ANALYTICS_OPERATION_NAMES
                            ),
                        }
                        for row in report["entryPointOperations"]
                        if row["entryPoint"] == entry_point
                    ][:5],
                    label_field="friendlyName",
                    value_field="actions",
                ),
                "",
            ]
        )
    return sections


def _analytics_feature_name(value: object) -> str:
    return _analytics_name(value, _ANALYTICS_HERO_FEATURE_NAMES)


def _analytics_habit_sections(habits: dict[str, object]) -> list[str]:
    days = habits["windowDays"]
    sessions = habits["assistantSessions"]
    returning = habits["returningUsers"]
    sections = [
        "## How people work",
        "",
        f"These views use the last **{days} days** unless stated otherwise. "
        "Groups with fewer than "
        f"**{_analytics_cell(habits['minimumUsers'])} users** are not shown.",
        "",
        "### Size of AI assistant sessions",
        "",
        "A session is one run of the MCP Server inside an AI assistant. The "
        "command line is left out because every `excelcli` command runs on its "
        f"own. Out of **{_analytics_cell(sessions['sessions'])} sessions**, the "
        f"typical one had **{_analytics_cell(sessions['medianActions'])} actions**, "
        f"and **{_analytics_cell(sessions['multiFeatureSharePct'])}%** used two or "
        "more areas of Excel. A small number of long sessions do much of the work.",
        "",
        _analytics_paired_bar_chart(
            [
                {
                    **row,
                    "friendlyName": f"{row['size']} action"
                    + ("" if row["size"] == "1" else "s"),
                }
                for row in sessions["sizes"]
            ],
            label_field="friendlyName",
            first_field="sessionSharePct",
            second_field="actionSharePct",
            first_name="Share of sessions",
            second_name="Share of actions",
        ),
        "",
    ]
    if habits["featurePairs"]:
        sections.extend(
            [
                "### Areas used together",
                "",
                "How often two areas of Excel appear in the same AI assistant "
                "session, as a share of all sessions.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": f"{_analytics_feature_name(row['first'])} + "
                            f"{_analytics_feature_name(row['second'])}",
                            "shown": f"{_analytics_cell(row['sharePct'])}% "
                            f"({_analytics_cell(row['sessions'])} sessions)",
                        }
                        for row in habits["featurePairs"]
                    ],
                    label_field="friendlyName",
                    value_field="sharePct",
                    display_field="shown",
                ),
                "",
            ]
        )
    if returning["newUsers"]:
        sections.extend(
            [
                "### Do new users come back?",
                "",
                f"Of **{_analytics_cell(returning['newUsers'])} people** first seen "
                "between 4 and 12 weeks ago, this is how many used Excel MCP again "
                "later. People first seen before the 90-day window may be counted as new.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            "label": "Came back after a week or more",
                            "value": returning["returnedAfterWeekPct"],
                            "shown": f"{_analytics_cell(returning['returnedAfterWeekPct'])}% "
                            f"({_analytics_cell(returning['returnedAfterWeek'])} people)",
                        },
                        {
                            "label": "Came back after three weeks or more",
                            "value": returning["returnedAfterThreeWeeksPct"],
                            "shown": f"{_analytics_cell(returning['returnedAfterThreeWeeksPct'])}% "
                            f"({_analytics_cell(returning['returnedAfterThreeWeeks'])} people)",
                        },
                    ],
                    label_field="label",
                    value_field="value",
                    display_field="shown",
                ),
                "",
            ]
        )
    if habits["featureWait"]:
        sections.extend(
            [
                "### Typical wait by area",
                "",
                "How long a typical action takes from request to answer, including "
                "Excel's own work. The second number is the wait that 1 in 10 "
                "actions goes past. Waits depend on workbook size and the computer, "
                "so they show which areas are heavier, not how fast Excel MCP is.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": _analytics_feature_name(row["name"]),
                            "shown": f"{_analytics_cell(row['typicalSeconds'])} s typical; "
                            f"1 in 10 over {_analytics_cell(row['slowSeconds'])} s",
                        }
                        for row in habits["featureWait"]
                    ],
                    label_field="friendlyName",
                    value_field="typicalSeconds",
                    display_field="shown",
                    work=True,
                ),
                "",
            ]
        )
    if habits["weekdays"]:
        sections.extend(
            [
                "### Weekdays and weekends",
                "",
                f"Total actions on each day of the week, added up over the last "
                f"**{habits['weekdayWeeks']} complete weeks** (UTC). A single "
                f"weekday averaged **{_analytics_cell(habits['workdayAverageActions'])} "
                "actions**, compared with "
                f"**{_analytics_cell(habits['weekendAverageActions'])}** on a "
                "single weekend day.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "shown": f"{_analytics_cell(row['actions'])} actions, "
                            f"{_analytics_cell(row['users'])} users",
                        }
                        for row in habits["weekdays"]
                    ],
                    label_field="day",
                    value_field="actions",
                    display_field="shown",
                ),
                "",
            ]
        )
    if habits["firstAdvancedUse"]:
        sections.extend(
            [
                "### When people first try advanced features",
                "",
                "Of the people who used each advanced area, the share who first "
                "used it within a day of their first Excel MCP action.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": _analytics_feature_name(row["name"]),
                            "shown": f"{_analytics_cell(row['firstDayPct'])}% on day one "
                            f"({_analytics_cell(row['firstDay'])} of "
                            f"{_analytics_cell(row['users'])}); "
                            f"{_analytics_cell(row['later'])} after a week or more",
                        }
                        for row in habits["firstAdvancedUse"]
                    ],
                    label_field="friendlyName",
                    value_field="firstDayPct",
                    display_field="shown",
                ),
                "",
            ]
        )
    return sections


def _analytics_download_sections(
    downloads: dict[str, object], date_format: str
) -> list[str]:
    """Render public download counters from npm, NuGet, GitHub, and VS Code."""
    collected = datetime.fromisoformat(
        str(downloads["collectedUtc"]).replace("Z", "+00:00")
    )
    channel_rows = [
        {
            **row,
            "shown": f"{_analytics_cell(row['total'])} "
            + ("installs" if row["key"] == "vscode" else "downloads"),
        }
        for row in downloads["channels"]
    ]
    sections = [
        "## Where people get Excel MCP",
        "",
        "These numbers come from the public download counters on npm, NuGet, "
        "GitHub, and the Visual Studio Marketplace, checked on "
        f"**{collected.strftime(date_format)}**.",
        "",
        "!!! note \"Downloads are not people\"\n"
        "    One person can download Excel MCP many times, updates and automatic "
        "installs add to the counts, and the same person can appear in more than "
        "one place. Use these numbers to compare channels and spot trends, not to "
        "count users, and do not add them together.",
        "",
        "### Downloads by channel",
        "",
        "npm counts cover the last 12 months. The other counts include "
        "everything since each channel started.",
        "",
        _analytics_bar_chart(
            channel_rows,
            label_field="label",
            value_field="total",
            display_field="shown",
        ),
        "",
    ]
    npm_weekly = list(downloads["npmWeekly"])
    if npm_weekly:
        sections.extend(
            [
                "### npm downloads each week",
                "",
                "npm keeps a daily history, so each bar is one full week of "
                "downloads for the MCP Server and command line packages together.",
                "",
                _analytics_week_chart(
                    npm_weekly,
                    value_field="total",
                    title="npm downloads each week",
                ),
                "",
            ]
        )
    releases = list(downloads["releases"])
    if releases:
        sections.extend(
            [
                "### Downloads of recent releases",
                "",
                "Files downloaded from each GitHub release: the MCP Server and "
                "command line packages, the VS Code extension files, and the "
                "Claude Desktop bundle. Newer releases have had less time to "
                "collect downloads.",
                "",
                _analytics_bar_chart(
                    [
                        {
                            **row,
                            "friendlyName": f"{row['version']} ("
                            + datetime.fromisoformat(str(row["published"])).strftime(
                                "%b %d"
                            )
                            + ")",
                        }
                        for row in releases
                    ],
                    label_field="friendlyName",
                    value_field="downloads",
                ),
                "",
            ]
        )
    gains = list(downloads["weeklyGains"])
    sections.extend(["### New downloads each week", ""])
    if gains:
        sections.extend(
            [
                "NuGet, GitHub releases, and the VS Code Marketplace only publish "
                "running totals. This chart shows how much those totals grew "
                "between one weekly report and the next.",
                "",
                _analytics_week_chart(
                    [{**row, "total": max(0, int(row["total"]))} for row in gains],
                    value_field="total",
                    title="New NuGet, GitHub release, and VS Code downloads",
                ),
                "",
            ]
        )
    else:
        sections.extend(
            [
                "NuGet, GitHub releases, and the VS Code Marketplace only publish "
                "running totals. This report saves those totals every week, so a "
                "chart of new downloads each week appears from the next report on.",
                "",
            ]
        )
    return sections


def render_usage_analytics() -> str:
    source_rel = ".github/usage-analytics.json"
    report = json.loads(read(source_rel))
    schema_version = report.get("schemaVersion")
    if schema_version not in (2, 3):
        raise ValueError("usage analytics has an unsupported schema version")
    weighted = schema_version >= 3
    interpretation = report.get("interpretation")
    if not isinstance(interpretation, str) or not interpretation.strip():
        raise ValueError("usage analytics is missing its validated interpretation")

    summary = report["summary"]
    comparison = report["comparison"]
    generated = datetime.fromisoformat(report["generatedAtUtc"].replace("Z", "+00:00"))
    reporting_days = int(report["windows"]["reportingDays"])
    comparison_days = int(report["windows"]["comparisonDays"])
    reporting_start = generated - timedelta(days=reporting_days)
    current_start = generated - timedelta(days=comparison_days)
    previous_start = generated - timedelta(days=comparison_days * 2)
    date_format = "%b %d, %Y"

    hero_rows = [
        {
            **row,
            "friendlyName": _analytics_name(
                row["name"], _ANALYTICS_HERO_FEATURE_NAMES
            ),
        }
        for row in report["heroFeatures"]
    ]
    operation_rows = [
        {
            **row,
            "friendlyName": _analytics_name(row["name"], _ANALYTICS_OPERATION_NAMES),
        }
        for row in report["operations"]
    ]
    comparison_rows = [
        {
            "metric": "Users",
            "current": comparison["currentUsers"],
            "previous": comparison["previousUsers"],
            "change": f"{comparison['userChangePct']}%",
        },
        {
            "metric": "Actions",
            "current": comparison["currentInvocations"],
            "previous": comparison["previousInvocations"],
            "change": f"{comparison['invocationChangePct']}%",
        },
    ]
    if weighted:
        comparison_rows.append(
            {
                "metric": "Estimated work",
                "current": comparison["currentWorkUnits"],
                "previous": comparison["previousWorkUnits"],
                "change": f"{comparison['workChangePct']}%",
            }
        )
    sections = [
        "Excel MCP Server lets GitHub Copilot, Claude, and other AI assistants "
        "automate the real Microsoft Excel application. This public report shows "
        "how the open-source project is used and where people get it.",
        "",
        "New to the project? [Install Excel MCP Server](/installation/) to get started.",
        "",
        "!!! info \"Anonymous public report\"\n"
        "    This page shows broad usage patterns, not individual activity. "
        "Names, file details, locations, and the content of workbooks are never "
        "included.",
        "",
        f"**Last updated:** {generated.strftime(date_format)}  \n"
        f"**Period covered:** {reporting_start.strftime(date_format)} to "
        f"{generated.strftime(date_format)}",
        "",
        "## At a glance",
        "",
        '<div class="grid cards analytics-cards" markdown>',
        "",
        f"- :material-account-group: **{_analytics_cell(summary['users'])} users**",
        "",
        f"    Used Excel MCP during the last {reporting_days} days.",
        "",
        f"- :material-lightning-bolt: **{_analytics_cell(summary['toolInvocations'])} actions**",
        "",
        "    Recorded across workbooks, cells, data, charts, and automation.",
        "",
        f"- :material-calendar-refresh: **{_analytics_cell(summary['repeatUserRate'])}% returned**",
        "",
        "    Used Excel MCP on at least two different days.",
        "",
        "</div>",
        "",
        f"## Usage over the last {report['windows']['trendWeeks']} complete weeks",
        "",
        "Each bar is one full week, which makes changes easier to compare.",
        "",
        _analytics_week_chart(
            report["weekly"],
            value_field="users",
            title="Users each week",
        ),
        "",
        _analytics_week_chart(
            report["weekly"],
            value_field="actions",
            title="Actions each week",
        ),
        "",
        *(
            [
                _analytics_week_chart(
                    report["weekly"],
                    value_field="workUnits",
                    title="Estimated work each week",
                ),
                "",
            ]
            if weighted
            else []
        ),
        "## Release upgrades over time",
        "",
        "Each column is one week; the final column is the current week so far. A "
        "user appears once, under the latest release they used that week. This "
        "makes it easy to see newer releases replace older ones without overall "
        "user growth changing the scale. Less common releases are grouped as "
        "**Other**.",
        "",
        _analytics_version_chart(report["versionAdoption"]),
        "",
        "## The latest two weeks",
        "",
        f"The latest period is **{current_start.strftime(date_format)} to "
        f"{generated.strftime(date_format)}**. It is compared with "
        f"**{previous_start.strftime(date_format)} to "
        f"{current_start.strftime(date_format)}**.",
        "",
        _analytics_paired_bar_chart(
            [
                {
                    **row,
                    "shown": f"{_analytics_cell(row['previous'])} → "
                    f"{_analytics_cell(row['current'])} ({row['change']})",
                }
                for row in comparison_rows
            ],
            label_field="metric",
            first_field="previous",
            second_field="current",
            first_name=f"Previous {comparison_days} days",
            second_name=f"Latest {comparison_days} days",
            value_suffix="",
            display_field="shown",
            scale_each_row=True,
        ),
        "",
        "## What the numbers tell us",
        "",
        "!!! note \"Summary written by GitHub Copilot\"\n"
        "    Copilot reads only the anonymous totals used to build this page. Its "
        "summary is checked automatically so it cannot add private details or "
        "numbers that are not in the report.",
        "",
        interpretation.strip(),
        "",
        "## What people use most",
        "",
        *(_analytics_weighted_feature_section(report, hero_rows) if weighted else [
            "The bars group actions by the main features highlighted on the Excel MCP "
            "homepage. The percentage is each feature's share of meaningful actions. "
            "Smaller capabilities are grouped as **Other features**.",
            "",
            _analytics_bar_chart(
                hero_rows,
                label_field="friendlyName",
                value_field="sharePct",
                value_suffix="%",
            ),
            "",
        ]),
        "## Most common actions",
        "",
        _analytics_bar_chart(
            [
                {
                    **row,
                    "shown": f"{_analytics_cell(row['invocations'])} times by "
                    f"{_analytics_cell(row['users'])} users",
                }
                for row in operation_rows[:8]
            ],
            label_field="friendlyName",
            value_field="invocations",
            display_field="shown",
        ),
        "",
    ]
    if weighted:
        sections.extend(_analytics_work_sections(report))
        sections.extend(_analytics_entry_point_sections(report, date_format))
        if "habits" in report:
            sections.extend(_analytics_habit_sections(report["habits"]))
    if "downloads" in report:
        sections.extend(_analytics_download_sections(report["downloads"], date_format))
    sections.extend(
        [
            "## How this report protects privacy",
            "",
            "The report is built from anonymous counts and percentages. "
            "We do not publish or give Copilot user or session codes, file "
            "fingerprints, locations, messages, workbook content, error messages, "
            "or technical error details.",
            "",
            "Excel MCP never intentionally collects workbook contents, cell values, "
            "formulas, prompts, messages, file names or paths, names, email addresses, "
            "or account details. Read the full [privacy policy](/privacy/).",
            "",
            "You can inspect exactly how the report is built in "
            "[`Update-UsageAnalytics.ps1`](https://github.com/sbroenne/"
            "mcp-server-excel/blob/main/scripts/Update-UsageAnalytics.ps1) and "
            "[`usage-analytics.yml`](https://github.com/sbroenne/mcp-server-excel/"
            "blob/main/.github/workflows/usage-analytics.yml).",
        ]
    )
    return "\n".join(sections) + "\n"
