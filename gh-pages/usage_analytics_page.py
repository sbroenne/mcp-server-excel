"""Render the public usage-analytics page from ``.github/usage-analytics.json``.

Called by ``generate.py`` before the site build. Kept separate because it is a
report renderer, not documentation plumbing.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from html import escape

def _analytics_cell(value: object) -> str:
    """Format a validated aggregate value for a Markdown table."""
    if isinstance(value, float):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
    elif isinstance(value, int):
        text = f"{value:,}"
    else:
        text = str(value)
    return text.replace("|", r"\|").replace("\r", " ").replace("\n", " ")


def _analytics_table(
    headings: tuple[str, ...],
    fields: tuple[str, ...],
    rows: list[dict[str, object]],
) -> str:
    lines = [
        "| " + " | ".join(headings) + " |",
        "|" + "|".join("---" for _ in headings) + "|",
    ]
    for row in rows:
        lines.append(
            "| "
            + " | ".join(_analytics_cell(row[field]) for field in fields)
            + " |"
        )
    return "\n".join(lines)


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
}

_ANALYTICS_FAILURE_CLASS_NAMES = {
    "expected-negative": "Expected validation or state result",
    "input-state": "Input or workbook state",
    "external-dependency": "External dependency",
    "timeout-cancellation": "Timeout or cancellation",
    "excel-runtime": "Excel runtime",
    "internal-product-fault": "Excel MCP product fault",
    "unclassified": "Unclassified failure",
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
) -> str:
    """Render an accessible horizontal comparison chart."""
    maximum = max((float(row[value_field]) for row in rows), default=0)
    lines = ['<div class="analytics-bars" role="list">']
    for row in rows:
        value = float(row[value_field])
        width = 0 if maximum == 0 else max(2, value / maximum * 100)
        label = escape(str(row[label_field]))
        display_value = f"{_analytics_cell(row[value_field])}{value_suffix}"
        lines.extend(
            [
                '  <div class="analytics-bars__row" role="listitem">',
                '    <div class="analytics-bars__label">',
                f"      <span>{label}</span><strong>{escape(display_value)}</strong>",
                "    </div>",
                '    <div class="analytics-bars__track" aria-hidden="true">',
                f'      <span style="width: {width:.2f}%"></span>',
                "    </div>",
                "  </div>",
            ]
        )
    lines.append("</div>")
    return "\n".join(lines)


def _analytics_week_chart(
    rows: list[dict[str, object]],
    *,
    value_field: str,
    title: str,
) -> str:
    """Render weekly values as an accessible compact bar chart."""
    maximum = max((float(row[value_field]) for row in rows), default=0)
    midpoint = maximum / 2
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


def render(report: dict) -> str:
    if report.get("schemaVersion") != 2:
        raise ValueError("usage analytics has an unsupported schema version")
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
    categorized_reliability_since = datetime.fromisoformat(
        report["windows"]["categorizedReliabilitySinceUtc"].replace("Z", "+00:00")
    )
    categorized_reliability_version = report["windows"][
        "categorizedReliabilityMinimumVersion"
    ]
    date_format = "%b %d, %Y"

    hero_rows = [
        {
            **row,
            "friendlyName": _analytics_name(
                row["name"], _ANALYTICS_HERO_FEATURE_NAMES
            ),
            "share": f"{_analytics_cell(row['sharePct'])}%",
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
    reliability_rows = [
        {
            **row,
            "friendlyName": _analytics_name(row["name"], _ANALYTICS_OPERATION_NAMES),
            "failureRateDisplay": f"{_analytics_cell(row['failureRate'])}%",
        }
        for row in report["reliability"]
    ]
    failure_class_rows = [
        {
            **row,
            "friendlyName": _analytics_name(
                row["name"], _ANALYTICS_FAILURE_CLASS_NAMES
            ),
        }
        for row in report["failureClasses"]
    ]
    release_rows = [
        {
            **row,
            "failureRateDisplay": f"{_analytics_cell(row['failureRate'])}%",
        }
        for row in report["versionReliability"]
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
    sections = [
        "Excel MCP Server lets GitHub Copilot, Claude, and other AI assistants "
        "automate the real Microsoft Excel application. This public report shows "
        "how the open-source project is used and where reliability can improve.",
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
        _analytics_table(
            ("Measure", f"Latest {comparison_days} days", f"Previous {comparison_days} days", "Change"),
            ("metric", "current", "previous", "change"),
            comparison_rows,
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
        _analytics_table(
            ("Homepage feature", "Actions", "Users", "Share"),
            ("friendlyName", "invocations", "users", "share"),
            hero_rows,
        ),
        "",
        "## Most common actions",
        "",
        _analytics_table(
            ("Action", "Times used", "Users"),
            ("friendlyName", "invocations", "users"),
            operation_rows[:8],
        ),
        "",
    ]
    sections.extend(
        [
            "## Reliability measurement",
            "",
            f"Outcome classification starts with release "
            f"**{categorized_reliability_version}** from "
            f"**{categorized_reliability_since.strftime(date_format)}**. Earlier "
            "rows did not contain these labels and are not guessed or rewritten.",
            "",
            "A negative diagnostic result, such as finding that a workbook cannot "
            "be opened, is counted as an expected result rather than a product "
            "failure. Failures are grouped using fixed labels supplied by the "
            "software. Unknown labels remain visible as **Unclassified failure**.",
            "",
        ]
    )
    if failure_class_rows:
        sections.extend(
            [
                "### Outcomes and failure classes",
                "",
                _analytics_bar_chart(
                    failure_class_rows,
                    label_field="friendlyName",
                    value_field="actions",
                ),
                "",
                _analytics_table(
                    ("Outcome or failure class", "Actions", "Users"),
                    ("friendlyName", "actions", "users"),
                    failure_class_rows,
                ),
                "",
            ]
        )
    if reliability_rows:
        sections.extend(
            [
                "### Reliability by action",
                "",
                _analytics_table(
                    (
                        "Action",
                        "Actions",
                        "Expected negative",
                        "Failures",
                        "Failure rate",
                        "Users",
                    ),
                    (
                        "friendlyName",
                        "actions",
                        "expectedNegatives",
                        "failures",
                        "failureRateDisplay",
                        "users",
                    ),
                    reliability_rows[:15],
                ),
                "",
            ]
        )
    if release_rows:
        sections.extend(
            [
                "### Reliability by release",
                "",
                "This comparison can reveal a problem introduced in a release. It "
                "is not a direct quality score: different releases may be used for "
                "different kinds of work. The action count shows how much data each "
                "rate is based on.",
                "",
                _analytics_table(
                    (
                        "Release",
                        "Actions",
                        "Expected negative",
                        "Failures",
                        "Failure rate",
                        "Users",
                    ),
                    (
                        "version",
                        "actions",
                        "expectedNegatives",
                        "failures",
                        "failureRateDisplay",
                        "users",
                    ),
                    release_rows[:15],
                ),
                "",
            ]
        )
    sections.extend(
        [
        "## Problems we are watching",
        "",
        ]
    )
    exceptions = report["exceptions"]
    if exceptions:
        total_exceptions = sum(int(row["exceptions"]) for row in exceptions)
        sections.append(
            f"Excel MCP reported **{_analytics_cell(total_exceptions)} background "
            "task problems** during this period. These reports came from at least "
            f"**{_analytics_cell(max(int(row['users']) for row in exceptions))} "
            "users**. They are not the same as failed user actions, and one "
            "underlying problem can produce more than one report."
        )
    else:
        sections.append(
            "No broadly shared background problem appeared during this period."
        )
    sections.extend(
        [
            "",
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
