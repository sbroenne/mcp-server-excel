"""MkDocs build hook: generate documentation pages from canonical repo sources.

This preserves the project's single-source-of-truth design: several site pages
are generated from the authoritative Markdown files elsewhere in the repo
(README files, FEATURES.md, CHANGELOG.md, docs/*) so the website can never
drift from the real docs. It is the MkDocs equivalent of the old Jekyll
``build.sh`` script.

Generated files are written to ``gh-pages/_generated/`` (git-ignored, and
deliberately outside ``docs_dir``) and pulled into the thin wrapper pages under
``docs/`` via the ``pymdownx.snippets`` ``--8<--`` include syntax. Regeneration
happens automatically on every ``mkdocs build`` / ``mkdocs serve`` via the
``on_pre_build`` event.

Two smaller jobs live here as well:

* ``on_env`` hands ``overrides/sitemap.xml`` the git commit date behind every
  page, so ``<lastmod>`` reflects real content changes rather than the build
  date, plus the home page's video metadata.
* ``on_post_page`` gives Material's search dialog an accessible name. The logo
  and progress-bar equivalents are declarative partials under ``overrides/``;
  ``audit_site.py`` fails the build if any of the three stops applying.
"""

from __future__ import annotations

import json
import logging
import posixpath
import re
import subprocess
from datetime import datetime, timedelta
from html import escape
from pathlib import Path

log = logging.getLogger("mkdocs.hooks.generate")

# Home-page intro video. MkDocs' built-in sitemap is a plain URL sitemap and has
# no notion of embedded media, so overrides/sitemap.xml renders a Google
# video-sitemap <video:video> block into the home page's <url> entry. Keep these
# fields in sync with the VideoObject JSON-LD in docs/index.md.
VIDEO = {
    "page_url": "https://excelmcpserver.dev/",
    "thumbnail": "https://i.ytimg.com/vi/B6eIQ5BIbNc/maxresdefault.jpg",
    "title": "Introducing MCP Server for Excel - AI Coding for Excel",
    "description": (
        "See Excel MCP Server drive the real Microsoft Excel application from an "
        "AI assistant - Power Query, DAX, VBA, PivotTables and more."
    ),
    "player_loc": "https://www.youtube.com/embed/B6eIQ5BIbNc",
    "duration": "62",
    "publication_date": "2025-11-23T08:33:40-08:00",
}

# gh-pages/hooks.py -> gh-pages/ -> repo root
REPO_ROOT = Path(__file__).resolve().parent.parent
# Deliberately OUTSIDE docs_dir. When the generated files lived in
# docs/_generated/, every build rewrote files inside the directory `mkdocs
# serve` watches, so a single edit put the dev server into an endless
# rebuild loop. `.` is a snippets base_path, so the `--8<-- "_generated/..."`
# includes in the wrapper pages resolve here unchanged.
GEN_DIR = Path(__file__).resolve().parent / "_generated"

GITHUB_BLOB = "https://github.com/sbroenne/mcp-server-excel/blob/main/"
GITHUB_TREE = "https://github.com/sbroenne/mcp-server-excel/tree/main/"

# Repo-relative paths that have a dedicated site page: rewrite links to them so
# they resolve on the website instead of 404-ing.
SITE_PAGE_MAP = {
    "FEATURES.md": "/features/",
    ".github/usage-analytics.json": "/usage-analytics/",
    "docs/features/DATA-ANALYTICS.md": "/features/data-analytics/",
    "docs/features/CELLS-WORKBOOKS.md": "/features/cells-workbooks/",
    "docs/features/CHARTS-VISUALS.md": "/features/charts-visuals/",
    "docs/features/AUTOMATION-ADVANCED.md": "/features/automation-advanced/",
    "CHANGELOG.md": "/changelog/",
    "docs/INSTALLATION.md": "/installation/",
    "docs/INSTALLATION-MCP-SERVER.md": "/installation-mcp-server/",
    "docs/INSTALLATION-CLI.md": "/installation-cli/",
    "docs/ARCHITECTURE.md": "/architecture/",
    "docs/USE-CASES.md": "/use-cases/",
    "docs/guides/README.md": "/guides/",
    "docs/guides/REFRESH-POWER-QUERY.md": "/guides/refresh-power-query/",
    "docs/guides/AUTOMATE-PIVOTTABLES.md": "/guides/automate-pivottables/",
    "docs/guides/RUN-VBA-MACROS.md": "/guides/run-vba-macros/",
    "docs/guides/QUERY-DATA-MODEL-WITH-DAX.md": "/guides/query-data-model-with-dax/",
    "docs/guides/EXCEL-COM-VS-FILE-PARSERS.md": "/guides/excel-automation-vs-file-parsers/",
    "docs/CONTRIBUTING.md": "/contributing/",
    "SECURITY.md": "/security/",
    "PRIVACY.md": "/privacy/",
    "src/ExcelMcp.McpServer/README.md": "/mcp-server/",
    "src/ExcelMcp.CLI/README.md": "/cli/",
    "skills/README.md": "/skills/",
}

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")

SITE_URL = "https://excelmcpserver.dev/"

# Raw Markdown of every built page, captured in on_page_markdown with --8<--
# includes resolved, and emitted in on_post_build as /llms-full.txt plus one
# Markdown mirror per page. Keyed by the page's site path.
_PAGE_MARKDOWN: dict[str, dict[str, str]] = {}

_SNIPPET = re.compile(r'^[ \t]*(?:-{2,}8<-{2,})[ \t]+"([^"]+)"[ \t]*$', re.MULTILINE)
_FRONTMATTER = re.compile(r"\A---\r?\n.*?\r?\n---\r?\n", re.DOTALL)

FEATURE_SOURCES = {
    "features-data.md": "docs/features/DATA-ANALYTICS.md",
    "features-workbooks.md": "docs/features/CELLS-WORKBOOKS.md",
    "features-visualization.md": "docs/features/CHARTS-VISUALS.md",
    "features-automation.md": "docs/features/AUTOMATION-ADVANCED.md",
}

# Canonical task guides -> intent-focused website pages. Same contract as the
# feature references: the wrapper owns presentation and SEO metadata only.
GUIDE_SOURCES = {
    "guides-index.md": "docs/guides/README.md",
    "guides-refresh-power-query.md": "docs/guides/REFRESH-POWER-QUERY.md",
    "guides-automate-pivottables.md": "docs/guides/AUTOMATE-PIVOTTABLES.md",
    "guides-run-vba-macros.md": "docs/guides/RUN-VBA-MACROS.md",
    "guides-query-data-model-with-dax.md": "docs/guides/QUERY-DATA-MODEL-WITH-DAX.md",
    "guides-excel-com-vs-file-parsers.md": "docs/guides/EXCEL-COM-VS-FILE-PARSERS.md",
}


# skills/shared/*.md: the expert reference corpus shipped inside the skill
# packages and MCP prompts. Published verbatim so the site and the agent
# guidance can never disagree. Value = (output name, page title).
SKILL_SOURCES = {
    "workflows.md": ("skills-workflows.md", "Key Constraints & Sequencing"),
    "behavioral-rules.md": ("skills-behavioral-rules.md", "Behavioral Rules"),
    "anti-patterns.md": ("skills-anti-patterns.md", "Anti-Patterns to Avoid"),
    "gotchas.md": ("skills-gotchas.md", "Gotchas & Known Limits"),
    "excel_agent_mode.md": ("skills-agent-mode.md", "Agent Mode in Excel"),
    "workbook.md": ("skills-workbook.md", "Workbook Lifecycle"),
    "worksheet.md": ("skills-worksheet.md", "Worksheet Operations"),
    "range.md": ("skills-range.md", "Ranges, Number Formats & Formatting"),
    "table.md": ("skills-table.md", "Excel Tables"),
    "powerquery.md": ("skills-powerquery.md", "Power Query"),
    "m-code-syntax.md": ("skills-m-code-syntax.md", "M Code Syntax"),
    "datamodel.md": ("skills-datamodel.md", "Data Model & DAX"),
    "dmv-reference.md": ("skills-dmv-reference.md", "DMV Query Reference"),
    "pivottable.md": ("skills-pivottable.md", "PivotTables"),
    "querytable.md": ("skills-querytable.md", "QueryTables"),
    "analysis.md": ("skills-analysis.md", "What-If Analysis"),
    "chart.md": ("skills-chart.md", "Charts"),
    "conditionalformat.md": ("skills-conditionalformat.md", "Conditional Formatting"),
    "slicer.md": ("skills-slicer.md", "Slicers"),
    "drawing.md": ("skills-drawing.md", "Drawing Objects"),
    "screenshot.md": ("skills-screenshot.md", "Screenshots & Visual Verification"),
    "dashboard.md": ("skills-dashboard.md", "Dashboards & Reports"),
    "window.md": ("skills-window.md", "Window Management"),
    "xmlmap.md": ("skills-xmlmap.md", "XML Maps"),
}

_SKILL_SLUGS = {
    name: output.removeprefix("skills-").removesuffix(".md")
    for name, (output, _title) in SKILL_SOURCES.items()
}
SITE_PAGE_MAP.update(
    {f"skills/shared/{name}": f"/reference/{slug}/" for name, slug in _SKILL_SLUGS.items()}
)


def _rewrite_links(text: str, source_rel: str) -> str:
    """Resolve links in pulled-in content so they work on the site.

    Two cases:

    - Repo-relative links: rewritten to the published page when we publish one,
      otherwise to an absolute GitHub URL.
    - Absolute GitHub URLs into this repo: rewritten *back* to the published
      page when we publish one. Sources that are also rendered outside GitHub -
      the NuGet package READMEs - have to spell links out absolutely, because
      NuGet.org resolves relative links against the package root and they 404.
      Without this the website would link out to GitHub for pages it publishes
      itself.

    External links, anchors and site-absolute links are left alone.
    """
    source_dir = posixpath.dirname(source_rel)

    def repl(match: re.Match) -> str:
        label, url = match.group(1), match.group(2)

        for prefix in (GITHUB_BLOB, GITHUB_TREE):
            if url.startswith(prefix):
                remainder = url[len(prefix) :]
                target, _, anchor = remainder.partition("#")
                anchor = f"#{anchor}" if anchor else ""
                if target.rstrip("/") in SITE_PAGE_MAP:
                    return f"[{label}]({SITE_PAGE_MAP[target.rstrip('/')]}{anchor})"
                return match.group(0)

        if url.startswith(("http://", "https://", "#", "/", "mailto:", "<")):
            return match.group(0)

        anchor = ""
        target = url
        if "#" in target:
            target, anchor = target.split("#", 1)
            anchor = "#" + anchor
        if target == "":
            return match.group(0)  # pure in-page anchor

        resolved = posixpath.normpath(posixpath.join(source_dir, target))
        if resolved.startswith(".."):
            return match.group(0)  # points outside the repo; leave as-is

        if resolved in SITE_PAGE_MAP:
            return f"[{label}]({SITE_PAGE_MAP[resolved]}{anchor})"

        base = GITHUB_TREE if url.endswith("/") else GITHUB_BLOB
        return f"[{label}]({base}{resolved}{anchor})"

    return _MD_LINK.sub(repl, text)


def _strip_header(
    text: str,
    *,
    drop_prefixes: tuple[str, ...] = (),
    end_on_blank: bool = False,
    end_on_hr: bool = False,
    demote_h1: bool = False,
) -> str:
    """Drop the leading H1 title block from a source file, optionally demoting
    any remaining H1 headings to H2.

    Mirrors the awk transforms in the previous Jekyll ``build.sh``:
    - the first ``# Title`` line is always dropped, and header mode begins;
    - while in the header, lines starting with any ``drop_prefixes`` are dropped;
    - the header ends on the first blank line (``end_on_blank``) or ``---`` rule
      (``end_on_hr``); leading blank lines before content are also dropped;
    - when ``demote_h1`` is set, any later ``# `` heading becomes ``## ``.
    """
    in_header = False
    header_done = False
    out: list[str] = []

    for line in text.splitlines():
        if not header_done and line.startswith("# "):
            in_header = True
            continue
        if in_header:
            if any(line.startswith(p) for p in drop_prefixes):
                continue
            if end_on_hr and line.startswith("---"):
                in_header = False
                header_done = True
                continue
            if line.strip() == "":
                if end_on_blank:
                    in_header = False
                    header_done = True
                continue
            # Any other lingering header line is dropped.
            continue
        if not header_done and line.strip() == "":
            # Skip leading blank lines before real content begins.
            continue
        header_done = True
        if demote_h1 and line.startswith("# "):
            line = "#" + line  # "# " -> "## "
        out.append(line)

    return "\n".join(out).strip() + "\n"


def _add_stable_feature_anchors(text: str) -> str:
    """Give feature headings stable IDs that do not include operation counts."""
    heading = re.compile(r"^## (?P<title>.+?) \(\d+ operations\)$", re.MULTILINE)

    def replace(match: re.Match) -> str:
        title = match.group("title")
        slug = re.sub(r"[^\w\s-]", "", title, flags=re.UNICODE).strip().lower()
        slug = re.sub(r"[-\s]+", "-", slug)
        return f"{match.group(0)} {{ #{slug} }}"

    return heading.sub(replace, text)


def _read(rel: str) -> str:
    path = REPO_ROOT / rel
    if not path.is_file():
        raise FileNotFoundError(f"Source doc not found: {path}")
    return path.read_text(encoding="utf-8")


def _write(name: str, source_rel: str, content: str) -> None:
    GEN_DIR.mkdir(parents=True, exist_ok=True)
    content = _rewrite_links(content, source_rel)
    (GEN_DIR / name).write_text(content, encoding="utf-8")
    MIRROR_SOURCES[name] = source_rel
    log.info("generated _generated/%s", name)


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


def _render_usage_analytics() -> str:
    source_rel = ".github/usage-analytics.json"
    report = json.loads(_read(source_rel))
    if report.get("schemaVersion") != 1:
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
    reliability_since = datetime.fromisoformat(
        report["windows"]["reliabilitySinceUtc"].replace("Z", "+00:00")
    )
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
            "errorRateDisplay": f"{_analytics_cell(row['errorRate'])}%",
        }
        for row in report["reliability"]
    ]
    release_rows = [
        {
            **row,
            "errorRateDisplay": f"{_analytics_cell(row['errorRate'])}%",
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
    if reliability_rows:
        sections.extend(
            [
                "## Actions reporting errors",
                "",
                f"Accurate error counting began on "
                f"**{reliability_since.strftime(date_format)}** with the latest "
                "patch release. Earlier releases are excluded because they did not "
                "count every kind of failed action.",
                "",
                _analytics_table(
                    ("Action", "Actions", "Errors", "Error rate", "Users"),
                    (
                        "friendlyName",
                        "actions",
                        "errors",
                        "errorRateDisplay",
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
                "## Errors by release",
                "",
                "This comparison can reveal a problem introduced in a release. It "
                "is not a direct quality score: different releases may be used for "
                "different kinds of work. The action count shows how much data each "
                "rate is based on.",
                "",
                _analytics_table(
                    ("Release", "Actions", "Errors", "Error rate", "Users"),
                    ("version", "actions", "errors", "errorRateDisplay", "users"),
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


# Generated file name (e.g. "features-data.md") -> repo-relative canonical
# source. Populated by _write during on_pre_build and read back when dating
# sitemap entries: a wrapper page's real "last modified" is driven by the
# canonical file it mirrors, not by the two-line wrapper.
MIRROR_SOURCES: dict[str, str] = {}

# Matches the snippet includes in the wrapper pages, e.g.
#     --8<-- "_generated/features-data.md"
_GEN_INCLUDE = re.compile(r'--8<--\s*"_generated/([^"]+)"')


def _git_lastmod_index() -> dict[str, str]:
    """Map every tracked repo-relative path to its last commit date (W3C).

    One ``git log`` walk over the whole history, newest first: the first time a
    path appears is by definition its most recent change. This replaces the
    previous behaviour of stripping ``<lastmod>`` altogether, which was done
    because MkDocs stamps every URL with the *build* date - a false freshness
    signal on every page in every deploy.

    Returns an empty index (so ``<lastmod>`` is simply omitted) when git is
    unavailable, which keeps ``mkdocs build`` working from a source tarball.
    """
    try:
        proc = subprocess.run(
            ["git", "log", "--format=%cI", "--name-only", "--no-renames"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=True,
        )
    except (OSError, subprocess.CalledProcessError) as exc:
        log.warning("git log failed (%s); sitemap will omit <lastmod>", exc)
        return {}

    index: dict[str, str] = {}
    date = ""
    for line in proc.stdout.splitlines():
        if not line:
            continue
        # Commit-date lines are the only ones that can start with a 4-digit year
        # followed by '-'; paths in this repo never do.
        if len(line) >= 5 and line[:4].isdigit() and line[4] == "-":
            date = line
        elif date:
            index.setdefault(line, date)
    return index


def _git_is_shallow() -> bool:
    """True when the checkout has truncated history.

    Worth reporting explicitly: a shallow clone still lists every tracked file,
    just all under the tip commit's date, so the lastmod index looks perfectly
    healthy while every date in it is wrong.
    """
    try:
        proc = subprocess.run(
            ["git", "rev-parse", "--is-shallow-repository"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return False
    return proc.stdout.strip() == "true"


def _page_lastmod(files) -> dict[str, str]:
    """Map each page's ``src_uri`` to the newest git date that affects it.

    For a wrapper page that is nothing but an ``--8<--`` include, that is the
    date of the canonical source; the wrapper itself contributes its own date
    too, so editing either one refreshes the entry.
    """
    index = _git_lastmod_index()
    if not index:
        return {}
    if _git_is_shallow():
        # A shallow clone (actions/checkout's default fetch-depth: 1) still lists
        # every tracked file - all under the tip commit's date. So the index
        # looks healthy and only the dates are wrong; audit_site.py catches it by
        # noticing that every page claims the same <lastmod>.
        log.warning(
            "shallow git clone: every sitemap <lastmod> will be the tip "
            "commit's date - the workflow needs fetch-depth: 0"
        )

    lastmod: dict[str, str] = {}
    for file in files.documentation_pages():
        candidates = []
        wrapper_rel = f"gh-pages/docs/{file.src_uri}"
        if wrapper_rel in index:
            candidates.append(index[wrapper_rel])
        try:
            text = Path(file.abs_src_path).read_text(encoding="utf-8")
        except OSError:
            text = ""
        for name in _GEN_INCLUDE.findall(text):
            source_rel = MIRROR_SOURCES.get(name)
            if source_rel and source_rel in index:
                candidates.append(index[source_rel])
        if candidates:
            # git's %cI keeps each committer's UTC offset, so the strings are
            # not directly comparable as instants - parse before taking the max.
            lastmod[file.src_uri] = max(candidates, key=datetime.fromisoformat)
    return lastmod


def on_env(env, config, files, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Expose sitemap data to overrides/sitemap.xml."""
    env.globals["page_lastmod"] = _page_lastmod(files)
    env.globals["video"] = VIDEO
    return env


DOCS_DIR = Path(__file__).resolve().parent / "docs"
# Mirrors the snippets `base_path` in mkdocs.yml, in the same order. Kept in
# sync so the llms.txt/mirror output resolves exactly what the site renders.
SNIPPET_BASE_PATHS = (DOCS_DIR, Path(__file__).resolve().parent)


def _resolve_snippets(text: str, depth: int = 0) -> str:
    """Expand ``--8<-- "path"`` includes.

    ``on_page_markdown`` fires before the snippets extension runs, so the raw
    Markdown still contains include directives. Resolving them here is what makes
    the Markdown mirrors and ``llms-full.txt`` complete rather than a list of
    stub pages.
    """
    if depth > 5:
        return text

    def repl(match: re.Match) -> str:
        for base in SNIPPET_BASE_PATHS:
            target = base / match.group(1)
            if target.is_file():
                return _resolve_snippets(target.read_text(encoding="utf-8"), depth + 1)
        log.warning("snippet not found while building llms output: %s", match.group(1))
        return ""

    return _SNIPPET.sub(repl, text)


def _page_url(page) -> str:
    return SITE_URL + page.url


# The resolved Navigation object, captured in on_nav. config["nav"] holds the raw
# YAML nav, which has no page objects to correlate with captured Markdown.
_NAV: list = []


def on_nav(nav, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _NAV.clear()
    _NAV.extend(nav.items)
    return nav


def on_page_markdown(markdown, page, config, **kwargs):  # noqa: D401 - MkDocs hook
    """Capture each page's full Markdown for the LLM-facing outputs."""
    body = _resolve_snippets(_FRONTMATTER.sub("", markdown)).strip()
    _PAGE_MARKDOWN[page.file.src_uri] = {
        "title": page.title or page.file.src_uri,
        "url": _page_url(page),
        "description": (page.meta or {}).get("description", "").strip(),
        "markdown": body,
        "dest": page.file.dest_uri,
    }

    faq = _faq_jsonld(body)
    if faq:
        page.meta["faq_jsonld"] = faq
    return markdown


_FAQ_ADMONITION = re.compile(r'^\?{3}\+?\s+question\s+"([^"]+)"\s*$')
_FAQ_HEADING = re.compile(r"^###\s+(.+?)\s*$")
_FAQ_MIN_ENTITIES = 3


def _faq_jsonld(markdown: str) -> str:
    """Build FAQPage JSON-LD from a page's own question blocks.

    Two source forms are recognised:

    * ``### Some question?`` headings - preferred, because each answer keeps a
      stable anchor that can be deep-linked from another page or straight from a
      search result, and shows up in the page table of contents.
    * ``??? question "..."`` collapsible admonitions, which have no anchor at
      all, kept so a page written either way still works.

    Either way the structured data is derived from the page body rather than
    maintained separately, so the two cannot diverge.
    """
    items: list[tuple[str, list[str]]] = []
    current: list[str] | None = None
    indented = False

    for line in markdown.splitlines():
        admonition = _FAQ_ADMONITION.match(line)
        if admonition:
            current = []
            indented = True
            items.append((admonition.group(1), current))
            continue

        heading = _FAQ_HEADING.match(line)
        if heading:
            text = heading.group(1).strip()
            if text.endswith("?"):
                current = []
                indented = False
                items.append((text, current))
            else:
                current = None
            continue

        if current is None:
            continue

        # A heading of any level ends a heading-sourced answer.
        if not indented and line.startswith("#"):
            current = None
            continue

        if not line.strip():
            current.append("")
        elif indented and not line.startswith((" ", "\t")):
            current = None
        else:
            current.append(line.strip())

    entities = []
    for question, answer_lines in items:
        # Fenced code blocks and table rows are useful on the page but pure noise
        # inside a structured answer, so they are dropped here.
        prose: list[str] = []
        in_fence = False
        for raw in answer_lines:
            if raw.startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence or raw.startswith("|"):
                continue
            # Strip the list marker only where it starts a line, so a dash used
            # mid-sentence survives into the structured answer.
            prose.append(re.sub(r"^[-*+]\s+", "", raw))

        answer = " ".join(x for x in prose if x).strip()
        if not answer:
            continue
        # Strip inline Markdown so the structured answer is plain prose.
        answer = _MD_LINK.sub(r"\1", answer)
        answer = re.sub(r"[*_`]+", "", answer)
        answer = re.sub(r"\s{2,}", " ", answer).strip()
        entities.append(
            {
                "@type": "Question",
                "name": question,
                "acceptedAnswer": {"@type": "Answer", "text": answer},
            }
        )

    # A page with one or two question-shaped headings is a guide that happens to
    # ask a question, not an FAQ; emitting FAQPage there is a false signal.
    if len(entities) < _FAQ_MIN_ENTITIES:
        return ""

    return json.dumps(
        {"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": entities},
        ensure_ascii=False,
    )


def _nav_entries(items, out: list) -> None:
    for item in items:
        if getattr(item, "children", None):
            _nav_entries(item.children, out)
        elif getattr(item, "file", None) is not None:
            out.append(item)


def _write_llm_outputs(config) -> None:
    """Emit /llms.txt, /llms-full.txt and one Markdown mirror per page.

    ``llms.txt`` follows the llmstxt.org convention: an H1, a blockquote summary,
    then link sections. Both files and the mirrors are derived from the same
    captured Markdown, so they cannot drift from the site.
    """
    site_dir = Path(config["site_dir"])

    # Markdown mirrors: /guides/refresh-power-query/index.md next to index.html.
    mirrored = 0
    for entry in _PAGE_MARKDOWN.values():
        dest = site_dir / entry["dest"]
        if dest.suffix != ".html":
            continue
        md_path = dest.with_suffix(".md")
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text(
            entry["markdown"] + "\n",
            encoding="utf-8",
            newline="\n",
        )
        mirrored += 1

    # Section-aware index, ordered exactly like the site navigation.
    lines = [
        "# Excel MCP Server",
        "",
        "> Excel MCP Server (ExcelMcp) automates the real Microsoft Excel "
        "application through its COM API, exposing 31 tools and 325 operations to AI assistants "
        "over the Model Context Protocol and to scripts through "
        "the `excelcli` command line. Unlike file-parser libraries it can refresh "
        "Power Query, evaluate DAX against the Data Model, refresh PivotTables, "
        "and run VBA, because Excel itself does the work. Windows-only; requires "
        "Microsoft Excel 2016 or later.",
        "",
        "Every page below is also available as Markdown by appending `index.md` "
        "to its URL. The complete corpus is at "
        f"{SITE_URL}llms-full.txt.",
        "",
    ]

    def link_line(entry: dict) -> str:
        url = entry["url"].rstrip("/")
        url = f"{url}/index.md" if entry["dest"].endswith("index.html") else url
        desc = f": {entry['description']}" if entry["description"] else ""
        return f"- [{entry['title']}]({url}){desc}"

    seen: set[str] = set()
    for section in _NAV:
        pages: list = []
        _nav_entries([section], pages)
        title = section.title if getattr(section, "title", None) else "Documentation"
        rendered = []
        for item in pages:
            entry = _PAGE_MARKDOWN.get(item.file.src_uri)
            if entry is None or item.file.src_uri in seen:
                continue
            seen.add(item.file.src_uri)
            rendered.append(link_line(entry))
        if rendered:
            lines.append(f"## {title}")
            lines.append("")
            lines.extend(rendered)
            lines.append("")

    (site_dir / "llms.txt").write_text("\n".join(lines), encoding="utf-8", newline="\n")

    # Full corpus, same order as llms.txt.
    full = ["# Excel MCP Server - complete documentation", ""]
    ordered: list = []
    _nav_entries(_NAV, ordered)
    emitted: set[str] = set()
    for item in ordered:
        entry = _PAGE_MARKDOWN.get(item.file.src_uri)
        if entry is None or item.file.src_uri in emitted:
            continue
        emitted.add(item.file.src_uri)
        full.append(f"# {entry['title']}")
        full.append("")
        full.append(f"Source: {entry['url']}")
        full.append("")
        full.append(entry["markdown"])
        full.append("")
        full.append("---")
        full.append("")
    (site_dir / "llms-full.txt").write_text(
        "\n".join(full), encoding="utf-8", newline="\n"
    )

    log.info(
        "wrote llms.txt, llms-full.txt and %d Markdown mirrors", mirrored
    )



def on_pre_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    _write(
        "usage-analytics.md",
        ".github/usage-analytics.json",
        _render_usage_analytics(),
    )

    # Canonical feature references -> focused website pages. The wrappers add
    # presentation and SEO metadata but never duplicate operation details.
    for output_name, source_rel in FEATURE_SOURCES.items():
        _write(
            output_name,
            source_rel,
            _add_stable_feature_anchors(
                _strip_header(_read(source_rel), end_on_hr=True)
            ),
        )

    # Canonical task guides -> intent-focused website pages. The H1 lives in the
    # wrapper, so drop it here and demote any remaining H1 to H2.
    for output_name, source_rel in GUIDE_SOURCES.items():
        _write(
            output_name,
            source_rel,
            _strip_header(_read(source_rel), end_on_blank=True, demote_h1=True),
        )

    # CHANGELOG.md -> changelog (drop title + description line, demote H1)
    _write(
        "changelog.md",
        "CHANGELOG.md",
        _strip_header(
            _read("CHANGELOG.md"),
            drop_prefixes=("This changelog",),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # docs/INSTALLATION.md -> installation (drop title + description line, demote H1)
    _write(
        "installation.md",
        "docs/INSTALLATION.md",
        _strip_header(
            _read("docs/INSTALLATION.md"),
            drop_prefixes=("Complete installation",),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # docs/INSTALLATION-MCP-SERVER.md -> installation-mcp-server (drop title + description line, demote H1)
    _write(
        "installation-mcp-server.md",
        "docs/INSTALLATION-MCP-SERVER.md",
        _strip_header(
            _read("docs/INSTALLATION-MCP-SERVER.md"),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # docs/INSTALLATION-CLI.md -> installation-cli (drop title + description line, demote H1)
    _write(
        "installation-cli.md",
        "docs/INSTALLATION-CLI.md",
        _strip_header(
            _read("docs/INSTALLATION-CLI.md"),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # Canonical architecture and examples guides.
    _write(
        "architecture.md",
        "docs/ARCHITECTURE.md",
        _strip_header(_read("docs/ARCHITECTURE.md"), end_on_blank=True),
    )
    _write(
        "use-cases.md",
        "docs/USE-CASES.md",
        _strip_header(_read("docs/USE-CASES.md"), end_on_blank=True),
    )

    # src/ExcelMcp.McpServer/README.md -> mcp-server (drop title, mcp-name, badges)
    _write(
        "mcp-server.md",
        "src/ExcelMcp.McpServer/README.md",
        _strip_header(
            _read("src/ExcelMcp.McpServer/README.md"),
            drop_prefixes=("<!-- mcp-name", "mcp-name:", "[!["),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # src/ExcelMcp.CLI/README.md -> cli (drop title + badges, demote H1)
    _write(
        "cli.md",
        "src/ExcelMcp.CLI/README.md",
        _strip_header(
            _read("src/ExcelMcp.CLI/README.md"),
            drop_prefixes=("[![",),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # skills/README.md -> skills (drop title, demote H1)
    _write(
        "skills.md",
        "skills/README.md",
        _strip_header(
            _read("skills/README.md"),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # skills/shared/*.md -> reference pages (drop the H1, wrapper owns the title)
    for name, (output_name, _title) in SKILL_SOURCES.items():
        _write(
            output_name,
            f"skills/shared/{name}",
            _strip_header(
                _read(f"skills/shared/{name}"), end_on_blank=True, demote_h1=True
            ),
        )

    # Verbatim copies (these keep their own H1 as the page title).
    _write("contributing.md", "docs/CONTRIBUTING.md", _read("docs/CONTRIBUTING.md").strip() + "\n")
    _write("security.md", "SECURITY.md", _read("SECURITY.md").strip() + "\n")
    _write("privacy.md", "PRIVACY.md", _read("PRIVACY.md").strip() + "\n")


def _write_tools_json(config) -> None:
    """Emit /tools.json: every tool and operation as structured JSON.

    Derived from the canonical ``docs/features/*.md`` references, so the machine
    -readable catalogue is generated from the same source as the human pages and
    cannot drift. Totals are asserted against the documented headline counts.
    """
    category_titles = {
        "docs/features/DATA-ANALYTICS.md": "Data & Analytics",
        "docs/features/CELLS-WORKBOOKS.md": "Cells & Workbooks",
        "docs/features/CHARTS-VISUALS.md": "Charts & Visualization",
        "docs/features/AUTOMATION-ADVANCED.md": "Automation & Advanced",
    }
    site_page = {
        "docs/features/DATA-ANALYTICS.md": "/features/data-analytics/",
        "docs/features/CELLS-WORKBOOKS.md": "/features/cells-workbooks/",
        "docs/features/CHARTS-VISUALS.md": "/features/charts-visuals/",
        "docs/features/AUTOMATION-ADVANCED.md": "/features/automation-advanced/",
    }

    heading = re.compile(r"^## (?:\W+\s+)?(?P<name>.+?) \((?P<count>\d+) operations\)$")
    operation = re.compile(r"^- \*\*(?P<name>[^:*]+):\*\*\s*(?P<desc>.+)$")

    # Headline counts live in FEATURES.md and are enforced against code by
    # scripts/check-doc-counts.ps1, so read them rather than restating them.
    headline = re.search(
        r"\*\*(?P<tools>\d+) specialized tools with (?P<ops>\d+) operations",
        _read("FEATURES.md"),
    )
    if headline is None:
        raise RuntimeError("could not read the headline tool/operation counts from FEATURES.md")
    headline_tools = int(headline.group("tools"))
    headline_ops = int(headline.group("ops"))

    categories = []
    total_ops = 0

    for source_rel, title in category_titles.items():
        groups: list[dict] = []
        current: dict | None = None
        for line in _read(source_rel).splitlines():
            match = heading.match(line)
            if match:
                current = {
                    "name": match.group("name").strip(),
                    "operationCount": int(match.group("count")),
                    "operations": [],
                }
                groups.append(current)
                continue
            if current is None:
                continue
            op = operation.match(line)
            if op:
                current["operations"].append(
                    {
                        "name": op.group("name").strip(),
                        "description": op.group("desc").strip(),
                    }
                )

        total_ops += sum(g["operationCount"] for g in groups)
        categories.append(
            {
                "name": title,
                "url": SITE_URL.rstrip("/") + site_page[source_rel],
                "operationCount": sum(g["operationCount"] for g in groups),
                "featureGroups": groups,
            }
        )

    if total_ops != headline_ops:
        raise RuntimeError(
            "tools.json operation total does not match the FEATURES.md headline: "
            f"parsed {total_ops}, expected {headline_ops}. "
            "Fix the feature reference headings or the headline."
        )

    payload = {
        "name": "Excel MCP Server",
        "url": SITE_URL,
        "repository": "https://github.com/sbroenne/mcp-server-excel",
        "description": (
            "Automates the real Microsoft Excel application through its COM API, "
            "exposing Excel to AI assistants over the Model Context Protocol and "
            "to scripts through the excelcli command line."
        ),
        "requirements": {
            "operatingSystem": "Windows",
            "application": "Microsoft Excel desktop 2016 or later",
        },
        "entryPoints": ["mcp-server", "cli"],
        "toolCount": headline_tools,
        "operationCount": total_ops,
        "categories": categories,
    }

    (Path(config["site_dir"]) / "tools.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    log.info("wrote tools.json (%d tools, %d operations)", headline_tools, total_ops)


def on_post_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Write the LLM-facing outputs that MkDocs itself has no notion of.

    The sitemap used to be rewritten here - stripping ``<lastmod>`` and splicing
    in a ``<video:video>`` block with regexes, then re-gzipping by hand. Both
    jobs now happen declaratively in ``overrides/sitemap.xml``, which also means
    MkDocs writes ``sitemap.xml.gz`` from the same rendered output instead of the
    two being kept in step manually.
    """
    _write_llm_outputs(config)
    _write_tools_json(config)


def on_post_page(output, page, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Give Material's search dialog an accessible name.

    A role="dialog" with no name is a WCAG 4.1.2 failure. Unlike the logo and
    progress-bar fixes - now declarative partials under ``overrides/`` - this one
    stays a string patch on purpose: upstream's ``partials/search.html`` is ~45
    lines of markup, icon lookups and feature flags, so copying it into
    ``overrides/`` to add one attribute would pin a large slice of Material
    internals and silently miss every upstream change to the search UI.

    Two variants because mkdocs-minify strips attribute quotes.
    """
    output = output.replace(
        '<div class="md-search" data-md-component="search" role="dialog">',
        '<div class="md-search" data-md-component="search" role="dialog" '
        'aria-label="Search documentation">',
    )
    output = output.replace(
        "<div class=md-search data-md-component=search role=dialog>",
        '<div class=md-search data-md-component=search role=dialog '
        'aria-label="Search documentation">',
    )
    return output
