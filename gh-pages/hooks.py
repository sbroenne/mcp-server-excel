"""MkDocs build hook: generate documentation pages from canonical repo sources.

This preserves the project's single-source-of-truth design: several site pages
are generated from the authoritative Markdown files elsewhere in the repo
(README files, FEATURES.md, CHANGELOG.md, docs/*) so the website can never
drift from the real docs. It is the MkDocs equivalent of the old Jekyll
``build.sh`` script.

Generated files are written to ``docs/_generated/`` (git-ignored) and pulled
into the thin wrapper pages under ``docs/`` via the ``pymdownx.snippets``
``--8<--`` include syntax. Regeneration happens automatically on every
``mkdocs build`` / ``mkdocs serve`` via the ``on_pre_build`` event.
"""

from __future__ import annotations

import gzip
import json
import logging
import posixpath
import re
from pathlib import Path

log = logging.getLogger("mkdocs.hooks.generate")

# Home-page intro video. MkDocs' built-in sitemap is a plain URL sitemap and has
# no notion of embedded media, so we enrich the home page's <url> entry with a
# Google video-sitemap <video:video> block in on_post_build. Keep these fields in
# sync with the VideoObject JSON-LD in docs/index.md.
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

_VIDEO_NS = "http://www.google.com/schemas/sitemap-video/1.1"


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )

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
    """Resolve repo-relative links in pulled-in content so they work on the site.

    Links that point at a page we publish are rewritten to that page's URL;
    everything else that resolves inside the repo is rewritten to an absolute
    GitHub URL. External links, anchors and site-absolute links are left alone.
    """
    source_dir = posixpath.dirname(source_rel)

    def repl(match: re.Match) -> str:
        label, url = match.group(1), match.group(2)
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
    log.info("generated _generated/%s", name)


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


_FAQ_QUESTION = re.compile(r'^\?{3}\+?\s+question\s+"([^"]+)"\s*$')


def _faq_jsonld(markdown: str) -> str:
    """Build FAQPage JSON-LD from ``??? question "..."`` admonitions.

    Derived from the page body rather than maintained separately, so the
    structured data and the visible FAQ cannot diverge.
    """
    items: list[tuple[str, list[str]]] = []
    current: list[str] | None = None

    for line in markdown.splitlines():
        match = _FAQ_QUESTION.match(line)
        if match:
            current = []
            items.append((match.group(1), current))
            continue
        if current is None:
            continue
        if not line.strip():
            current.append("")
        elif line.startswith((" ", "\t")):
            current.append(line.strip())
        else:
            current = None

    entities = []
    for question, answer_lines in items:
        answer = " ".join(x for x in answer_lines if x).strip()
        if not answer:
            continue
        # Strip inline Markdown so the structured answer is plain prose.
        answer = _MD_LINK.sub(r"\1", answer)
        answer = re.sub(r"[*_`]+", "", answer)
        entities.append(
            {
                "@type": "Question",
                "name": question,
                "acceptedAnswer": {"@type": "Answer", "text": answer},
            }
        )

    if not entities:
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
        "application through its COM API, exposing 31 tools and 326 operations "
        "to AI assistants over the Model Context Protocol and to scripts through "
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
    """Normalize and enrich the generated sitemap.

    MkDocs writes a plain URL sitemap that cannot describe the home-page intro
    video, so we add the ``video`` namespace to ``<urlset>`` and inject a
    ``<video:video>`` block into the home page's ``<url>``. Both ``sitemap.xml``
    and its gzipped twin are updated so Search Console reads the enriched copy.
    MkDocs also stamps every URL with the build date, even when its content did
    not change. Those unreliable ``lastmod`` values are removed rather than
    sending search engines a false freshness signal.
    """
    site_dir = Path(config["site_dir"])
    _write_llm_outputs(config)
    _write_tools_json(config)

    sitemap = site_dir / "sitemap.xml"
    if not sitemap.is_file():
        log.warning("sitemap.xml not found; skipping video-sitemap enrichment")
        return

    xml = sitemap.read_text(encoding="utf-8")
    xml = re.sub(r"\s*<lastmod>[^<]+</lastmod>", "", xml)

    if 'xmlns:video=' not in xml:
        xml = xml.replace(
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"'
            f' xmlns:video="{_VIDEO_NS}">',
            1,
        )

    video_block = (
        "        <video:video>\n"
        f"            <video:thumbnail_loc>{_xml_escape(VIDEO['thumbnail'])}</video:thumbnail_loc>\n"
        f"            <video:title>{_xml_escape(VIDEO['title'])}</video:title>\n"
        f"            <video:description>{_xml_escape(VIDEO['description'])}</video:description>\n"
        f"            <video:player_loc>{_xml_escape(VIDEO['player_loc'])}</video:player_loc>\n"
        f"            <video:duration>{VIDEO['duration']}</video:duration>\n"
        f"            <video:publication_date>{VIDEO['publication_date']}</video:publication_date>\n"
        "            <video:family_friendly>yes</video:family_friendly>\n"
        "            <video:live>no</video:live>\n"
        "        </video:video>\n"
    )

    # Insert the video block inside the home page's <url>...</url> element.
    home_url = re.compile(
        r"(<url>\s*<loc>"
        + re.escape(VIDEO["page_url"])
        + r"</loc>.*?)(</url>)",
        re.DOTALL,
    )
    if "<video:video>" not in xml:
        new_xml, count = home_url.subn(rf"\1{video_block}    \2", xml, count=1)
        if count:
            xml = new_xml
        else:
            log.warning(
                "home page <url> not found in sitemap; video markup not added"
            )

    sitemap.write_text(xml, encoding="utf-8", newline="\n")

    gz = site_dir / "sitemap.xml.gz"
    if gz.exists():
        # mtime=0 keeps the output reproducible, matching MkDocs' own gzip call.
        with gzip.GzipFile(gz, "wb", mtime=0) as fh:
            fh.write(xml.encode("utf-8"))

    log.info("normalized sitemap dates and added home-page video markup")


def on_post_page(output, page, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Add accessibility metadata omitted by the upstream Material partials."""
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
    output = output.replace(
        '<div class="md-progress" data-md-component="progress" role="progressbar">',
        '<div class="md-progress" data-md-component="progress" role="progressbar" '
        'aria-label="Page loading progress">',
    )
    output = output.replace(
        "<div class=md-progress data-md-component=progress role=progressbar>",
        '<div class=md-progress data-md-component=progress role=progressbar '
        'aria-label="Page loading progress">',
    )
    output = re.sub(
        r'<img src="([^"]*assets/images/logo\.png)" alt="logo">',
        r'<img src="\1" alt="Excel MCP Server" width="256" height="256">',
        output,
    )
    output = re.sub(
        r"<img src=([^\s>]*assets/images/logo\.png) alt=logo>",
        r'<img src="\1" alt="Excel MCP Server" width="256" height="256">',
        output,
    )
    return output
