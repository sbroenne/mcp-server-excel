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
GEN_DIR = Path(__file__).resolve().parent / "docs" / "_generated"

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
    "docs/CONTRIBUTING.md": "/contributing/",
    "SECURITY.md": "/security/",
    "PRIVACY.md": "/privacy/",
    "src/ExcelMcp.McpServer/README.md": "/mcp-server/",
    "src/ExcelMcp.CLI/README.md": "/cli/",
    "skills/README.md": "/skills/",
}

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")

FEATURE_SOURCES = {
    "features-data.md": "docs/features/DATA-ANALYTICS.md",
    "features-workbooks.md": "docs/features/CELLS-WORKBOOKS.md",
    "features-visualization.md": "docs/features/CHARTS-VISUALS.md",
    "features-automation.md": "docs/features/AUTOMATION-ADVANCED.md",
}


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

    # Verbatim copies (these keep their own H1 as the page title).
    _write("contributing.md", "docs/CONTRIBUTING.md", _read("docs/CONTRIBUTING.md").strip() + "\n")
    _write("security.md", "SECURITY.md", _read("SECURITY.md").strip() + "\n")
    _write("privacy.md", "PRIVACY.md", _read("PRIVACY.md").strip() + "\n")


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
