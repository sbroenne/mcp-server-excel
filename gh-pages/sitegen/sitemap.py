"""Sitemap data: real git ``<lastmod>`` dates and per-page video entries.

``overrides/sitemap.xml`` renders both. MkDocs' stock sitemap stamps every URL
with the *build* date, a false freshness signal on every page in every deploy.
"""

from __future__ import annotations

import logging
import re
import subprocess
from datetime import datetime
from pathlib import Path

from sitegen.sources import MIRROR_SOURCES, REPO_ROOT

log = logging.getLogger("mkdocs.hooks.generate")

# Preserve the homepage introduction and its matching VideoObject in docs/index.md.
VIDEOS = [
    {
        "page_path": "",
        "thumbnail": "https://i.ytimg.com/vi/wbw3-hPcE2o/maxresdefault.jpg",
        "title": "Excel MCP Server: Real Excel Automation for AI Agents",
        "description": (
            "Learn what Excel MCP Server is, when to use it, and how AI agents automate "
            "Power Query, DAX, PivotTables, VBA, Python, and calculations through real "
            "Microsoft Excel."
        ),
        "player_loc": "https://www.youtube.com/embed/wbw3-hPcE2o",
        "duration": "121",
        "publication_date": "2026-09-12T07:07:06-07:00",
    },
    {
        "page_path": "samples/world-in-motion/",
        "thumbnail": "https://i.ytimg.com/vi/47HJPZbcta4/maxresdefault.jpg",
        "title": "AI-Built Excel Dashboards | ExcelMCP in Action",
        "description": (
            "See a real Excel workbook built by GPT-6 Astra through ExcelMCP, "
            "using World Bank data, Power Query, a Data Model, DAX, PivotTables "
            "and interactive dashboards. Excel powers the workbook. Download "
            "the sample and ask your agent to adapt it."
        ),
        "player_loc": "https://www.youtube.com/embed/47HJPZbcta4",
        "duration": "154",
    },
]

# Matches the snippet includes in the wrapper pages, e.g.
#     --8<-- "_generated/features-data.md"
_GEN_INCLUDE = re.compile(r'--8<--\s*"_generated/([^"]+)"')


def _git_lastmod_index() -> dict[str, str]:
    """Map every tracked repo-relative path to its last commit date (W3C).

    One ``git log`` walk over the whole history, newest first: the first time a
    path appears is by definition its most recent change.

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


def page_lastmod(files) -> dict[str, str]:
    """Map each page's ``src_uri`` to the newest git date that affects it.

    For a wrapper page that is nothing but an ``--8<--`` include, that is the
    date of the canonical source; the wrapper itself contributes its own date
    too, so editing either one refreshes the entry.
    """
    index = _git_lastmod_index()
    if not index:
        return {}
    if _git_is_shallow():
        # actions/checkout defaults to fetch-depth: 1. audit_site.py catches the
        # result by noticing that every page claims the same <lastmod>.
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
            if source_rel in index:
                candidates.append(index[source_rel])
        if candidates:
            # git's %cI keeps each committer's UTC offset, so the strings are
            # not directly comparable as instants - parse before taking the max.
            lastmod[file.src_uri] = max(candidates, key=datetime.fromisoformat)
    return lastmod
