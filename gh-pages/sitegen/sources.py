"""Which repository documents the website publishes, and how each one is adapted.

``PAGES`` is the single list of published canonical documents. Everything else
derives from it: the generated snippet files, repo-link rewriting
(``SITE_PAGE_MAP``), sitemap dating, ``tools.json`` categories, and the deploy
workflow path check in ``check_deploy_paths.py``.
"""

from __future__ import annotations

import logging
import posixpath
import re
from dataclasses import dataclass
from pathlib import Path

from mkdocs.structure.files import File

log = logging.getLogger("mkdocs.hooks.generate")

GH_PAGES = Path(__file__).resolve().parent.parent
REPO_ROOT = GH_PAGES.parent
DOCS_DIR = GH_PAGES / "docs"
# Deliberately OUTSIDE docs_dir. When the generated files lived in
# docs/_generated/, every build rewrote files inside the directory `mkdocs
# serve` watches, so a single edit put the dev server into an endless
# rebuild loop. `.` is a snippets base_path, so the `--8<-- "_generated/..."`
# includes in the wrapper pages resolve here unchanged.
GEN_DIR = GH_PAGES / "_generated"

# The single generated count file every count consumer reads; see
# scripts/check-doc-counts.ps1.
DOC_COUNTS = "doc-counts.json"


@dataclass(frozen=True)
class Page:
    """One canonical repository document published as a website page.

    ``header`` controls how the source's own title block is removed (the H1
    lives in the hand-written wrapper under ``docs/``):

    * ``"blank"`` - drop the H1 block up to the first blank line;
    * ``"hr"`` - drop it up to the first ``---`` rule;
    * ``"keep"`` - publish verbatim, keeping the source's H1;
    * ``"rendered"`` - the content is rendered from data, not copied.
    """

    output: str
    source: str
    url: str
    header: str = "blank"
    demote_h1: bool = True
    drop_prefixes: tuple[str, ...] = ()
    # Set for docs/features pages: tools.json category name; also adds stable
    # heading anchors that survive operation-count changes.
    feature_title: str = ""


def _feature(output: str, source: str, url: str, title: str) -> Page:
    return Page(output, source, url, header="hr", demote_h1=False, feature_title=title)


# Canonical documentation references (docs/reference/<name>.md ->
# /reference/<slug>/). Only report formatting is also packaged as a skill
# reference.
_REFERENCE_PAGES = {
    "workflows.md": "skills-workflows.md",
    "behavioral-rules.md": "skills-behavioral-rules.md",
    "workbook.md": "skills-workbook.md",
    "worksheet.md": "skills-worksheet.md",
    "range.md": "skills-range.md",
    "table.md": "skills-table.md",
    "powerquery.md": "skills-powerquery.md",
    "m-code-syntax.md": "skills-m-code-syntax.md",
    "datamodel.md": "skills-datamodel.md",
    "dmv-reference.md": "skills-dmv-reference.md",
    "pivottable.md": "skills-pivottable.md",
    "querytable.md": "skills-querytable.md",
    "analysis.md": "skills-analysis.md",
    "chart.md": "skills-chart.md",
    "conditionalformat.md": "skills-conditionalformat.md",
    "slicer.md": "skills-slicer.md",
    "drawing.md": "skills-drawing.md",
    "screenshot.md": "skills-screenshot.md",
    "report-formatting.md": "skills-report-formatting.md",
    "window.md": "skills-window.md",
    "xmlmap.md": "skills-xmlmap.md",
    "calculation.md": "reference-calculation.md",
}

PAGES: tuple[Page, ...] = (
    Page("world-in-motion.md", "samples/world-bank-dashboard/README.md", "/samples/world-in-motion/"),
    Page("usage-analytics.md", ".github/usage-analytics.json", "/usage-analytics/", header="rendered"),
    Page("features.md", "FEATURES.md", "/features/", demote_h1=False),
    _feature("features-data.md", "docs/features/DATA-ANALYTICS.md",
             "/features/data-analytics/", "Data & Analytics"),
    _feature("features-workbooks.md", "docs/features/CELLS-WORKBOOKS.md",
             "/features/cells-workbooks/", "Cells & Workbooks"),
    _feature("features-visualization.md", "docs/features/CHARTS-VISUALS.md",
             "/features/charts-visuals/", "Charts & Visualization"),
    _feature("features-automation.md", "docs/features/AUTOMATION-ADVANCED.md",
             "/features/automation-advanced/", "Automation & Advanced"),
    Page("guides-index.md", "docs/guides/README.md", "/guides/"),
    Page("guides-refresh-power-query.md", "docs/guides/REFRESH-POWER-QUERY.md",
         "/guides/refresh-power-query/"),
    Page("guides-automate-pivottables.md", "docs/guides/AUTOMATE-PIVOTTABLES.md",
         "/guides/automate-pivottables/"),
    Page("guides-run-vba-macros.md", "docs/guides/RUN-VBA-MACROS.md", "/guides/run-vba-macros/"),
    Page("guides-query-data-model-with-dax.md", "docs/guides/QUERY-DATA-MODEL-WITH-DAX.md",
         "/guides/query-data-model-with-dax/"),
    Page("guides-excel-com-vs-file-parsers.md", "docs/guides/EXCEL-COM-VS-FILE-PARSERS.md",
         "/guides/excel-automation-vs-file-parsers/"),
    Page("guides-claude-desktop.md", "docs/guides/CLAUDE-DESKTOP.md", "/guides/claude-desktop/"),
    Page("changelog.md", "CHANGELOG.md", "/changelog/", drop_prefixes=("This changelog",)),
    Page("installation.md", "docs/INSTALLATION.md", "/installation/",
         drop_prefixes=("Complete installation",)),
    Page("installation-mcp-server.md", "docs/INSTALLATION-MCP-SERVER.md", "/installation-mcp-server/"),
    Page("installation-cli.md", "docs/INSTALLATION-CLI.md", "/installation-cli/"),
    Page("architecture.md", "docs/ARCHITECTURE.md", "/architecture/", demote_h1=False),
    Page("use-cases.md", "docs/USE-CASES.md", "/use-cases/", demote_h1=False),
    Page("mcp-server.md", "src/ExcelMcp.McpServer/README.md", "/mcp-server/",
         drop_prefixes=("<!-- mcp-name", "mcp-name:", "[![")),
    Page("cli.md", "src/ExcelMcp.CLI/README.md", "/cli/", drop_prefixes=("[![",)),
    Page("skills.md", "docs/AGENT-SKILLS.md", "/skills/"),
    *(
        Page(output, f"docs/reference/{name}", f"/reference/{name.removesuffix('.md')}/")
        for name, output in _REFERENCE_PAGES.items()
    ),
    Page("reference-index.md", "docs/reference/README.md", "/reference/"),
    Page("contributing.md", "docs/CONTRIBUTING.md", "/contributing/", header="keep"),
    Page("security.md", "SECURITY.md", "/security/", header="keep"),
    Page("privacy.md", "PRIVACY.md", "/privacy/", header="keep"),
)

SAMPLE_ASSETS = {
    "downloads/world-in-motion.xlsx": "samples/world-bank-dashboard/world-in-motion.xlsx",
    "downloads/world-bank-sources.json": "samples/world-bank-dashboard/data/sources.json",
    "downloads/world-bank-indicators.csv": "samples/world-bank-dashboard/data/indicators.csv",
    "assets/images/world-in-motion/overview.png": "videos/world-in-motion-demo/capture/assets/overview.png",
    "assets/images/world-in-motion/growth.png": "videos/world-in-motion-demo/capture/assets/growth.png",
    "assets/images/world-in-motion/progress.png": "videos/world-in-motion-demo/capture/assets/progress.png",
}
SITE_ASSET_MAP = {source: "/" + destination for destination, source in SAMPLE_ASSETS.items()}

# Repo-relative source -> site path, for rewriting links into published pages.
SITE_PAGE_MAP = {page.source: page.url for page in PAGES}
# Generated snippet name -> canonical source, for dating sitemap entries.
MIRROR_SOURCES = {page.output: page.source for page in PAGES}
# docs/features generated snippet name -> canonical source.
FEATURE_SOURCES = {page.output: page.source for page in PAGES if page.feature_title}
# Every repository file the build reads; check_deploy_paths.py keeps the deploy
# workflow's paths filter in step with it.
SOURCE_FILES = frozenset({*SITE_PAGE_MAP, *SAMPLE_ASSETS.values(), DOC_COUNTS})


def add_sample_assets(files, config):
    for destination, source in SAMPLE_ASSETS.items():
        path = REPO_ROOT / source
        if not path.is_file():
            raise FileNotFoundError(f"Sample asset not found: {path}")
        if files.get_file_from_path(destination) is not None:
            raise ValueError(f"Duplicate sample asset destination: {destination}")
        files.append(File.generated(config, destination, abs_src_path=str(path)))
    return files


_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")


def read(rel: str) -> str:
    path = REPO_ROOT / rel
    if not path.is_file():
        raise FileNotFoundError(f"Source doc not found: {path}")
    return path.read_text(encoding="utf-8")


def rewrite_links(text: str, source_rel: str, repo_url: str) -> str:
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
    repo_url = repo_url.rstrip("/")
    github_blob = f"{repo_url}/blob/main/"
    github_tree = f"{repo_url}/tree/main/"
    source_dir = posixpath.dirname(source_rel)

    def repl(match: re.Match) -> str:
        label, url = match.group(1), match.group(2)

        for prefix in (github_blob, github_tree):
            if url.startswith(prefix):
                remainder = url[len(prefix) :]
                target, _, anchor = remainder.partition("#")
                anchor = f"#{anchor}" if anchor else ""
                published = SITE_PAGE_MAP.get(target.rstrip("/")) or SITE_ASSET_MAP.get(target.rstrip("/"))
                if published:
                    return f"[{label}]({published}{anchor})"
                return match.group(0)

        if url.startswith(("http://", "https://", "#", "/", "mailto:", "<")):
            return match.group(0)

        target, _, anchor = url.partition("#")
        anchor = f"#{anchor}" if anchor else ""
        if target == "":
            return match.group(0)  # pure in-page anchor

        resolved = posixpath.normpath(posixpath.join(source_dir, target))
        if resolved.startswith(".."):
            return match.group(0)  # points outside the repo; leave as-is

        published = SITE_PAGE_MAP.get(resolved) or SITE_ASSET_MAP.get(resolved)
        if published:
            return f"[{label}]({published}{anchor})"

        base = github_tree if url.endswith("/") else github_blob
        return f"[{label}]({base}{resolved}{anchor})"

    return _MD_LINK.sub(repl, text)


def strip_header(
    text: str,
    *,
    drop_prefixes: tuple[str, ...] = (),
    end_on_blank: bool = False,
    end_on_hr: bool = False,
    demote_h1: bool = False,
) -> str:
    """Drop the leading H1 title block from a source file, optionally demoting
    any remaining H1 headings to H2.

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


def add_stable_feature_anchors(text: str) -> str:
    """Give feature headings stable IDs that do not include operation counts."""
    heading = re.compile(r"^## (?P<title>.+?) \(\d+ operations\)$", re.MULTILINE)

    def replace(match: re.Match) -> str:
        title = match.group("title")
        slug = re.sub(r"[^\w\s-]", "", title, flags=re.UNICODE).strip().lower()
        slug = re.sub(r"[-\s]+", "-", slug)
        return f"{match.group(0)} {{ #{slug} }}"

    return heading.sub(replace, text)


def adapt(page: Page, text: str) -> str:
    """Turn a canonical source into the website snippet for ``page``."""
    if page.header == "keep":
        return text.strip() + "\n"
    content = strip_header(
        text,
        drop_prefixes=page.drop_prefixes,
        end_on_blank=page.header == "blank",
        end_on_hr=page.header == "hr",
        demote_h1=page.demote_h1,
    )
    return add_stable_feature_anchors(content) if page.feature_title else content


def write(page: Page, content: str, repo_url: str) -> None:
    GEN_DIR.mkdir(parents=True, exist_ok=True)
    content = rewrite_links(content, page.source, repo_url)
    (GEN_DIR / page.output).write_text(content, encoding="utf-8")
    log.info("generated _generated/%s", page.output)
