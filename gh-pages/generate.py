"""Prepare the website sources before ``zensical build``.

Run from ``gh-pages/``::

    python generate.py
    zensical build --clean --strict

Zensical has no build hooks, so everything the site needs beyond the stock
build is produced here, ahead of time:

* ``_generated/*.md`` - website copies of the canonical repo docs (README files,
  FEATURES.md, CHANGELOG.md, docs/*). The thin wrapper pages under ``docs/``
  pull them in with the ``pymdownx.snippets`` ``--8<--`` syntax, so the site can
  never drift from the real docs. Links are rewritten to the site or GitHub.
* ``overrides/generated/lastmod/<page>index.html`` - the git commit date behind
  each page, included by ``overrides/sitemap.xml`` as ``<lastmod>``.
* ``overrides/generated/faq/<page>index.html`` - FAQPage JSON-LD derived from a
  page's own question headings, included by ``overrides/main.html``.

Both output folders are git-ignored and rebuilt from scratch on every run.
"""

from __future__ import annotations

import json
import posixpath
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import usage_analytics_page

GH_PAGES = Path(__file__).resolve().parent
REPO_ROOT = GH_PAGES.parent
DOCS_DIR = GH_PAGES / "docs"
GEN_DIR = GH_PAGES / "_generated"
PARTIALS_DIR = GH_PAGES / "overrides" / "generated"

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
    "docs/AGENT-SKILLS.md": "/skills/",
}

_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")
_SNIPPET = re.compile(r'^[ \t]*(?:-{2,}8<-{2,})[ \t]+"([^"]+)"[ \t]*$', re.MULTILINE)
_FRONTMATTER = re.compile(r"\A---\r?\n.*?\r?\n---\r?\n", re.DOTALL)
# Matches the snippet includes in the wrapper pages, e.g.
#     --8<-- "_generated/features-data.md"
_GEN_INCLUDE = re.compile(r'--8<--\s*"_generated/([^"]+)"')
# Same order as the snippets base_path in zensical.toml.
SNIPPET_BASE_PATHS = (DOCS_DIR, GH_PAGES)

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


# Canonical docs/reference/*.md pages. Each becomes _generated/reference-<name>
# and is published at /reference/<name without .md>/.
REFERENCE_SOURCES = (
    "workflows.md",
    "behavioral-rules.md",
    "workbook.md",
    "worksheet.md",
    "range.md",
    "table.md",
    "powerquery.md",
    "m-code-syntax.md",
    "datamodel.md",
    "dmv-reference.md",
    "pivottable.md",
    "querytable.md",
    "analysis.md",
    "chart.md",
    "conditionalformat.md",
    "slicer.md",
    "drawing.md",
    "screenshot.md",
    "report-formatting.md",
    "window.md",
    "xmlmap.md",
    "calculation.md",
)

SITE_PAGE_MAP.update(
    {f"docs/reference/{name}": f"/reference/{name.removesuffix('.md')}/" for name in REFERENCE_SOURCES}
)
SITE_PAGE_MAP["docs/reference/README.md"] = "/reference/"
SITE_PAGE_MAP["docs/guides/CLAUDE-DESKTOP.md"] = "/guides/claude-desktop/"


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


def _git_lastmod_index() -> dict[str, str]:
    """Map every tracked repo-relative path to its last commit date (W3C).

    One ``git log`` walk over the whole history, newest first: the first time a
    path appears is by definition its most recent change. A build date would be
    a false freshness signal on every page in every deploy.

    Returns an empty index (so ``<lastmod>`` is simply omitted) when git is
    unavailable, which keeps the build working from a source tarball.
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
        print(f"WARNING: git log failed ({exc}); sitemap will omit <lastmod>", file=sys.stderr)
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

def _write(name: str, source_rel: str, content: str) -> None:
    content = _rewrite_links(content, source_rel)
    (GEN_DIR / name).write_text(content, encoding="utf-8", newline="\n")
    MIRROR_SOURCES[name] = source_rel


# Generated file name -> repo-relative canonical source, filled by _write. A
# wrapper page's real "last modified" date comes from the file it mirrors.
MIRROR_SOURCES: dict[str, str] = {}


_FAQ_ADMONITION = re.compile(r'^\?{3}\+?\s+question\s+"([^"]+)"\s*$')
_FAQ_HEADING = re.compile(r"^###\s+(.+?)\s*$")
_FAQ_MIN_ENTITIES = 3


def generate_pages() -> None:
    _write(
        "usage-analytics.md",
        ".github/usage-analytics.json",
        usage_analytics_page.render(json.loads(_read(".github/usage-analytics.json"))),
    )

    _write(
        "features.md",
        "FEATURES.md",
        _strip_header(_read("FEATURES.md"), end_on_blank=True),
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

    # docs/AGENT-SKILLS.md -> skills (drop title, demote H1)
    _write(
        "skills.md",
        "docs/AGENT-SKILLS.md",
        _strip_header(
            _read("docs/AGENT-SKILLS.md"),
            end_on_blank=True,
            demote_h1=True,
        ),
    )

    # Canonical documentation -> reference pages (the wrapper owns the title).
    for name in REFERENCE_SOURCES:
        _write(
            f"reference-{name}",
            f"docs/reference/{name}",
            _strip_header(
                _read(f"docs/reference/{name}"), end_on_blank=True, demote_h1=True
            ),
        )
    _write("reference-index.md", "docs/reference/README.md",
           _strip_header(_read("docs/reference/README.md"), end_on_blank=True, demote_h1=True))
    _write("guides-claude-desktop.md", "docs/guides/CLAUDE-DESKTOP.md",
           _strip_header(_read("docs/guides/CLAUDE-DESKTOP.md"), end_on_blank=True, demote_h1=True))

    # Verbatim copies (these keep their own H1 as the page title).
    _write("contributing.md", "docs/CONTRIBUTING.md", _read("docs/CONTRIBUTING.md").strip() + "\n")
    _write("security.md", "SECURITY.md", _read("SECURITY.md").strip() + "\n")
    _write("privacy.md", "PRIVACY.md", _read("PRIVACY.md").strip() + "\n")


def _resolve_snippets(text: str, depth: int = 0) -> str:
    """Expand ``--8<-- "path"`` includes the same way the site build does."""
    if depth > 5:
        return text

    def repl(match: re.Match) -> str:
        for base in SNIPPET_BASE_PATHS:
            target = base / match.group(1)
            if target.is_file():
                return _resolve_snippets(target.read_text(encoding="utf-8"), depth + 1)
        raise FileNotFoundError(f"snippet not found: {match.group(1)}")

    return _SNIPPET.sub(repl, text)


def _page_url(page: Path) -> str:
    """Site URL path of a docs page: ``faq.md`` -> ``faq/``, ``index.md`` -> ``""``."""
    rel = page.relative_to(DOCS_DIR).as_posix().removesuffix(".md")
    if rel == "index" or rel.endswith("/index"):
        rel = rel.removesuffix("index")
        return rel
    return rel + "/"


def _write_partial(kind: str, url: str, content: str) -> None:
    # Always .html: the theme copies any other file under overrides/ into the
    # published site as a static asset instead of treating it as a template.
    path = PARTIALS_DIR / kind / url / "index.html"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def generate_page_partials() -> None:
    """Write the per-page sitemap ``<lastmod>`` and FAQ JSON-LD partials.

    A page's ``<lastmod>`` is the newest git date of the wrapper page and of the
    canonical source it mirrors, so editing either refreshes the entry. When git
    is unavailable the partials are simply not written and the sitemap omits
    ``<lastmod>`` rather than claiming the build date.
    """
    index = _git_lastmod_index()
    if index and _git_is_shallow():
        # A shallow clone still lists every file, all under the tip commit's
        # date. audit_site.py catches it by noticing every page has one date.
        print(
            "WARNING: shallow git clone - every sitemap <lastmod> will be the "
            "tip commit's date; the workflow needs fetch-depth: 0",
            file=sys.stderr,
        )

    for page in sorted(DOCS_DIR.rglob("*.md")):
        url = _page_url(page)
        text = page.read_text(encoding="utf-8")

        candidates = []
        wrapper_rel = f"gh-pages/docs/{page.relative_to(DOCS_DIR).as_posix()}"
        if wrapper_rel in index:
            candidates.append(index[wrapper_rel])
        for name in _GEN_INCLUDE.findall(text):
            source_rel = MIRROR_SOURCES.get(name)
            if source_rel in index:
                candidates.append(index[source_rel])
        if candidates:
            # git's %cI keeps each committer's UTC offset, so parse before max.
            date = max(candidates, key=datetime.fromisoformat)
            _write_partial("lastmod", url, f"<lastmod>{date}</lastmod>")

        faq = _faq_jsonld(_resolve_snippets(_FRONTMATTER.sub("", text)))
        if faq:
            # "</" cannot appear inside a <script> element.
            faq = faq.replace("</", "<\\/")
            _write_partial(
                "faq", url, f'<script type="application/ld+json">{faq}</script>'
            )


def main() -> int:
    for folder in (GEN_DIR, PARTIALS_DIR):
        shutil.rmtree(folder, ignore_errors=True)
        folder.mkdir(parents=True)
    generate_pages()
    generate_page_partials()
    print(f"generate.py: wrote {len(MIRROR_SOURCES)} pages to {GEN_DIR.name}/ and page partials")
    return 0


if __name__ == "__main__":
    sys.exit(main())
