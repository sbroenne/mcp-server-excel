"""Machine-readable outputs for LLMs and answer engines.

* ``/llms.txt`` - navigation-ordered page index (llmstxt.org convention);
* ``/llms-full.txt`` - every page's Markdown with snippet includes resolved;
* one ``index.md`` Markdown mirror next to every page's ``index.html``;
* ``/tools.json`` - capability summaries derived from ``docs/features``;
* FAQPage JSON-LD built from a page's own question blocks.

All of them are derived from the Markdown MkDocs is building, so they cannot
drift from the rendered site.
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path

from sitegen.sources import DOC_COUNTS, DOCS_DIR, GH_PAGES, PAGES, read

log = logging.getLogger("mkdocs.hooks.generate")

# Raw Markdown of every built page, captured in on_page_markdown with --8<--
# includes resolved, keyed by the page's src_uri.
_PAGE_MARKDOWN: dict[str, dict[str, str]] = {}
# The resolved Navigation items, captured in on_nav. config["nav"] holds the raw
# YAML nav, which has no page objects to correlate with captured Markdown.
_NAV: list = []

_SNIPPET = re.compile(r'^[ \t]*(?:-{2,}8<-{2,})[ \t]+"([^"]+)"[ \t]*$', re.MULTILINE)
_FRONTMATTER = re.compile(r"\A---\r?\n.*?\r?\n---\r?\n", re.DOTALL)
_MD_LINK = re.compile(r"(?<!!)\[([^\]]+)\]\(([^)\s]+)\)")
# Mirrors the snippets `base_path` in mkdocs.yml, in the same order, so the
# llms.txt/mirror output resolves exactly what the site renders.
SNIPPET_BASE_PATHS = (DOCS_DIR, GH_PAGES)

_FAQ_ADMONITION = re.compile(r'^\?{3}\+?\s+question\s+"([^"]+)"\s*$')
_FAQ_HEADING = re.compile(r"^###\s+(.+?)\s*$")
_FAQ_MIN_ENTITIES = 3


def headline_counts() -> tuple[int, int]:
    """Read the canonical tool/operation totals from ``doc-counts.json``.

    Written by ``scripts/check-doc-counts.ps1 -Update``. Every count consumer --
    this site, release notes, packaging metadata -- reads that one file instead
    of deriving or parsing totals from Markdown headline text.
    """
    counts = json.loads(read(DOC_COUNTS))
    return int(counts["tools"]), int(counts["operations"])


def resolve_snippets(text: str, depth: int = 0) -> str:
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
                return resolve_snippets(target.read_text(encoding="utf-8"), depth + 1)
        log.warning("snippet not found while building llms output: %s", match.group(1))
        return ""

    return _SNIPPET.sub(repl, text)


def set_nav(items) -> None:
    _NAV.clear()
    _NAV.extend(items)


def capture_page(markdown: str, page, site_url: str) -> None:
    """Record a page's full Markdown and attach any FAQ structured data."""
    body = resolve_snippets(_FRONTMATTER.sub("", markdown)).strip()
    _PAGE_MARKDOWN[page.file.src_uri] = {
        "title": page.title or page.file.src_uri,
        "url": site_url + page.url,
        "description": (page.meta or {}).get("description", "").strip(),
        "markdown": body,
        "dest": page.file.dest_uri,
    }

    faq = faq_jsonld(body)
    if faq:
        page.meta["faq_jsonld"] = faq


def faq_jsonld(markdown: str) -> str:
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


def write_llm_outputs(site_dir: Path, site_url: str) -> None:
    """Emit /llms.txt, /llms-full.txt and one Markdown mirror per page."""
    headline_tools, headline_operations = headline_counts()

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
        f"application, exposing {headline_tools} tools and "
        f"{headline_operations} operations to AI assistants "
        "over the Model Context Protocol and to scripts through "
        "the `excelcli` command line. Unlike file-parser libraries it can refresh "
        "Power Query, evaluate DAX against the Data Model, refresh PivotTables, "
        "and run VBA on Windows, because Excel itself does the work. "
        "Windows requires Microsoft Excel 2016 or later. Apple Silicon macOS "
        "support is experimental beta with a capability-gated subset; consult "
        "the Mac support page for unavailable features and recovery rules.",
        "",
        "Every page below is also available as Markdown by appending `index.md` "
        "to its URL. The complete corpus is at "
        f"{site_url}llms-full.txt.",
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
        full.extend(
            [f"# {entry['title']}", "", f"Source: {entry['url']}", "", entry["markdown"], "", "---", ""]
        )
    (site_dir / "llms-full.txt").write_text(
        "\n".join(full), encoding="utf-8", newline="\n"
    )

    log.info("wrote llms.txt, llms-full.txt and %d Markdown mirrors", mirrored)


def write_tools_json(site_dir: Path, site_url: str, repo_url: str) -> None:
    """Emit /tools.json: capability summaries with tool and operation totals.

    Summaries describe grouped capabilities, not individual commands. The
    summaries and operation total come from the canonical ``docs/features/*.md``
    pages; the tool total comes from ``doc-counts.json``.
    """
    heading = re.compile(r"^## (?:\W+\s+)?(?P<name>.+?) \((?P<count>\d+) operations\)$")
    capability = re.compile(r"^- \*\*(?P<name>[^:*]+):\*\*\s*(?P<desc>.+)$")

    headline_tools, _ = headline_counts()

    categories = []
    total_ops = 0

    for page in PAGES:
        if not page.feature_title:
            continue
        groups: list[dict] = []
        current: dict | None = None
        for line in read(page.source).splitlines():
            match = heading.match(line)
            if match:
                current = {
                    "name": match.group("name").strip(),
                    "operationCount": int(match.group("count")),
                    "capabilities": [],
                }
                groups.append(current)
                continue
            if current is None:
                continue
            summary = capability.match(line)
            if summary:
                current["capabilities"].append(
                    {
                        "name": summary.group("name").strip(),
                        "description": summary.group("desc").strip(),
                    }
                )

        total_ops += sum(g["operationCount"] for g in groups)
        categories.append(
            {
                "name": page.feature_title,
                "url": site_url.rstrip("/") + page.url,
                "operationCount": sum(g["operationCount"] for g in groups),
                "featureGroups": groups,
            }
        )

    payload = {
        "name": "Excel MCP Server",
        "url": site_url,
        "repository": repo_url,
        "description": (
            "Automates the real Microsoft Excel application through Windows COM "
            "or capability-gated Apple Events on Apple Silicon macOS, "
            "exposing Excel to AI assistants over the Model Context Protocol and "
            "to scripts through the excelcli command line. "
            "Feature groups describe capabilities, not individual commands."
        ),
        "requirements": {
            "operatingSystem": ["Windows", "Apple Silicon macOS (experimental beta)"],
            "application": "Microsoft Excel desktop; Windows requires 2016 or later",
        },
        "catalogueScope": "Full Windows capability catalogue; Mac availability is recorded separately.",
        "platformSupport": {
            "Windows": {"status": "supported", "backend": "COM"},
            "macOS": {
                "status": "experimental beta",
                "architecture": "Apple Silicon",
                "backend": "Apple Events",
                "supportPage": site_url.rstrip("/") + "/macos-support/",
                "capabilityInventory": repo_url.rstrip("/") + "/blob/main/docs/MACOS-ACTION-INVENTORY.md",
                "unsupportedFeatures": [
                    "Power Query",
                    "Public VBA module/source actions and arbitrary macro execution",
                    "Data Model, DAX, and OLAP",
                    "Tables, PivotTables, charts, and slicers",
                    "Connections and QueryTables",
                    "Screenshots and XML Maps",
                    "Advanced worksheet, visual, calculation, and file variants",
                ],
            },
        },
        "entryPoints": ["mcp-server", "cli"],
        "toolCount": headline_tools,
        "operationCount": total_ops,
        "categories": categories,
    }

    (site_dir / "tools.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    log.info("wrote tools.json (%d tools, %d operations)", headline_tools, total_ops)
