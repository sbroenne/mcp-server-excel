#!/usr/bin/env python3
"""Audit the built MkDocs site for SEO and LLM-discoverability regressions.

Run after ``mkdocs build`` from the ``gh-pages`` directory::

    python -m mkdocs build --strict --clean
    python audit_site.py

Exits non-zero and prints every failure if the built site regresses. This runs in
the ``Docs Site`` CI job on every pull request and in the Pages deploy workflow,
rather than in the pre-commit hook, so the docs-only pre-commit fast path stays
fast.
"""

from __future__ import annotations

import gzip
import json
import re
import sys
import zlib
from pathlib import Path
from urllib.parse import urlsplit

SITE_DIR = Path(__file__).resolve().parent / "_site"
MKDOCS_YML = Path(__file__).resolve().parent / "mkdocs.yml"
SITE_URL = "https://excelmcpserver.dev/"

# Imported rather than duplicated: this is the same mapping hooks.py uses to
# rewrite links, so the audit cannot drift away from what the build produces.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from hooks import SITE_PAGE_MAP as SOURCE_TO_SITE  # noqa: E402

# Google truncates around these lengths; well outside them is a real problem.
TITLE_MAX = 70
DESCRIPTION_MIN = 50
DESCRIPTION_MAX = 200

# The home page legitimately renders site_description; every other page must not.
HOMEPAGE = "index.html"

# Remote status badges: their pixel dimensions are not knowable at build time, and
# the URL may carry no file extension at all (e.g. img.shields.io/...?style=flat),
# so they are matched on host rather than on the path.
BADGE_HOSTS = ("img.shields.io", "cdn.jsdelivr.net", "vsmarketplacebadges.dev", "badgen.net")

# Material emits the theme logo from its own header partial and sizes it via CSS;
# it cannot carry width/height without overriding that partial.
THEME_LOGO_SUFFIXES = ("/logo.png", "/icon.png")

failures: list[str] = []
checked = 0


def _site_description() -> str:
    """Read ``site_description`` out of mkdocs.yml.

    Parsed rather than hardcoded on purpose: a copy of the string here would stop
    matching the moment someone rewords mkdocs.yml, and the fallback check below
    would then pass forever while guarding nothing - a silent failure inside a
    detector whose whole job is catching silent failures.

    A full ``yaml.safe_load`` is not an option because mkdocs.yml carries custom
    ``!!python/name:`` tags, so only this one key is parsed. Plain, quoted and
    ``>``/``|`` block scalar forms are all handled.
    """
    lines = MKDOCS_YML.read_text(encoding="utf-8").splitlines()
    for index, line in enumerate(lines):
        match = re.match(r"^site_description:\s*(.*?)\s*$", line)
        if not match:
            continue
        value = match.group(1)
        if value[:1] in (">", "|"):
            block: list[str] = []
            for follow in lines[index + 1 :]:
                if not follow.strip():
                    block.append("")
                    continue
                if not follow[:1].isspace():
                    break
                block.append(follow.strip())
            value = " ".join(x for x in block if x)
        else:
            value = value.strip("'\"")
        return " ".join(value.split())

    print("ERROR: could not find site_description in mkdocs.yml", file=sys.stderr)
    sys.exit(2)


SITE_DESCRIPTION = _site_description()


def fail(message: str) -> None:
    failures.append(message)


def page_name(path: Path) -> str:
    return path.relative_to(SITE_DIR).as_posix()


def audit_html(path: Path) -> None:
    global checked
    checked += 1
    html = path.read_text(encoding="utf-8", errors="replace")
    name = page_name(path)

    canonical = re.search(r'<link[^>]+rel=["\']?canonical["\']?[^>]*>', html)
    if not canonical:
        fail(f"{name}: no canonical link")
    elif SITE_URL not in canonical.group(0):
        fail(f"{name}: canonical does not point at {SITE_URL}")

    titles = re.findall(r"<title>(.*?)</title>", html, re.DOTALL)
    if not titles:
        fail(f"{name}: no <title>")
    elif len(titles[0].strip()) > TITLE_MAX:
        fail(f"{name}: <title> is {len(titles[0].strip())} chars (max {TITLE_MAX})")

    description = re.search(
        r'<meta[^>]+name=["\']?description["\']?[^>]+content=["\'](.*?)["\']',
        html,
        re.DOTALL,
    )
    if not description:
        fail(f"{name}: no meta description")
    else:
        text = " ".join(description.group(1).split())
        length = len(text)
        if not DESCRIPTION_MIN <= length <= DESCRIPTION_MAX:
            fail(
                f"{name}: meta description is {length} chars "
                f"(want {DESCRIPTION_MIN}-{DESCRIPTION_MAX})"
            )
        # Material falls back to site_description whenever a page has no usable
        # per-page description, and MkDocs reports nothing when that happens.
        # Two ways to trigger it, both silent: a double quote inside an unquoted
        # `description:` value terminates the rendered content="..." attribute
        # early, and an unquoted YAML scalar containing ": " makes the whole
        # front-matter block unparseable - dropping title, description and
        # keywords together.
        if name != HOMEPAGE and text == SITE_DESCRIPTION:
            fail(
                f"{name}: meta description fell back to site_description - this "
                f"page's YAML front matter did not apply (check the description "
                f'value for a double quote or an unquoted ": ")'
            )

    for prop in ("og:title", "og:description", "og:image", "og:url", "og:type"):
        if f'property="{prop}"' not in html and f"property={prop}" not in html:
            fail(f"{name}: missing {prop}")
    for prop in ("twitter:card", "twitter:title", "twitter:description", "twitter:image"):
        if f'name="{prop}"' not in html and f"name={prop}" not in html:
            fail(f"{name}: missing {prop}")

    # Exactly one <h1> per page.
    h1_count = len(re.findall(r"<h1[\s>]", html))
    if h1_count != 1:
        fail(f"{name}: found {h1_count} <h1> elements (want exactly 1)")

    # Raster content images need explicit dimensions to avoid layout shift.
    # SVGs carry intrinsic dimensions in the file itself, remote badges have no
    # build-time dimensions (and often no file extension either, so they are
    # matched on host), and the theme logo is emitted by Material's own partials.
    for img in re.findall(r"<img\b[^>]*>", html):
        src_match = re.search(r'src=["\']?([^"\'\s>]+)', img)
        src = src_match.group(1) if src_match else ""
        parts = urlsplit(src)
        exempt = (
            parts.netloc in BADGE_HOSTS
            or parts.path.endswith(".svg")
            or parts.path.endswith(THEME_LOGO_SUFFIXES)
        )
        if "data:image" in img or exempt:
            continue
        if "width=" not in img or "height=" not in img:
            fail(f"{name}: <img> without width/height: {src or img[:60]}")

    # Markdown alternate must be advertised and must exist.
    if 'type="text/markdown"' not in html and "type=text/markdown" not in html:
        fail(f"{name}: no <link rel=alternate type=text/markdown>")


def audit_offsite_links(html_files: list[Path]) -> None:
    """No page may link out to GitHub for content we publish ourselves.

    Canonical sources that are also rendered outside GitHub (the NuGet package
    READMEs) spell their links out as absolute GitHub URLs, because NuGet.org
    resolves relative links against the package root and they 404. hooks.py maps
    those back to the published page. If that mapping is missed - a new absolute
    link, or a page added without a SITE_PAGE_MAP entry - the site silently
    starts sending readers to GitHub instead of its own page, losing both the
    reader and the internal link equity.
    """
    # Quotes are optional: the minify plugin strips them from attribute values.
    pattern = re.compile(
        r'href=["\']?https://github\.com/sbroenne/mcp-server-excel/(?:blob|tree)/main/([^"\'\s>#]+)'
    )
    for path in html_files:
        html = path.read_text(encoding="utf-8", errors="replace")
        name = page_name(path)
        for target in pattern.findall(html):
            mapped = SOURCE_TO_SITE.get(target.rstrip("/"))
            if mapped is not None:
                fail(
                    f"{name}: links to GitHub for {target}, which is published at "
                    f"{mapped} - add the mapping in hooks.py instead"
                )


def audit_internal_links(html_files: list[Path]) -> None:
    """Every site-absolute internal link must resolve to something we built."""
    for path in html_files:
        html = path.read_text(encoding="utf-8", errors="replace")
        name = page_name(path)
        for href in re.findall(r'href=["\']?(/[^"\'\s>#]*)', html):
            target = urlsplit(href).path
            if target.startswith("//"):
                continue
            candidate = SITE_DIR / target.lstrip("/")
            if candidate.is_dir():
                candidate = candidate / "index.html"
            elif target.endswith("/"):
                candidate = SITE_DIR / target.strip("/") / "index.html"
            if not candidate.exists():
                fail(f"{name}: broken internal link {href}")


def audit_sitemap() -> None:
    sitemap = SITE_DIR / "sitemap.xml"
    if not sitemap.is_file():
        fail("sitemap.xml is missing")
        return
    xml = sitemap.read_text(encoding="utf-8")
    if "<lastmod>" in xml:
        fail("sitemap.xml still contains unreliable <lastmod> values")
    if "<video:video>" not in xml:
        fail("sitemap.xml is missing the home-page video markup")
    locs = re.findall(r"<loc>([^<]+)</loc>", xml)
    if not locs:
        fail("sitemap.xml contains no <loc> entries")
    for loc in locs:
        if not loc.startswith(SITE_URL):
            fail(f"sitemap.xml has an off-site <loc>: {loc}")
    if not (SITE_DIR / "sitemap.xml.gz").is_file():
        fail("sitemap.xml.gz is missing")
    else:
        # The gzipped twin is what many crawlers actually fetch, and hooks.py
        # rewrites it separately after enriching sitemap.xml. Checking only that
        # it exists would let the two silently diverge - publishing a .gz still
        # carrying the lastmod values the plain file had dropped.
        # gzip surfaces corruption three ways and only one of them is an OSError:
        # BadGzipFile (wrong format / CRC failure) subclasses it, but EOFError
        # (truncated write) and zlib.error (damaged deflate stream) inherit
        # straight from Exception. Catching OSError alone would let the two most
        # likely real-world cases escape as a traceback, burying every other
        # finding this audit produced. The type name is included because
        # "EOFError" says truncated write, where "BadGzipFile" says wrong format.
        try:
            with gzip.open(SITE_DIR / "sitemap.xml.gz", "rt", encoding="utf-8") as fh:
                if fh.read() != xml:
                    fail("sitemap.xml.gz does not match sitemap.xml")
        except (OSError, EOFError, zlib.error) as exc:
            kind = type(exc).__qualname__
            if type(exc).__module__ != "builtins":
                # zlib.error's bare name is just "error", which says nothing.
                kind = f"{type(exc).__module__}.{kind}"
            fail(f"sitemap.xml.gz could not be read: {kind}: {exc}")


def audit_llms(html_files: list[Path]) -> None:
    llms = SITE_DIR / "llms.txt"
    if not llms.is_file():
        fail("llms.txt is missing")
    else:
        text = llms.read_text(encoding="utf-8")
        lines = [x for x in text.splitlines() if x.strip()]
        if not lines or not lines[0].startswith("# "):
            fail("llms.txt must start with an H1")
        if not any(x.startswith("> ") for x in lines[:5]):
            fail("llms.txt must contain a blockquote summary near the top")
        if not any(x.startswith("## ") for x in lines):
            fail("llms.txt contains no sections")
        if text.count("](") < 10:
            fail("llms.txt lists suspiciously few pages")

    full = SITE_DIR / "llms-full.txt"
    if not full.is_file():
        fail("llms-full.txt is missing")
    else:
        text = full.read_text(encoding="utf-8")
        if len(text) < 50_000:
            fail(f"llms-full.txt is only {len(text)} chars; expected the full corpus")
        if "8<--" in text:
            fail("llms-full.txt contains unresolved --8<-- includes")

    # Every built page needs a Markdown mirror.
    for path in html_files:
        mirror = path.with_suffix(".md")
        if not mirror.is_file():
            fail(f"{page_name(path)}: no Markdown mirror at {page_name(mirror)}")
            continue
        content = mirror.read_text(encoding="utf-8")
        if content.startswith("---"):
            fail(f"{page_name(mirror)}: mirror still contains YAML front matter")
        if "8<--" in content:
            fail(f"{page_name(mirror)}: mirror contains unresolved --8<-- includes")
        if not content.strip():
            fail(f"{page_name(mirror)}: mirror is empty")


def audit_tools_json() -> None:
    path = SITE_DIR / "tools.json"
    if not path.is_file():
        fail("tools.json is missing")
        return
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        fail(f"tools.json is not valid JSON: {exc}")
        return

    features = (SITE_DIR.parent.parent / "FEATURES.md").read_text(encoding="utf-8")
    headline = re.search(
        r"\*\*(?P<tools>\d+) specialized tools with (?P<ops>\d+) operations", features
    )
    if headline is None:
        fail("could not read headline counts from FEATURES.md")
        return

    if data.get("toolCount") != int(headline.group("tools")):
        fail(
            f"tools.json toolCount {data.get('toolCount')} != "
            f"FEATURES.md headline {headline.group('tools')}"
        )
    if data.get("operationCount") != int(headline.group("ops")):
        fail(
            f"tools.json operationCount {data.get('operationCount')} != "
            f"FEATURES.md headline {headline.group('ops')}"
        )
    if not data.get("categories"):
        fail("tools.json has no categories")


def audit_robots() -> None:
    path = SITE_DIR / "robots.txt"
    if not path.is_file():
        fail("robots.txt is missing")
        return
    text = path.read_text(encoding="utf-8")
    for agent in ("GPTBot", "ClaudeBot", "PerplexityBot", "Google-Extended", "OAI-SearchBot"):
        if f"User-agent: {agent}" not in text:
            fail(f"robots.txt has no explicit policy for {agent}")
    if "Sitemap:" not in text:
        fail("robots.txt does not declare the sitemap")


def audit_faq() -> None:
    """The FAQ page must carry parseable FAQPage structured data.

    The JSON-LD is derived from the page's own ``###`` question headings by
    hooks.py, so a parser regression shows up as too few questions rather than
    as a build failure.
    """
    faq = SITE_DIR / "faq" / "index.html"
    if not faq.is_file():
        fail("FAQ page is missing")
        return
    html = faq.read_text(encoding="utf-8")
    blocks = re.findall(
        r'<script type=["\']?application/ld\+json["\']?>(.*?)</script>', html, re.DOTALL
    )
    faq_blocks = []
    for block in blocks:
        try:
            data = json.loads(block)
        except json.JSONDecodeError:
            continue  # audit_jsonld reports the parse error itself
        if data.get("@type") == "FAQPage":
            faq_blocks.append(data)
    if not faq_blocks:
        fail("faq/index.html: no FAQPage structured data")
    elif len(faq_blocks[0].get("mainEntity", [])) < 5:
        fail(
            f"faq/index.html: FAQPage has only "
            f"{len(faq_blocks[0].get('mainEntity', []))} questions - parser regression?"
        )

    if not (SITE_DIR / "troubleshooting" / "index.html").is_file():
        fail("troubleshooting page is missing")


def audit_jsonld(html_files: list[Path]) -> None:
    """Every JSON-LD block must actually parse, or search engines ignore it."""
    for path in html_files:
        html = path.read_text(encoding="utf-8", errors="replace")
        for block in re.findall(
            r'<script type=["\']?application/ld\+json["\']?>(.*?)</script>',
            html,
            re.DOTALL,
        ):
            try:
                json.loads(block)
            except json.JSONDecodeError as exc:
                fail(f"{page_name(path)}: invalid JSON-LD: {exc}")


def main() -> int:
    if not SITE_DIR.is_dir():
        print(f"ERROR: {SITE_DIR} not found - run 'mkdocs build' first", file=sys.stderr)
        return 2

    html_files = sorted(
        p
        for p in SITE_DIR.rglob("*.html")
        # 404.html is not an indexable page: MkDocs renders it with no canonical
        # URL and no page metadata, so every metadata check would fire on it.
        if p.name != "404.html" and "assets" not in p.relative_to(SITE_DIR).parts
    )
    if not html_files:
        print("ERROR: no HTML pages found in the built site", file=sys.stderr)
        return 2

    for path in html_files:
        audit_html(path)
    audit_internal_links(html_files)
    audit_offsite_links(html_files)
    audit_jsonld(html_files)
    audit_sitemap()
    audit_llms(html_files)
    audit_tools_json()
    audit_robots()
    audit_faq()

    if failures:
        print(f"Site audit FAILED - {len(failures)} issue(s) across {checked} pages:\n")
        for item in failures:
            print(f"  - {item}")
        return 1

    print(f"Site audit passed: {checked} pages, no issues found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
