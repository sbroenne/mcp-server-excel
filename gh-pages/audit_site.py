#!/usr/bin/env python3
"""Audit the built MkDocs site for SEO and LLM-discoverability regressions.

Run after ``mkdocs build`` from the ``gh-pages`` directory::

    python -m mkdocs build --strict --clean
    python audit_site.py

Exits non-zero and prints every failure if the built site regresses. This runs in
the Pages deploy workflow rather than the pre-commit hook, so the docs-only
pre-commit fast path stays fast.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import urlsplit

SITE_DIR = Path(__file__).resolve().parent / "_site"
SITE_URL = "https://excelmcpserver.dev/"

# Google truncates around these lengths; well outside them is a real problem.
TITLE_MAX = 70
DESCRIPTION_MIN = 50
DESCRIPTION_MAX = 200

failures: list[str] = []
checked = 0


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
        length = len(description.group(1).strip())
        if not DESCRIPTION_MIN <= length <= DESCRIPTION_MAX:
            fail(
                f"{name}: meta description is {length} chars "
                f"(want {DESCRIPTION_MIN}-{DESCRIPTION_MAX})"
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
    # SVG badges carry intrinsic dimensions in the file itself, so they are exempt.
    svg_hosts = ("img.shields.io", "cdn.jsdelivr.net")
    for img in re.findall(r"<img\b[^>]*>", html):
        src_match = re.search(r'src=["\']?([^"\'\s>]+)', img)
        src = src_match.group(1) if src_match else ""
        is_svg = src.split("?")[0].endswith(".svg") or any(h in src for h in svg_hosts)
        if "data:image" in img or is_svg:
            continue
        if "width=" not in img or "height=" not in img:
            fail(f"{name}: <img> without width/height: {src or img[:60]}")

    # Markdown alternate must be advertised and must exist.
    if 'type="text/markdown"' not in html and "type=text/markdown" not in html:
        fail(f"{name}: no <link rel=alternate type=text/markdown>")


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
    path = SITE_DIR / "troubleshooting" / "index.html"
    if not path.is_file():
        fail("troubleshooting page is missing")
        return
    if "FAQPage" not in path.read_text(encoding="utf-8"):
        fail("troubleshooting page has no FAQPage structured data")


def main() -> int:
    if not SITE_DIR.is_dir():
        print(f"ERROR: {SITE_DIR} not found - run 'mkdocs build' first", file=sys.stderr)
        return 2

    html_files = sorted(
        p
        for p in SITE_DIR.rglob("*.html")
        if p.name != "404.html" and "assets" not in p.relative_to(SITE_DIR).parts
    )
    if not html_files:
        print("ERROR: no HTML pages found in the built site", file=sys.stderr)
        return 2

    for path in html_files:
        audit_html(path)
    audit_internal_links(html_files)
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
