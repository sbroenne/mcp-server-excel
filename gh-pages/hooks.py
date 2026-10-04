"""MkDocs build hooks for excelmcpserver.dev.

The website publishes canonical Markdown from all over the repository (README
files, FEATURES.md, CHANGELOG.md, docs/*) so it can never drift from the real
docs. This file only wires MkDocs events to the helpers in ``sitegen/``:

* ``sitegen/sources.py`` - the single ``PAGES`` list of published documents and
  how each is adapted (written to ``_generated/``, pulled into the thin wrapper
  pages under ``docs/`` via ``--8<--`` includes);
* ``sitegen/analytics.py`` - the usage-analytics page;
* ``sitegen/sitemap.py`` - git-derived ``<lastmod>`` and video data for
  ``overrides/sitemap.xml``;
* ``sitegen/llm.py`` - llms.txt, llms-full.txt, Markdown mirrors, tools.json
  and FAQ structured data.

MkDocs puts this file's folder on ``sys.path`` while loading it, which is what
makes the ``sitegen`` import below resolve.
"""

from __future__ import annotations

from pathlib import Path

from sitegen import llm, sitemap
from sitegen.analytics import render_usage_analytics
from sitegen.sources import PAGES, adapt, read, write


def on_pre_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    for page in PAGES:
        if page.header == "rendered":
            content = render_usage_analytics()
        else:
            content = adapt(page, read(page.source))
        write(page, content, config["repo_url"])


def on_nav(nav, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    llm.set_nav(nav.items)
    return nav


def on_page_markdown(markdown, page, config, **kwargs):  # noqa: D401 - MkDocs hook
    """Capture each page's full Markdown for the LLM-facing outputs."""
    llm.capture_page(markdown, page, config["site_url"])
    return markdown


def on_env(env, config, files, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Expose sitemap data to overrides/sitemap.xml."""
    env.globals["page_lastmod"] = sitemap.page_lastmod(files)
    env.globals["video"] = {**sitemap.VIDEO, "page_url": config["site_url"]}
    return env


def on_post_build(config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Write the LLM-facing outputs that MkDocs itself has no notion of."""
    site_dir = Path(config["site_dir"])
    llm.write_llm_outputs(site_dir, config["site_url"])
    llm.write_tools_json(site_dir, config["site_url"], config["repo_url"])


def on_post_page(output, page, config, **kwargs):  # noqa: D401 - MkDocs hook signature
    """Give Material's search dialog an accessible name.

    A role="dialog" with no name is a WCAG 4.1.2 failure. Unlike the logo and
    progress-bar fixes - declarative partials under ``overrides/`` - this one
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
