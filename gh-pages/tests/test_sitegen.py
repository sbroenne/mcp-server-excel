"""Unit tests for the site build helpers. Run from ``gh-pages/``::

    python -m unittest discover -s tests
"""

from __future__ import annotations

import json
import re
import unittest

from sitegen.llm import faq_jsonld, resolve_snippets
from sitegen.sources import (
    DOCS_DIR,
    PAGES,
    REPO_ROOT,
    SITE_PAGE_MAP,
    add_stable_feature_anchors,
    adapt,
    rewrite_links,
    strip_header,
)

REPO_URL = "https://github.com/sbroenne/mcp-server-excel"


class RewriteLinksTests(unittest.TestCase):
    def rewrite(self, link: str, source: str = "docs/guides/RUN-VBA-MACROS.md") -> str:
        return rewrite_links(f"[x]({link})", source, REPO_URL)

    def test_relative_link_to_published_page_becomes_site_link(self):
        self.assertEqual(self.rewrite("../INSTALLATION.md#setup"), "[x](/installation/#setup)")

    def test_relative_link_to_unpublished_file_points_at_github(self):
        self.assertEqual(
            self.rewrite("../../src/ExcelMcp.Core/Foo.cs"),
            f"[x]({REPO_URL}/blob/main/src/ExcelMcp.Core/Foo.cs)",
        )

    def test_relative_directory_link_uses_tree_url(self):
        self.assertEqual(self.rewrite("../../scripts/"), f"[x]({REPO_URL}/tree/main/scripts)")

    def test_absolute_github_link_to_published_page_comes_back_to_site(self):
        self.assertEqual(
            self.rewrite(f"{REPO_URL}/blob/main/CHANGELOG.md#v1", "src/ExcelMcp.CLI/README.md"),
            "[x](/changelog/#v1)",
        )

    def test_absolute_github_link_to_unpublished_file_is_unchanged(self):
        link = f"{REPO_URL}/blob/main/global.json"
        self.assertEqual(self.rewrite(link), f"[x]({link})")

    def test_external_anchor_site_and_outside_repo_links_are_unchanged(self):
        for link in ("https://example.com", "#here", "/faq/", "mailto:a@b.c", "../../../x.md"):
            self.assertEqual(self.rewrite(link), f"[x]({link})")

    def test_images_are_not_rewritten(self):
        text = "![alt](../img.png)"
        self.assertEqual(rewrite_links(text, "docs/guides/README.md", REPO_URL), text)


class StripHeaderTests(unittest.TestCase):
    def test_drops_title_block_up_to_blank_line_and_demotes(self):
        source = "# Title\nintro line\n\nBody\n# Later\n"
        self.assertEqual(
            strip_header(source, end_on_blank=True, demote_h1=True), "Body\n## Later\n"
        )

    def test_drops_title_block_up_to_rule(self):
        source = "# Title\n\nSummary paragraph\n---\nBody\n"
        self.assertEqual(strip_header(source, end_on_hr=True), "Body\n")

    def test_drop_prefixes_only_apply_inside_header(self):
        source = "# Title\n[![badge](b)](l)\n\n[![kept](b)](l)\n"
        self.assertEqual(
            strip_header(source, drop_prefixes=("[![",), end_on_blank=True),
            "[![kept](b)](l)\n",
        )

    def test_keep_header_publishes_verbatim(self):
        page = next(p for p in PAGES if p.header == "keep")
        self.assertEqual(adapt(page, "\n# Security\n\nText\n\n"), "# Security\n\nText\n")


class FeatureAnchorTests(unittest.TestCase):
    def test_anchor_ignores_operation_count(self):
        self.assertEqual(
            add_stable_feature_anchors("## Power Query & M (12 operations)"),
            "## Power Query & M (12 operations) { #power-query-m }",
        )

    def test_other_headings_untouched(self):
        self.assertEqual(add_stable_feature_anchors("## Overview"), "## Overview")


class FaqJsonLdTests(unittest.TestCase):
    def test_question_headings_become_faq_entities(self):
        markdown = "\n".join(
            f"### Question {n}?\n\nAnswer **{n}** with [link](/x/).\n" for n in range(3)
        )
        data = json.loads(faq_jsonld(markdown))
        self.assertEqual(data["@type"], "FAQPage")
        self.assertEqual(len(data["mainEntity"]), 3)
        self.assertEqual(data["mainEntity"][0]["acceptedAnswer"]["text"], "Answer 0 with link.")

    def test_admonition_questions_drop_code_and_tables(self):
        block = '??? question "Q{n}?"\n    Text {n}.\n    ```\n    code\n    ```\n    | a | b |\n'
        data = json.loads(faq_jsonld("\n".join(block.format(n=n) for n in range(3))))
        self.assertEqual([e["acceptedAnswer"]["text"] for e in data["mainEntity"]],
                         ["Text 0.", "Text 1.", "Text 2."])

    def test_fewer_than_three_questions_is_not_an_faq(self):
        self.assertEqual(faq_jsonld("### One?\n\nYes.\n\n### Two?\n\nYes.\n"), "")


class PageTableTests(unittest.TestCase):
    def test_every_source_exists(self):
        missing = [p.source for p in PAGES if not (REPO_ROOT / p.source).is_file()]
        self.assertEqual(missing, [])

    def test_outputs_sources_and_urls_are_unique(self):
        for field in ("output", "source", "url"):
            values = [getattr(p, field) for p in PAGES]
            self.assertEqual(len(values), len(set(values)), field)

    def test_every_wrapper_include_has_a_page(self):
        outputs = {p.output for p in PAGES}
        include = re.compile(r'--8<--\s*"_generated/([^"]+)"')
        used = {
            name
            for path in DOCS_DIR.rglob("*.md")
            for name in include.findall(path.read_text(encoding="utf-8"))
        }
        self.assertEqual(used - outputs, set())
        self.assertEqual(outputs - used, set())

    def test_site_page_map_matches_pages(self):
        self.assertEqual(SITE_PAGE_MAP, {p.source: p.url for p in PAGES})


class SnippetTests(unittest.TestCase):
    def test_missing_snippet_resolves_to_empty(self):
        with self.assertLogs("mkdocs.hooks.generate", level="WARNING"):
            self.assertEqual(resolve_snippets('--8<-- "_generated/does-not-exist.md"'), "")


if __name__ == "__main__":
    unittest.main()
