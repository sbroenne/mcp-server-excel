"""Check that the sample download is shipped unchanged with the website."""

from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET

from jinja2 import Environment, FileSystemLoader
from mkdocs.config import load_config
from mkdocs.structure.files import Files

import hooks
from sitegen import sitemap, sources

GH_PAGES = Path(__file__).resolve().parent.parent


class SampleDownloadTests(unittest.TestCase):
    def test_sitemap_describes_both_videos_on_their_own_pages(self):
        pages = [
            SimpleNamespace(
                src_uri=source,
                page=SimpleNamespace(is_link=False, canonical_url=url, abs_url=url),
            )
            for source, url in [
                ("index.md", "https://excelmcpserver.dev/"),
                ("samples/world-in-motion.md", "https://excelmcpserver.dev/samples/world-in-motion/"),
            ]
        ]
        env = Environment(loader=FileSystemLoader(GH_PAGES / "overrides"))
        with patch.object(sitemap, "page_lastmod", return_value={}):
            hooks.on_env(env, {"site_url": "https://excelmcpserver.dev/"}, Files([]))
        root = ET.fromstring(env.get_template("sitemap.xml").render(pages=pages))
        ns = {
            "s": "http://www.sitemaps.org/schemas/sitemap/0.9",
            "v": "http://www.google.com/schemas/sitemap-video/1.1",
        }
        entries = {entry.findtext("s:loc", namespaces=ns): entry for entry in root}
        expected = {
            "https://excelmcpserver.dev/": ("wbw3-hPcE2o", "121"),
            "https://excelmcpserver.dev/samples/world-in-motion/": ("47HJPZbcta4", "154"),
        }
        for url, (video_id, duration) in expected.items():
            with self.subTest(url=url):
                videos = entries[url].findall("v:video", ns)
                self.assertEqual(len(videos), 1)
                video = videos[0]
                self.assertEqual(
                    video.findtext("v:player_loc", namespaces=ns),
                    f"https://www.youtube.com/embed/{video_id}",
                )
                self.assertEqual(video.findtext("v:duration", namespaces=ns), duration)
                self.assertTrue(video.findtext("v:thumbnail_loc", namespaces=ns))
                self.assertTrue(video.findtext("v:title", namespaces=ns))
                self.assertTrue(video.findtext("v:description", namespaces=ns))

    def test_workbook_link_points_to_built_download(self):
        text = sources.rewrite_links(
            "[Download](world-in-motion.xlsx)",
            "samples/world-bank-dashboard/README.md",
            "https://github.com/sbroenne/mcp-server-excel",
        )
        self.assertEqual(text, "[Download](/downloads/world-in-motion.xlsx)")

    def test_sample_assets_are_included_in_deploy_source_coverage(self):
        self.assertTrue(set(sources.SAMPLE_ASSETS.values()).issubset(sources.SOURCE_FILES))

    def test_sitemap_uses_the_configured_site_root(self):
        env = Environment()
        with patch.object(sitemap, "page_lastmod", return_value={}):
            hooks.on_env(env, {"site_url": "https://example.test/docs/"}, Files([]))
        self.assertEqual(
            [video["page_url"] for video in env.globals["videos"]],
            ["https://example.test/docs/", "https://example.test/docs/samples/world-in-motion/"],
        )

    def test_assets_are_copied_without_changing_the_source(self):
        with tempfile.TemporaryDirectory() as directory:
            config = load_config(
                config_file=str(GH_PAGES / "mkdocs.yml"),
                site_dir=directory,
            )
            with patch.object(config.plugins, "_current_plugin", "hooks.py", create=True):
                files = hooks.on_files(Files([]), config)
            self.assertIn("downloads/world-in-motion.xlsx", sources.SAMPLE_ASSETS)
            for destination, source in sources.SAMPLE_ASSETS.items():
                with self.subTest(destination=destination):
                    file = files.get_file_from_path(destination)
                    self.assertIsNotNone(file)
                    file.copy_file()
                    self.assertEqual(
                        (Path(directory) / destination).read_bytes(),
                        (sources.REPO_ROOT / source).read_bytes(),
                    )

    def test_missing_sample_stops_the_build(self):
        with tempfile.TemporaryDirectory() as directory:
            config = load_config(
                config_file=str(GH_PAGES / "mkdocs.yml"),
                site_dir=directory,
            )
            with patch.dict(sources.SAMPLE_ASSETS, {"downloads/missing.xlsx": "missing.xlsx"}, clear=True):
                with self.assertRaises(FileNotFoundError):
                    hooks.on_files(Files([]), config)


if __name__ == "__main__":
    unittest.main()
