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
        env = Environment(loader=FileSystemLoader(Path(__file__).parent / "overrides"))
        with patch.object(hooks, "_page_lastmod", return_value={}):
            hooks.on_env(env, None, Files([]))
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
        text = hooks._rewrite_links(
            "[Download](world-in-motion.xlsx)",
            "samples/world-bank-dashboard/README.md",
        )
        self.assertEqual(text, "[Download](/downloads/world-in-motion.xlsx)")

    def test_assets_are_copied_without_changing_the_source(self):
        with tempfile.TemporaryDirectory() as directory:
            config = load_config(
                config_file=str(Path(__file__).with_name("mkdocs.yml")),
                site_dir=directory,
            )
            with patch.object(config.plugins, "_current_plugin", "hooks.py", create=True):
                files = hooks.on_files(Files([]), config)
            self.assertIn("downloads/world-in-motion.xlsx", hooks.SAMPLE_ASSETS)
            for destination, source in hooks.SAMPLE_ASSETS.items():
                with self.subTest(destination=destination):
                    file = files.get_file_from_path(destination)
                    self.assertIsNotNone(file)
                    file.copy_file()
                    self.assertEqual(
                        (Path(directory) / destination).read_bytes(),
                        (hooks.REPO_ROOT / source).read_bytes(),
                    )

    def test_missing_sample_stops_the_build(self):
        with tempfile.TemporaryDirectory() as directory:
            config = load_config(
                config_file=str(Path(__file__).with_name("mkdocs.yml")),
                site_dir=directory,
            )
            with patch.dict(hooks.SAMPLE_ASSETS, {"downloads/missing.xlsx": "missing.xlsx"}, clear=True):
                with self.assertRaises(FileNotFoundError):
                    hooks.on_files(Files([]), config)


if __name__ == "__main__":
    unittest.main()
