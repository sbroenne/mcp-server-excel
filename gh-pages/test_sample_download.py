"""Check that the sample download is shipped unchanged with the website."""

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from mkdocs.config import load_config
from mkdocs.structure.files import Files

import hooks


class SampleDownloadTests(unittest.TestCase):
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
