#!/usr/bin/env python3
"""Verify the Pages deploy workflow rebuilds the site for every canonical source.

``sitegen/sources.py`` lists the canonical Markdown the site publishes, but
``.github/workflows/deploy-gh-pages.yml`` decides *when* to rebuild from a
hand-written ``paths:`` filter. Nothing kept the two in sync, so adding a new
mirrored source silently produced a stale website until the next nightly cron.

This script fails when a source that the build reads is not covered by the
workflow filter. Run it from the ``gh-pages`` directory::

    python check_deploy_paths.py
"""

from __future__ import annotations

import fnmatch
import sys
from pathlib import Path

import yaml

GH_PAGES = Path(__file__).resolve().parent
REPO_ROOT = GH_PAGES.parent
WORKFLOW = REPO_ROOT / ".github" / "workflows" / "deploy-gh-pages.yml"

def mirrored_sources() -> set[str]:
    sys.path.insert(0, str(GH_PAGES))
    from sitegen.sources import SOURCE_FILES  # noqa: PLC0415 - needs sys.path above

    return set(SOURCE_FILES)


def workflow_paths() -> list[str]:
    config = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
    # YAML 1.1 parses a bare `on:` key as the boolean True.
    triggers = config.get("on", config.get(True, {})) or {}
    return list((triggers.get("push") or {}).get("paths") or [])


def main() -> int:
    if not WORKFLOW.is_file():
        print(f"ERROR: {WORKFLOW} not found", file=sys.stderr)
        return 2

    patterns = workflow_paths()
    if not patterns:
        print("ERROR: the deploy workflow declares no push paths filter", file=sys.stderr)
        return 2

    missing: list[str] = []
    for source in sorted(mirrored_sources()):
        if not (REPO_ROOT / source).is_file():
            missing.append(f"{source}: listed in sitegen/sources.py but does not exist in the repo")
            continue
        if not any(fnmatch.fnmatch(source, pattern) for pattern in patterns):
            missing.append(
                f"{source}: mirrored into the site but no paths: entry in "
                f"{WORKFLOW.relative_to(REPO_ROOT).as_posix()} matches it"
            )

    if missing:
        print(
            "Deploy path check FAILED - editing these files would not rebuild "
            f"the site:\n"
        )
        for item in missing:
            print(f"  - {item}")
        return 1

    print(
        f"Deploy path check passed: {len(mirrored_sources())} mirrored sources, "
        f"all covered by {len(patterns)} paths entries."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
