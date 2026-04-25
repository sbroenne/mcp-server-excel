---
name: "plugin-overlay-bundling"
description: "Pattern for keeping published plugin templates stable while layering source-owned helper files and skill content."
---

## Context

Use this when plugin artifacts are built by copying validated templates from a published/distribution repo, but the source repo still needs to own release-coupled helper files or skill content.

## Pattern

1. Keep the validated published plugin structure as the base copy input.
2. Add a source-owned overlay directory (for example `.github/plugins/<plugin-name>/`) for files that should be authored in the source repo:
   - plugin README updates
   - install helpers
   - wrapper scripts
3. In the build script, apply the overlay after copying the template, then inject version metadata.
4. Prefer keeping published plugins lightweight; ship binaries through the main release artifacts unless the marketplace repo can safely host them.
5. Update the release workflow and source-side sync gate to watch both the build script and the overlay directory.
6. If the published marketplace repo is mid-migration, make the automation prefer the canonical marketplace manifest path/schema (`.github/plugin/marketplace.json`, `metadata.version`) but tolerate the legacy path/schema (`marketplace.json`, `version`) until the published repo is updated.

## Why

- Preserves the “copy validated templates” discipline without forcing the published repo to be the only authoring surface.
- Avoids pushing oversized release binaries into the published marketplace repo.
- Keeps plugin-specific helper scripts versioned in the source repo, where product/release changes are made.
- Lets source automation become more spec-compliant without forcing an all-at-once published-repo migration.

## Example

- `scripts/Build-Plugins.ps1` copies `../mcp-server-excel-plugins/plugins/excel-cli`, applies `.github/plugins/excel-cli/`, and refreshes the skill content without adding release binaries.
- `publish-plugins.yml` can publish the plugin bundle directly after `Build-Plugins.ps1` because the standalone CLI remains distributed through GitHub Releases and NuGet.
