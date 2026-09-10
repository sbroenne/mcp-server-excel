---
applyTo: "README.md,**/README.md,**/index.md,FEATURES.md,CHANGELOG.md,SECURITY.md,PRIVACY.md,docs/**/*.md,specs/**/*.md,skills/**/*.md,gh-pages/**"
excludeAgent: "code-review"
---

# Documentation sources

- `FEATURES.md` is navigation; operation references live in `docs/features/`.
  Permanent guides belong in `docs/`, decisions in `docs/ADR-*.md`, requirements
  in `specs/`. Agent guidance sources follow `mcp-llm-guidance.instructions.md`.
- Advertised counts come from `scripts\check-doc-counts.ps1`, not counts of tool
  files or CLI folders. Internal CLI diagnostics are excluded from the shared
  advertised surface.
- Website pages are thin wrappers over canonical repository docs.
  `gh-pages/hooks.py` writes gitignored `_generated` snippets and machine-readable
  outputs. Do not hand-copy or separately maintain that content.
- Adding/moving a published source requires its hook source map/write step,
  `SITE_PAGE_MAP`, wrapper snippet, MkDocs nav, and deploy path filter to agree.
  Use local website links for published targets.

Authoring procedures: `docs/CONTRIBUTING.md` and `gh-pages/README.md`.

Website checks, from `gh-pages`:

```powershell
python -m mkdocs build --strict --clean
python audit_site.py
python check_deploy_paths.py
```
