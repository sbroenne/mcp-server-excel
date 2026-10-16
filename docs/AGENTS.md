# Documentation sources

- `FEATURES.md` is navigation; capability summaries live in `docs/features/`.
  Workflow guidance lives in `docs/reference/`; current command specifications
  come from CLI help and MCP tool descriptions.
  Permanent guides belong in `docs/`, decisions in `docs/ADR-*.md` (follow
  [maintaining decisions](DECISIONS.md#maintaining-decisions) and keep the
  index aligned), and proposed
  requirements in GitHub issues. Core contracts define operation behavior; do
  not maintain parallel specification copies. Agent guidance sources follow
  [product guidance](../skills/AGENTS.md).
- EXL030 validates advertised counts and distribution website links during MCP
  compilation using documentation inputs and the compiled Core manifest.
  Update affected `doc-counts.json` values, headlines, and canonical feature
  headings using the exact expected values reported by the diagnostic, review
  them in the code PR, and rebuild. There is no separate count updater.
  Website and release automation consume the reviewed include; the website's
  capability catalogue also uses canonical feature sections. Do not count files
  or folders, or add a post-merge rewrite of `main`.
- Website pages are thin wrappers over canonical repository docs.
  `gh-pages/hooks.py` and `gh-pages/sitegen/` write gitignored `_generated`
  snippets and machine-readable outputs. Do not hand-copy or separately maintain
  that content.
- Adding/moving a published source requires its `PAGES` entry in
  `gh-pages/sitegen/sources.py`, wrapper snippet, MkDocs nav, and deploy path
  filter to agree.
  Use local website links for published targets.

Authoring procedures: `docs/CONTRIBUTING.md` and `gh-pages/README.md`.

Website checks, from `gh-pages`:

```powershell
python -m mkdocs build --strict --clean
python audit_site.py
python check_deploy_paths.py
```

## Instruction maintenance

- Keep only repository-specific constraints, non-obvious pitfalls, required
  checks, and source pointers. Omit generic coding advice and inventories.
- Give each rule one authoritative home: common rules in root `AGENTS.md`,
  area rules beside their owners. Maintain the root task map for cross-area
  work and clients that do not automatically discover descendant instructions.
- Target current clients' native `AGENTS.md` support, without per-client copies
  or compatibility wrappers. Document discovery conditions in
  [agent development](agents/development.md); links are not automatic imports.
- Keep review checks early in root instructions. Include owning generators and
  templates in task scopes; distinguish review from implementation requirements.
- Verify rules against source, checks, and vendor documentation. Remove a rule
  only when an implemented compiler check covers it; retain semantic ownership,
  runtime behavior, and validation duties. A suppression reason is not proof.
- ADRs explain decisions; instructions constrain work; developer docs hold
  procedures. Update inbound links when moving any of them.
