# Docs Site (MkDocs)

Source for [excelmcpserver.dev](https://excelmcpserver.dev/), built with MkDocs Material.
Most pages under `docs/` are thin wrappers that include canonical content from elsewhere in
the repo (root `README.md`, `FEATURES.md`, `docs/features/`, package READMEs,
`CHANGELOG.md`, etc.) so there is a single source of truth for documentation content.
Capability summaries live in `docs/features/`; workflow decisions and recovery
guidance live in `docs/reference/`. Current command specifications come from
CLI help and MCP tool descriptions, not a second hand-maintained reference.
`hooks.py` adapts the shared pages for the website without duplicating their prose.
It only wires MkDocs build events; the work lives in the `sitegen/` package:

| Module | Job |
| --- | --- |
| `sitegen/sources.py` | `PAGES`, the one table of published documents, plus sample downloads, link rewriting and snippet writing |
| `sitegen/llm.py` | `llms.txt`, `llms-full.txt`, Markdown mirrors, `tools.json`, FAQ structured data |
| `sitegen/sitemap.py` | Git-based sitemap dates and per-page video metadata |
| `sitegen/analytics.py` | The usage analytics page |

`mkdocs serve` loads `sitegen/` once; restart it after editing that code.

The feature overview is authored only in root `FEATURES.md`. Its website
wrapper, `docs/features.md`, keeps the page metadata, title, and illustration,
then includes `_generated/features.md`. Edit the root file to change categories,
capability navigation, task links, or headline counts. The site audit rejects
duplicate overview prose in the wrapper and checks the published Markdown copy.

## Publishing canonical documentation

Write operational content once in its repository source. Pages under
`gh-pages/docs/` own presentation: metadata, navigation, images, and Material
components. Preserve substantive examples and caveats at their canonical
destination before shortening or moving a page.

To add or move a published document:

1. Update the canonical source and links pointing to it.
2. Add a `Page(...)` entry to `PAGES` in `sitegen/sources.py`. That one entry
   generates the snippet and turns repository links to it into local website
   links.
3. Create a thin wrapper under `gh-pages/docs/` with metadata, one H1, and the
   generated snippet.
4. Update `nav` in `mkdocs.yml` and the deploy workflow's source path filters.
5. Run the strict build and the checks described below.

Example wrapper:

```markdown
---
title: Page Title
description: What this page helps the reader do.
keywords: relevant, search terms
---

# Page Title

--8<-- "_generated/page-name.md"
```

The hook writes snippets to gitignored `_generated/`, outside `docs/` to avoid
preview rebuild loops; do not edit those
files. Use local site links for published documents and GitHub links for source
code, issues, or documents without a site page.

### Generated reader and agent outputs

| Output | Source and purpose |
|--------|--------------------|
| `llms.txt` | Navigation-ordered page index and descriptions |
| `llms-full.txt` | Full Markdown with snippet content resolved |
| Page `index.md` mirrors | Markdown alternatives to rendered HTML |
| `tools.json` | Feature-group counts and capability summaries derived from the feature pages |
| FAQ structured data | Troubleshooting question blocks |

These are generated, not separately maintained. `tools.json` derives its
capability summaries and operation total from the feature groups and operation counts in
`docs/features/`, while its tool total comes from `doc-counts.json` (repo
root) - the single generated include file every count consumer reads.
Each `featureGroups` entry contains `capabilities` with names and descriptions,
not an `operations` command inventory. Its `operationCount` is the number of
supported operations, not the number of capability summaries. Current command
specifications come from CLI help and MCP tool descriptions.
`llms.txt` reads its advertised summary from that same file. Contributors run
`scripts\check-doc-counts.ps1 -Update` and review `doc-counts.json` and managed
headline changes in the source PR. CI rejects stale counts before merge;
do not substitute manual file or folder counts.

The catalogue describes the full Windows surface, not Mac availability.
`tools.json` separately declares experimental beta Mac support, unsupported
feature families, and the generated per-action inventory link. The `llms.txt`
summary carries the same platform boundary; the site audit rejects omissions.

## Theme overrides

`overrides/` holds the templates that change MkDocs/Material output:

| File | Why |
| --- | --- |
| `sitemap.xml` | Adds a real `<lastmod>` (the git commit date behind each page, supplied by `sitegen/sitemap.py`) and video details for the homepage introduction and sample-page dashboard demo. The stock template stamps the *build* date on every URL, which told crawlers all 52 pages changed on every deploy. |
| `partials/logo.html` | Upstream renders `alt="logo"` with no dimensions - a WCAG 1.1.1 failure and an unsized image. |
| `partials/progress.html` | Upstream's `role="progressbar"` has no accessible name (WCAG 4.1.2). |

Material's search dialog needs the same treatment, but its partial is ~45 lines
of markup and feature flags, so forking it to add one attribute would pin a
large slice of Material internals. That one stays a string patch in
`hooks.py`. `audit_site.py` asserts that these overrides remain active and also
checks content-image alt text, nested breadcrumbs, metadata, links, and
machine-readable outputs. An upstream change that breaks them therefore fails
the build instead of silently regressing accessibility.

## Setup (one-time)

Windows:

```powershell
cd gh-pages
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

macOS:

```bash
cd gh-pages
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

The site build is Excel-independent. On Mac use `.venv/bin/python` in the
commands below instead of the Windows `.venv\Scripts\python.exe` path.

## ⚠️ Always use the venv Python

A global `mkdocs` on `PATH` may resolve to a different Python install with
incompatible dependencies. Always invoke MkDocs through the project's venv:

```powershell
cd gh-pages
.\.venv\Scripts\python.exe -m mkdocs serve   # live preview with auto-reload
.\.venv\Scripts\python.exe -m mkdocs build --strict --clean   # verify before commit
```

(Alternatively, activate the venv first with `.\.venv\Scripts\Activate.ps1`, then plain
`mkdocs serve`/`mkdocs build` will use the correct interpreter.)

## Checks

These run in the `Docs Site` CI job on every pull request, and can be run locally
(the first two after a build):

```powershell
cd gh-pages
.\.venv\Scripts\python.exe audit_site.py           # SEO / a11y / LLM-discoverability audit
.\.venv\Scripts\python.exe check_deploy_paths.py   # deploy paths: filter covers every mirrored source
.\.venv\Scripts\python.exe -m unittest discover -s tests   # sitegen, sample packaging and video sitemap
```

The World in Motion sample page mirrors `samples/world-bank-dashboard/README.md`.
`SAMPLE_ASSETS` in `sitegen/sources.py` adds the original workbook, attribution files, and
verified dashboard stills to the build without keeping another workbook in
`docs/`. The packaging check confirms those files are copied unchanged and
missing sources fail the build. The sample page is available under
`/samples/world-in-motion/` after deployment; preparing or building it locally
does not publish it or replace the homepage video.
The sitemap describes the new video under that sample-page URL, including its
title, description, thumbnail, player link and duration. The homepage keeps the
original introduction's video entry. The packaging test also renders the sitemap
template and checks that both videos remain associated with the correct pages.

The World Bank Briefing page mirrors `samples/world-bank-briefing/README.md`
the same way. `SAMPLE_ASSETS` publishes the agent's workbook and deck unchanged,
plus four slide images from `videos/agentic-world-bank-briefing/assets/slides/fan/`.
The homepage links both sample pages from its video cards.

Both workflows that build the site check out with `fetch-depth: 0`, because the
sitemap dates come from `git log`. On a shallow clone every page would claim the
tip commit's date; `audit_site.py` fails the build when it sees that.

Two further checks run on a schedule rather than per pull request:

| Workflow | When | What |
| --- | --- | --- |
| `link-check.yml` | Weekly | Runs lychee over the built site and files an issue on link rot. Not a PR check: external endpoints rate-limit and would make it flaky. |
| `star-history.yml` | Daily, plus PRs touching the scripts | Records and persists the star snapshot. Split out of the Pages build so that job no longer needs `contents: write` while installing pip packages. |
