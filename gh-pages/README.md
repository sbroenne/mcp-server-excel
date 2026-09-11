# Docs Site (MkDocs)

Source for [excelmcpserver.dev](https://excelmcpserver.dev/), built with MkDocs Material.
Most pages under `docs/` are thin wrappers that include canonical content from elsewhere in
the repo (root `README.md`, `FEATURES.md`, `docs/features/`, package READMEs,
`CHANGELOG.md`, etc.) so there is a single source of truth for documentation content.
The canonical feature reference is organized into intent-based pages under `docs/features/`;
`hooks.py` adapts those pages for the website without copying operation details.

The feature overview is authored only in root `FEATURES.md`. Its website
wrapper, `docs/features.md`, keeps the page metadata, title, and illustration,
then includes `_generated/features.md`. Edit the root file to change categories,
tool-selection guidance, task links, or headline counts. The site audit rejects
duplicate overview prose in the wrapper and checks the published Markdown copy.

## Publishing canonical documentation

Write operational content once in its repository source. Pages under
`gh-pages/docs/` own presentation: metadata, navigation, images, and Material
components. Preserve substantive examples and caveats at their canonical
destination before shortening or moving a page.

To add or move a published document:

1. Update the canonical source and links pointing to it.
2. Register it in the appropriate source map or `_write()` step in `hooks.py`.
3. Add its repository path to `SITE_PAGE_MAP`, so repository-relative links
   become local website links.
4. Create a thin wrapper under `gh-pages/docs/` with metadata, one H1, and the
   generated snippet.
5. Update `nav` in `mkdocs.yml` and the deploy workflow's source path filters.
6. Run the strict build and both checks described below.

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
| `tools.json` | Tool/operation catalogue derived from canonical feature references |
| FAQ structured data | Troubleshooting question blocks |

These are generated, not separately maintained. `tools.json` generation checks
feature operation totals against the headline count in `FEATURES.md`. The
repository's `scripts/check-doc-counts.ps1` checks advertised counts against
code-derived metadata; do not substitute manual file or folder counts.

## Theme overrides

`overrides/` holds the templates that change MkDocs/Material output:

| File | Why |
| --- | --- |
| `sitemap.xml` | Adds a real `<lastmod>` (the git commit date behind each page, supplied by `hooks.py`) and the home page's `<video:video>` block. The stock template stamps the *build* date on every URL, which told crawlers all 52 pages changed on every deploy. |
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

```powershell
cd gh-pages
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

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

Both run in the `Docs Site` CI job on every pull request, and can be run locally
after a build:

```powershell
cd gh-pages
.\.venv\Scripts\python.exe audit_site.py           # SEO / a11y / LLM-discoverability audit
.\.venv\Scripts\python.exe check_deploy_paths.py   # deploy paths: filter covers every mirrored source
```

Both workflows that build the site check out with `fetch-depth: 0`, because the
sitemap dates come from `git log`. On a shallow clone every page would claim the
tip commit's date; `audit_site.py` fails the build when it sees that.

Two further checks run on a schedule rather than per pull request:

| Workflow | When | What |
| --- | --- | --- |
| `link-check.yml` | Weekly | Runs lychee over the built site and files an issue on link rot. Not a PR check: external endpoints rate-limit and would make it flaky. |
| `star-history.yml` | Daily, plus PRs touching the scripts | Records and persists the star snapshot. Split out of the Pages build so that job no longer needs `contents: write` while installing pip packages. |
