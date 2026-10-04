# Docs Site (Zensical)

Source for [excelmcpserver.dev](https://excelmcpserver.dev/), built with [Zensical](https://zensical.org/).
Most pages under `docs/` are thin wrappers that include canonical content from elsewhere in
the repo (root `README.md`, `FEATURES.md`, `docs/features/`, package READMEs,
`CHANGELOG.md`, etc.) so there is a single source of truth for documentation content.
Capability summaries live in `docs/features/`; workflow decisions and recovery
guidance live in `docs/reference/`. Current command specifications come from
CLI help and MCP tool descriptions, not a second hand-maintained reference.
`generate.py` adapts the shared pages for the website without duplicating their prose.
It runs before every build; Zensical itself has no build hooks.

The feature overview is authored only in root `FEATURES.md`. Its website
wrapper, `docs/features.md`, keeps the page metadata, title, and illustration,
then includes `_generated/features.md`. Edit the root file to change categories,
capability navigation, task links, or headline counts. The site audit rejects
duplicate overview prose in the wrapper and checks the published Markdown copy.

## Publishing canonical documentation

Write operational content once in its repository source. Pages under
`gh-pages/docs/` own presentation: metadata, navigation, images, and theme
components. Preserve substantive examples and caveats at their canonical
destination before shortening or moving a page.

To add or move a published document:

1. Update the canonical source and links pointing to it.
2. Register it in the appropriate source map or `_write()` step in `generate.py`.
3. Add its repository path to `SITE_PAGE_MAP`, so repository-relative links
   become local website links.
4. Create a thin wrapper under `gh-pages/docs/` with metadata, one H1, and the
   generated snippet.
5. Update `nav` in `zensical.toml` (and `plugins.llmstxt.sections` if the page
   belongs in `llms.txt`) and the deploy workflow's source path filters.
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

`generate.py` writes snippets to gitignored `_generated/`, outside `docs/` to avoid
preview rebuild loops; do not edit those
files. Use local site links for published documents and GitHub links for source
code, issues, or documents without a site page.

### Generated reader and agent outputs

| Output | Produced by | Purpose |
|--------|-------------|---------|
| `llms.txt` | Zensical `llmstxt` plugin | Page index grouped by the sections in `zensical.toml` |
| `llms-full.txt` | Zensical `llmstxt` plugin | Full Markdown of every page, snippets resolved |
| Page `index.md` mirrors | Zensical `llmstxt` plugin | Markdown alternatives to rendered HTML |
| FAQ structured data | `generate.py` | `FAQPage` JSON-LD from the `###` questions in `faq.md` |
| Sitemap `<lastmod>` | `generate.py` | Git commit date behind each page |

These are generated, not separately maintained. `generate.py` writes the FAQ
and sitemap pieces as small template files under gitignored
`overrides/generated/`, which `overrides/main.html` and `overrides/sitemap.xml`
include. Headline tool and operation counts on the site live in
`docs/index.md` and `docs/faq.md`; contributors run
`scripts\check-doc-counts.ps1 -Update` and CI rejects stale counts before merge.

## Theme overrides

`overrides/` holds the templates that change the theme's output:

| File | Why |
| --- | --- |
| `sitemap.xml` | Adds a real `<lastmod>` (the git commit date behind each page, supplied by `generate.py`) and the home page's `<video:video>` block. The stock template stamps the *build* date on every URL, which told crawlers all 52 pages changed on every deploy. |
| `partials/logo.html` | Upstream renders `alt="logo"` with no dimensions - a WCAG 1.1.1 failure and an unsized image. |
| `partials/progress.html` | Upstream's `role="progressbar"` has no accessible name (WCAG 4.1.2). |

The theme already labels its search dialog. `audit_site.py` asserts that
these accessible names remain present and also
checks content-image alt text, nested breadcrumbs, metadata, links, and
machine-readable outputs. An upstream change that breaks them therefore fails
the build instead of silently regressing accessibility.

## Setup (one-time)

```powershell
cd gh-pages
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Build and preview

Always use the project's venv, and run `generate.py` first. Rerun it after
editing any canonical source outside `gh-pages/docs/`; the preview server does
not watch those files.

```powershell
cd gh-pages
.\.venv\Scripts\python.exe generate.py
.\.venv\Scripts\zensical.exe serve                  # live preview
.\.venv\Scripts\zensical.exe build --clean --strict # verify before commit
```

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
