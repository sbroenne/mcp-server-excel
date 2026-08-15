# Docs Site (MkDocs)

Source for [excelmcpserver.dev](https://excelmcpserver.dev/), built with MkDocs Material.
Most pages under `docs/` are thin wrappers that include canonical content from elsewhere in
the repo (root `README.md`, `FEATURES.md`, `docs/features/`, package READMEs,
`CHANGELOG.md`, etc.) so there is a single source of truth for documentation content.
The canonical feature reference is organized into intent-based pages under `docs/features/`;
`hooks.py` adapts those pages for the website without copying operation details.

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
.\.venv\Scripts\python.exe audit_site.py           # SEO / LLM-discoverability audit
.\.venv\Scripts\python.exe check_deploy_paths.py   # deploy paths: filter covers every mirrored source
```

Two further checks run on a schedule rather than per pull request:

| Workflow | When | What |
| --- | --- | --- |
| `link-check.yml` | Weekly | Runs lychee over the built site and files an issue on link rot. Not a PR check: external endpoints rate-limit and would make it flaky. |
| `star-history.yml` | Daily, plus PRs touching the scripts | Records and persists the star snapshot. Split out of the Pages build so that job no longer needs `contents: write` while installing pip packages. |

