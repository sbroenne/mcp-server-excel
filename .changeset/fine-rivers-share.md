---
"excelmcp": minor
---

**Documentation site overhaul** — focused feature pages instead of one long reference, five new task guides (refresh Power Query, automate PivotTables, query the Data Model with DAX, run VBA macros, and how COM automation compares to file-parser libraries), and the 24-file expert reference corpus that previously shipped only inside the agent skills is now published on the web.

The site also gained a machine-readable layer for AI assistants: `/llms.txt`, `/llms-full.txt`, a Markdown mirror of every page (append `index.md` to any URL), `/tools.json` describing all 31 tools and 326 operations, FAQ structured data, and an explicit AI-crawler policy in `robots.txt`.

Distribution metadata was corrected and locked down: the Claude Desktop bundle description and the CLI NuGet description advertised outdated tool and operation counts, and the NuGet package pages now link to the documentation site. `scripts/check-doc-counts.ps1` guards both, so those counts and links cannot silently rot again.
