---
"excelmcp": patch
---

**Fixed broken documentation links on NuGet.org** — the MCP Server and CLI package pages linked to the installation guides, feature reference, and privacy policy using relative paths. NuGet.org resolves those against the package itself rather than the repository, so every one of them returned a 404 for anyone reading the package page. They now point at absolute repository URLs and work from NuGet.org, GitHub, and inside the shipped skill packages alike.

The documentation site is unaffected: those links still resolve to the site's own pages, and a new audit check fails the build if a published page ever starts sending readers to GitHub for content the site hosts itself.
