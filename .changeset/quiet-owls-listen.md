---
"excelmcp": patch
---

**Accurate documentation-site dates and accessibility fixes**: the sitemap now
reports each page's real last-changed date instead of the date the site was
built, so search engines no longer see all 52 pages change on every deploy. The
site logo, loading indicator and search box also gained the accessible names and
image dimensions they were missing, and the build now fails if any of those
regress.
