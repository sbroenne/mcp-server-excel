---
"excelmcp": patch
---

Correct the text, web, and XML connection type names. Reading or changing a legacy text or web import's refresh settings now uses its worksheet query and retains its source and import settings.

Fix moving an existing text or web import to another worksheet. Keep its extraction, formatting and refresh settings, its connection name and description after saving and reopening, and the original source cells. Loading again at the existing A1 refreshes that import without shifting nearby cells. If refreshing a copied import fails, report the original and copied imports and connection so the remaining changes can be inspected before retrying.
