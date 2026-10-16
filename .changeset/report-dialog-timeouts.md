---
"excelmcp": patch
---

When an Excel operation times out with a dialog open, include the dialog's title and readable text in the error instead of only suggesting a longer timeout. Existing timeout recovery still closes the failed session.
