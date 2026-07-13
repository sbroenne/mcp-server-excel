---
"excelmcp": patch
---

**More reliable Python in Excel results.** `pythoninexcel get-result` now detects when the Microsoft-hosted Python backend has finished computing by reading Excel's calculation state and the cell's `#BUSY!` placeholder directly, instead of guessing based on whether the value looked "stable" across repeated reads. The old heuristic could lock onto a stale placeholder and return the wrong value, which is why it needed retry loops to be dependable. A single `get-result` call now converges deterministically, and the default wait was raised from 15s to 30s to comfortably cover cold-start round-trips.
