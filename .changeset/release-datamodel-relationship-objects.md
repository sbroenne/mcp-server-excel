---
"excelmcp": patch
---

Creating, updating, and deleting Data Model relationships now releases temporary Excel objects even when inspection fails. Threaded-comment operations also release the selected cell if validation fails. These fixes reduce the risk of Excel remaining open after a workbook is closed.
