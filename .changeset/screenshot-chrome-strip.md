---
"excelmcp": patch
---

**Screenshots no longer include a strip of Excel window chrome** (#777): captured images picked up a
few pixels of the scroll bar and sheet tab strip along the bottom of every tile, which showed up as a
grey band at each seam of a stitched screenshot of a tall or wide range. The capture now measures the
actual worksheet grid area instead of Excel's reported workspace size, and sizes tiles so they line
up seamlessly.
