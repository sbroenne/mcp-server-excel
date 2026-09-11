"""Crop genuine Excel window captures without altering their contents."""
from pathlib import Path
from PIL import Image

root = Path(__file__).resolve().parent
for name, box in (
    ("sales-data", (0, 270, 620, 730)),
    ("sales-query", (0, 270, 620, 590)),
    ("sales-report", (0, 268, 1245, 850)),
):
    with Image.open(root / "assets" / f"{name}.png") as image:
        image.crop(box).save(root / "assets" / f"{name}-detail.png")
