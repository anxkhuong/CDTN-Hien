from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw


input_dir = Path(sys.argv[1])
output_path = Path(sys.argv[2])
paths = sorted(input_dir.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))

thumb_w, thumb_h = 170, 240
label_h = 24
cols = 6
rows = (len(paths) + cols - 1) // cols
sheet = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "#d9d9d9")
draw = ImageDraw.Draw(sheet)

for index, path in enumerate(paths):
    with Image.open(path) as page:
        page = page.convert("RGB")
        page.thumbnail((thumb_w - 10, thumb_h - 10))
        x0 = (index % cols) * thumb_w
        y0 = (index // cols) * (thumb_h + label_h)
        x = x0 + (thumb_w - page.width) // 2
        y = y0 + (thumb_h - page.height) // 2
        sheet.paste(page, (x, y))
        draw.rectangle([x, y, x + page.width - 1, y + page.height - 1], outline="#777777")
        draw.text((x0 + 6, y0 + thumb_h + 4), f"Trang {index + 1}", fill="black")

output_path.parent.mkdir(parents=True, exist_ok=True)
sheet.save(output_path)
print(output_path)
