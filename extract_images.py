from __future__ import annotations

import io
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


source = Path("Nội dung bài báo cáo CĐTN.docx")
output = Path("tmp/source_images")
output.mkdir(parents=True, exist_ok=True)

paths: list[Path] = []
with zipfile.ZipFile(source) as archive:
    for name in sorted(n for n in archive.namelist() if n.startswith("word/media/")):
        data = archive.read(name)
        target = output / Path(name).name
        target.write_bytes(data)
        paths.append(target)

thumb_w, thumb_h = 260, 210
label_h = 28
cols = 4
rows = (len(paths) + cols - 1) // cols
sheet = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "white")
draw = ImageDraw.Draw(sheet)

for index, path in enumerate(paths):
    with Image.open(path) as image:
        image = image.convert("RGB")
        image.thumbnail((thumb_w - 12, thumb_h - 12))
        x = (index % cols) * thumb_w + (thumb_w - image.width) // 2
        y0 = (index // cols) * (thumb_h + label_h)
        y = y0 + (thumb_h - image.height) // 2
        sheet.paste(image, (x, y))
        draw.rectangle(
            [index % cols * thumb_w, y0, (index % cols + 1) * thumb_w - 1, y0 + thumb_h + label_h - 1],
            outline="#999999",
        )
        draw.text((index % cols * thumb_w + 6, y0 + thumb_h + 4), path.name, fill="black")

sheet.save(output / "contact_sheet.png")
print(output / "contact_sheet.png")
