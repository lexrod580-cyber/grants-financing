from pathlib import Path
from PIL import Image, ImageOps, ImageDraw

src = Path(__file__).resolve().parent / "rendered_word"
out = Path(__file__).resolve().parent / "contact_sheets"
out.mkdir(exist_ok=True)
files = sorted(src.glob("page-*.png"))
for batch_index in range(0, len(files), 8):
    batch = files[batch_index:batch_index + 8]
    thumbs = []
    for file in batch:
        im = Image.open(file).convert("RGB")
        im.thumbnail((330, 470))
        canvas = Image.new("RGB", (350, 510), "white")
        canvas.paste(im, ((350 - im.width) // 2, 24))
        draw = ImageDraw.Draw(canvas)
        draw.text((12, 4), file.stem, fill="black")
        canvas = ImageOps.expand(canvas, border=1, fill="#888888")
        thumbs.append(canvas)
    sheet = Image.new("RGB", (704, 2048), "#d0d0d0")
    for idx, thumb in enumerate(thumbs):
        x = (idx % 2) * 352
        y = (idx // 2) * 512
        sheet.paste(thumb, (x, y))
    sheet.save(out / f"sheet-{batch_index // 8 + 1}.png")
print(len(files))
