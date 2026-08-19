from pathlib import Path

from PIL import Image, ImageDraw


def montage(prefix: str, output_name: str) -> None:
    paths = sorted(Path("work/export-qa").glob(f"{prefix}.xlsx-*.png"))
    images = [Image.open(path).convert("RGB") for path in paths]
    width = 1200
    panels = []

    for path, image in zip(paths, images):
        scale = min(1, width / image.width)
        resized = image.resize((int(image.width * scale), int(image.height * scale)))
        panel = Image.new("RGB", (width, resized.height + 34), "white")
        ImageDraw.Draw(panel).text((10, 8), path.stem, fill="black")
        panel.paste(resized, (0, 34))
        panels.append(panel)

    canvas = Image.new("RGB", (width, sum(panel.height for panel in panels)), "#dddddd")
    top = 0
    for panel in panels:
        canvas.paste(panel, (0, top))
        top += panel.height
    canvas.save(Path("work/export-qa") / output_name)


montage("validation", "validation-montage.png")
montage("reports", "reports-montage.png")
montage("hometax", "hometax-montage.png")
