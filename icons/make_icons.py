"""Generate the two program icons (run once; the .ico files are committed).

    python icons/make_icons.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).parent
SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def base(colour: str) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((8, 8, 248, 248), radius=48, fill=colour)
    return img, draw


def translator() -> Image.Image:
    img, draw = base("#1F5FAF")  # blue
    draw.rectangle((48, 52, 208, 90), fill="white")    # bar of the "T"
    draw.rectangle((108, 52, 148, 212), fill="white")  # stem of the "T"
    return img


def checker() -> Image.Image:
    img, draw = base("#1E8C3A")  # green
    draw.line([(56, 134), (108, 188), (204, 76)], fill="white", width=38, joint="curve")
    for x, y in [(56, 134), (204, 76)]:  # round line ends
        draw.ellipse((x - 19, y - 19, x + 19, y + 19), fill="white")
    return img


if __name__ == "__main__":
    for name, make in (("translator", translator), ("checker", checker)):
        img = make()
        img.save(HERE / f"{name}.ico", sizes=SIZES)
        img.save(HERE / f"{name}.png")
