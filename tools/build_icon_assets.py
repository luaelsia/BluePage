from pathlib import Path
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
MASTER = ROOT / "resources" / "BluePage.png"
ICON = ROOT / "resources" / "BluePage.ico"
OUT = ROOT / "resources" / "icons"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    master = Image.open(MASTER).convert("RGBA")

    sizes = [16, 20, 24, 32, 48, 64, 128, 256]
    for size in sizes:
        master.resize((size, size), Image.Resampling.LANCZOS).save(OUT / f"BluePage-{size}.png")

    master.save(ICON, format="ICO", sizes=[(size, size) for size in sizes])


if __name__ == "__main__":
    main()
