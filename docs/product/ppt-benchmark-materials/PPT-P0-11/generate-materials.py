"""Rebuild the reproducible image-only PDF from the licensed source article."""

from pathlib import Path
from subprocess import run
from tempfile import TemporaryDirectory

from PIL import Image, JpegImagePlugin  # noqa: F401 - registers PDF's JPEG encoder


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "PPT-P0-01" / "deardorff-2020-article.pdf"
OUTPUT = ROOT / "deardorff-2020-image-only.pdf"
TEXT = ROOT / "deardorff-2020-assistive-text.txt"


def main() -> None:
    with TemporaryDirectory() as temporary:
        prefix = Path(temporary) / "page"
        run(["pdftoppm", "-png", "-r", "120", str(SOURCE), str(prefix)], check=True)
        paths = sorted(prefix.parent.glob("page-*.png"))
        if len(paths) != 11:
            raise RuntimeError(f"expected 11 rendered pages, got {len(paths)}")
        pages = [Image.open(path).convert("RGB") for path in paths]
        pages[0].save(OUTPUT, save_all=True, append_images=pages[1:], resolution=120.0)
        for page in pages:
            page.close()
    run(["pdftotext", "-layout", str(SOURCE), str(TEXT)], check=True)


if __name__ == "__main__":
    main()
