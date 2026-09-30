"""Verify the fixed P0-11 narrowed-scope inputs on this machine."""

from hashlib import sha256
import json
from pathlib import Path
from subprocess import run
from tempfile import TemporaryDirectory


ROOT = Path(__file__).resolve().parent
FILES = (
    "naca-rm-l50b01-1950-real-scan.pdf",
    "deardorff-2020-article.pdf",
    "deardorff-2020-assistive-text.txt",
)


def digest(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def pages(path: Path) -> int:
    result = run(["pdfinfo", str(path)], capture_output=True, text=True, check=True)
    return int(
        next(
            line.split(":", 1)[1].strip()
            for line in result.stdout.splitlines()
            if line.startswith("Pages:")
        )
    )


def main() -> None:
    expected = {
        name: checksum
        for line in (ROOT / "SHA256SUMS").read_text().splitlines()
        if line.strip()
        for checksum, name in [line.split()]
    }
    for name in FILES:
        assert digest(ROOT / name) == expected[name], f"changed material: {name}"
    assert pages(ROOT / FILES[0]) == 30
    assert pages(ROOT / FILES[1]) == 11
    assert digest(ROOT / FILES[1]) == digest(ROOT.parent / "PPT-P0-01" / FILES[1])
    audit = json.loads((ROOT / "naca-rm-l50b01-scan-audit.json").read_text())
    assert audit["normalizedPdfSha256"] == digest(ROOT / FILES[0])
    assert audit["ocrReviewStatus"] == "unverified"
    assert len(audit["pageCoverage"]) == 30
    assert audit["pageCoverage"][1] == {
        "page": 2,
        "extractedCharacters": 0,
        "status": "no_text",
    }

    with TemporaryDirectory() as temporary:
        extracted = Path(temporary) / "article.txt"
        run(["pdftotext", "-layout", str(ROOT / FILES[1]), str(extracted)], check=True)
        assert extracted.read_bytes() == (ROOT / FILES[2]).read_bytes(), (
            "helper text is not the fixed PDF extraction"
        )

    scan_page = run(
        ["pdftotext", "-f", "2", "-l", "2", str(ROOT / FILES[0]), "-"],
        capture_output=True,
        check=True,
    )
    assert not scan_page.stdout.strip(), "the known blank scan page changed"
    print(
        "P0-11 narrowed-scope candidate inputs verified: "
        "30-page scan, 11-page article, exact same-source helper"
    )


if __name__ == "__main__":
    main()
